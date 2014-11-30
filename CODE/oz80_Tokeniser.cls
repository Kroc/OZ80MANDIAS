VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_TokenStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS:: oz80_TokenStream

'A Token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. This class handles _
 parsing the source code text into tokens; the assembler only works with tokens
 
'This class only represents one source code file. The assembler handles the specifics _
 of including additional source code files and spawns more TokenStream instances

'/// PRIVATE DEFS /////////////////////////////////////////////////////////////////////

'Used for time measurement
Private Stopwatch As PerformanceCounter

'The width of a tab stop (this affects the column number given)
Private Const TAB_WIDTH As Long = 8

'--------------------------------------------------------------------------------------

'For diaplying a disassembly
Private Const FORMAT_TOKEN          As String = "<tkn> $$ ($$,$$) $$ [$$] $$"
Private Const FORMAT_TOKEN_CHUNK    As String = "   CHUNK"
Private Const FORMAT_TOKEN_KEYWORD  As String = " KEYWORD"
Private Const FORMAT_TOKEN_LABEL    As String = "   LABEL"
Private Const FORMAT_TOKEN_NUMBER   As String = "  NUMBER"
Private Const FORMAT_TOKEN_OPERATOR As String = "OPERATOR"
Private Const FORMAT_TOKEN_PREFIX   As String = "  PREFIX"
Private Const FORMAT_TOKEN_PROPERTY As String = "PROPERTY"
Private Const FORMAT_TOKEN_QUOTE    As String = "   QUOTE"
Private Const FORMAT_TOKEN_RAM      As String = "     RAM"
Private Const FORMAT_TOKEN_REGISTER As String = "REGISTER"
Private Const FORMAT_TOKEN_SECTION  As String = " SECTION"
Private Const FORMAT_TOKEN_Z80      As String = "     Z80"
Private Const FORMAT_TOKEN_MEMZ80   As String = "  MEMZ80"

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'The error number!
Private My_Error As OZ80_ERROR

'Remember where the original source file is as the assembler will need to change to _
 its directory when handling relative references with includes
'TODO: Make a real property (should not be write-enabled)
Public SourceFile As String

'When parsing source text, we can save having to pass these through function calls _
 hundreds of times by making them global to this class
Private LineBegin As Long, ColBegin As Long

'We need to be able to tell the difference between a property being used as a _
 parameter on a Z80 instruction, and a label-property definition (within a chunk) _
 which occurs on a new line. We keep track of the previous line number used and _
 look for a property occuring on a new line
Private PrevLine As Long

'We will capture the word as we read it here. The bluString is both fast and _
 allows us to access the string as bytes or text
Private Word As bluString

'--------------------------------------------------------------------------------------

'Use a 100'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 100000

Private Tokens() As oz80Token           'The token stream will live here
Private TokenLimit As Long              'The length of the buffer used for the above
Private TokenCount As Long              'The current amount of space used in the array
Private TokenIndex As Long              'The current place within the tokenStream

Private Type oz80Token
    Kind As Byte                        '=OZ80_TOKEN, but use 1-byte instead of 4
    Value As Double                     'Associated value, e.g. for number tokens
    Line As Long                        'Line number in the original source text
    Col As Long                         'Column number in the original source text
End Type

'Caches of the current token to make multiple calls to this class a little quicker, _
 it also helps with a lot of internal referencing, mainly with logging
Private TokenKind As OZ80_TOKEN
Private TokenValue As Double
Private TokenLine As Long
Private TokenCol As Long

'We track the opening and closing of brackets to validate nesting of brackets _
 and to allow the assembler to skip over blocks
Private Type TokenNesting
    Token As OZ80_TOKEN                 'Which type of bracket `{`, `[`, `(`
    Index As Long                       'The token index of the opening bracket
End Type

'Look up tables to determine what category a Token belongs to; _
 these are exposed publicly through properties so other classes can do quick checks
Private My_IsInstruction(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister8Bit(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister16Bit(0 To [_TOKEN_LAST]) As Boolean
Private My_IsFlag(0 To [_TOKEN_LAST]) As Boolean
Private My_IsKeyword(0 To [_TOKEN_LAST]) As Boolean
Private My_IsOperator(0 To [_TOKEN_LAST]) As Boolean
Private My_IsExpression(0 To [_TOKEN_LAST]) As Boolean
Private My_IsParameter(0 To [_TOKEN_LAST]) As Boolean

'We could just unfold strings into a series of number tokens, but for quick access _
 to filenames with INCLUDE directives, we store strings in their own database
'NOTE: This is public so that the assembler can access the right string for a token _
 in order to normalise name items, e.g. `::section:label.property`
Public Strings As Scripting.Dictionary

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message( _
    ByRef LogLevel As OZ80_LOG, ByRef Text As String _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Set Stopwatch = New PerformanceCounter
    
    'Initialise the string database. _
     Since it's public, we want to ensure it's instantiated from the off
    Set Strings = New Scripting.Dictionary
    
    Dim ii As OZ80_TOKEN
    For ii = 0 To [_TOKEN_LAST]
        'Check what token category it belongs to
        Let My_IsInstruction(ii) = (ii > [_TOKEN_INSTRUCTIONS_BEGIN]) And _
                                   (ii < [_TOKEN_INSTRUCTIONS_END])
        Let My_IsOperator(ii) = (ii > [_TOKEN_OPERATORS_BEGIN]) And _
                                (ii < [_TOKEN_OPERATORS_END])
        Let My_IsKeyword(ii) = (ii > [_TOKEN_KEYWORDS_BEGIN]) And _
                               (ii < [_TOKEN_KEYWORDS_END])
        
        'Registers and flags are more complex as there's some overlap
        If (ii = TOKEN_Z80_C) Then
            'At the token level, we can't differentiate Register C and Flag C; _
             we need context for that and we can't gaurantee the TokenStream is _
             valid until we begin assembling, therefore we treat Register C and _
             Flag C as the same and the assembler will sort out the difference
            Let My_IsRegister8Bit(ii) = True
            Let My_IsParameter(ii) = True
            Let My_IsRegister(ii) = True
            Let My_IsFlag(ii) = True
        
        '8-Bit registers:
        ElseIf (ii = TOKEN_Z80_A) Or (ii = TOKEN_Z80_B) Or (ii = TOKEN_Z80_C) Or _
               (ii = TOKEN_Z80_D) Or (ii = TOKEN_Z80_E) Or (ii = TOKEN_Z80_H) Or _
               (ii = TOKEN_Z80_L) Or (ii = TOKEN_Z80_I) Or (ii = TOKEN_Z80_R) Or _
               (ii = TOKEN_Z80_IXL) Or (ii = TOKEN_Z80_IXH) Or _
               (ii = TOKEN_Z80_IYL) Or (ii = TOKEN_Z80_IYH) _
        Then
            Let My_IsRegister8Bit(ii) = True
            Let My_IsParameter(ii) = True
            Let My_IsRegister(ii) = True
        
        '16-Bit registers:
        ElseIf (ii = TOKEN_Z80_AF) Or (ii = TOKEN_Z80_BC) Or (ii = TOKEN_Z80_DE) Or _
               (ii = TOKEN_Z80_HL) Or (ii = TOKEN_Z80_IX) Or (ii = TOKEN_Z80_IY) Or _
               (ii = TOKEN_Z80_SP) Or (ii = TOKEN_Z80_PC) _
        Then
            Let My_IsRegister16Bit(ii) = True
            Let My_IsParameter(ii) = True
            Let My_IsRegister(ii) = True
        
        'Condition flags:
        ElseIf (ii = TOKEN_Z80_NC) Or (ii = TOKEN_Z80_M) Or (ii = TOKEN_Z80_P) Or _
               (ii = TOKEN_Z80_PE) Or (ii = TOKEN_Z80_PO) Or (ii = TOKEN_Z80_Z) Or _
               (ii = TOKEN_Z80_NZ) _
        Then
            Let My_IsParameter(ii) = True
            Let My_IsFlag(ii) = True
        
        'If the token consititutes an expression (a calculable value)
        ElseIf (ii = TOKEN_LABEL) Or (ii = TOKEN_NUMBER) Or _
               (ii = TOKEN_PROPERTY_USE) _
        Then
            Let My_IsExpression(ii) = True
            Let My_IsParameter(ii) = True
        
        'Z80 parameters are expressions, with the addition of Z80 register / memory _
         references, e.g. `(ix+$FF)`
        ElseIf (ii = TOKEN_Z80MEM_OPEN) Then
            Let My_IsParameter(ii) = True
        End If
    Next ii
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Clean up the timing class
    Set Stopwatch = Nothing
    'Clean up the look up tables
    Erase My_IsInstruction, My_IsRegister, My_IsRegister8Bit, My_IsRegister16Bit, _
          My_IsFlag, My_IsKeyword, My_IsOperator, My_IsExpression, My_IsParameter
    'Empty the token stream
    Erase Tokens
    'Clear the string database
    Set Strings = Nothing
    'Clear the original source file path
    Let Me.SourceFile = vbNullString
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Col : Column number in the source of the current token
'======================================================================================
Public Property Get Col() As Long: Let Col = TokenCol: End Property

'PROPERTY Count : Number of tokens
'======================================================================================
Public Property Get Count() As Long
    'Internally the token count is zero-based (easy to manage), _
     but externally it is one-based (expected behaviour)
    Let Count = TokenCount + 1
End Property

'PROPERTY EOF : End Of File?
'======================================================================================
Public Property Get EOF() As Boolean: Let EOF = (TokenIndex >= TokenCount): End Property

'PROPERTY Error : Expose the internal error number
'======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property

'PROPERTY Index : Current location within the tokenStream
'======================================================================================
Public Property Get Index() As Long: Let Index = TokenIndex: End Property
Public Property Let Index( _
    ByVal Value As Long _
)
    'Out of bounds?
    If (Value < 0) Or (Value > TokenCount) Then
        'Raise a standard VB error for 'array out of bounds'
        Call Err.Raise(9)
    Else
        Let TokenIndex = Value
    End If
End Property

'PROPERTY IsExpression : Is this token number considered a calculable value?
'======================================================================================
Public Property Get IsExpression() As Boolean
    Let IsExpression = My_IsExpression(TokenKind)
End Property

'PROPERTY IsFlag : Is this token number a condition flag?
'======================================================================================
Public Property Get IsFlag() As Boolean
    Let IsFlag = My_IsFlag(TokenKind)
End Property

'PROPERTY IsInstruction : Is this token number a Z80 instruction?
'======================================================================================
Public Property Get IsInstruction() As Boolean
    Let IsInstruction = My_IsInstruction(TokenKind)
End Property

'PROPERTY IsKeyword : Is this token number a OZ80 Keyword?
'======================================================================================
Public Property Get IsKeyword() As Boolean
    Let IsKeyword = My_IsKeyword(TokenKind)
End Property

'PROPERTY IsOperator : Is this token number an OZ80 operator?
'======================================================================================
Public Property Get IsOperator() As Boolean
    Let IsOperator = My_IsOperator(TokenKind)
End Property

'PROPERTY IsPrameter : Is this token number a Z80 instruction parameter?
'======================================================================================
Public Property Get IsParameter() As Boolean
    Let IsParameter = My_IsParameter(TokenKind)
End Property

'PROPERTY IsRegister : Is this token number a Register (8 or 16-Bit)?
'======================================================================================
Public Property Get IsRegister() As Boolean
    Let IsRegister = My_IsRegister(TokenKind)
End Property

'PROPERTY IsRegister8Bit : Is this token number an 8-bit Register?
'======================================================================================
Public Property Get IsRegister8Bit() As Boolean
    Let IsRegister8Bit = My_IsRegister8Bit(TokenKind)
End Property

'PROPERTY IsRegister16Bit : Is this token number a 16-bit Register?
'======================================================================================
Public Property Get IsRegister16Bit() As Boolean
    Let IsRegister16Bit = My_IsRegister16Bit(TokenKind)
End Property

'PROPERTY Kind : Token Kind of current token
'======================================================================================
Public Property Get Kind() As OZ80_TOKEN: Let Kind = TokenKind: End Property

'PROPERTY Line : Line number in the source of the current token
'======================================================================================
Public Property Get Line() As Long: Let Line = TokenLine: End Property

'PROPERTY Value : Token Value of current token
'======================================================================================
Public Property Get Value() As Double: Let Value = TokenValue: End Property

'Forward : Move onto the next token, returns false if the TokenStream hit the end
'======================================================================================
Public Function Forward() As Boolean
    'Generate a debug message with the details of the current token
    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken().Text)
    
    Let TokenIndex = TokenIndex + 1
    Call CacheToken
    
    If TokenIndex >= TokenCount Then Exit Function
    Let Forward = True
End Function

'Rewind : Go back to the start of the tokenStream
'======================================================================================
Public Sub Rewind()
    Let TokenIndex = 0
    Call CacheToken
End Sub

'Tokenise : Break the source file into individual words and convert to machine tokens
'======================================================================================
'FilePath       | Path to the source file to tokenise into this class instance
'Returns        | The error number
'======================================================================================
Public Function Tokenise( _
    ByRef FilePath As String _
) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    Let My_Error = OZ80_ERROR_NONE
    
    'Measure the time this takes
    Call Stopwatch.Reset
    
    'Just a loop counter
    Dim ii As Long
    
    'Prepare a large buffer to accept the token stream (using `ReDim` for every token _
     added is *incredibly* slow). If the tokens overflow this buffer, we'll increase _
     it by another chunk there and then
    Let TokenLimit = TOKEN_BUFFER: Let TokenCount = 0
    ReDim Tokens(0 To TokenLimit) As oz80Token
    '(call this after ReDimming, as it will cache the first token info)
    Call Me.Rewind
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    'Check the file even exists first...
    If Not FileExists(FilePath) _
        Then Call RaiseError(OZ80_ERROR_FILE_NOTFOUND): GoTo Finish
    
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_TOKENISE_READFILE)
    
    'Read the file into our fast string handler (though we won't be doing any string- _
     manipulation, we will be relying on the various methods available to us)
    Dim Source As bluString: Set Source = New bluString
    'Note the string will not be increasing in size, so we can trim the buffer _
     down to the same length as the content
    Call Source.ReadFile(FilePath).ReduceBuffer

    Call oz80.Profiler.EventEnd(EVENT_TOKENISE_READFILE)
    '*** TIMER END ***
    
    'Cache this for speed, we'll be referring to it every loop
    Dim SourceLen As Long: Let SourceLen = Source.Length
    
    RaiseEvent Message(OZ80_LOG_ACTION, _
        "Tokenising " & Chr(ASC_QUOT) & FilePath & Chr(ASC_QUOT) _
    )
    RaiseEvent Message(OZ80_LOG_INFO, _
        "Length " & Format$(SourceLen, "#,#") & " bytes" _
    )
    
    'Clear the string database, just to be extra safe
    Call Strings.RemoveAll
    
    'Begin parsing words: _
     ----------------------------------------------------------------------------------
    'Set the buffer for capturing the word to 1024 chars to avoid lots of memory _
     re-allocations if larger strings in the source are encountered
    Set Word = New bluString
    Let Word.Buffer = BUFFER_1K
    
    'These will track the "mode" of the text as we go
    Dim IsQuote As Long
    Dim IsComment As Long
    Dim IsCommentMulti As Long
    
    'These are used whilst converting words into real numbers
    Dim Number As Double, Multiplier As Long
    
    'We want to track the opening and closing of braces & brackets so that we can _
     attach information to allow skipping from the start to the end (and back)
    Dim Nesting() As TokenNesting
    Dim NestingDepth As Long
    
    'Which character in the source code we're parsing. Since the way the routine is _
     written (we increase the position before each word begins) begin with -1 so that _
     we may really begin at 0
    Dim SourcePos As Long
    Let SourcePos = -1
    
    'Line and column number in the source text. Used to record the position in the _
     original file of each word for producing helpful error messages. 1-based
    Dim Line As Long, Col As Long
    Let Line = 1: Let PrevLine = 1
    
    'Read a single word: _
     ..................................................................................
ReadWord:
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_TOKENISE_READWORD)
    
    'If the word has begun (first letter is read)
    Dim IsWord As Long: Let IsWord = False
    'Reset the captured word
    Call Word.Clear
SkipChar:
    'Move on to the next character in the source
    Let SourcePos = SourcePos + 1
ReadChar:
    'If the file ends?
    If SourcePos >= SourceLen Then
        'If the word is currently underway then we are all done
        If Not IsWord Then GoTo Finish
        'Otherwise process this final word
        GoTo EndWord
    End If
    
    'Read a character:
    Dim AscChar As ASCII
    Let AscChar = Source.Char(SourcePos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the tokenising stage)
    
    'New line:
    If AscChar = ASC_LF Then
        Let Line = Line + 1: Let Col = 0
    
    'NOTE: Carriage-Returns are ignored to avoid doubling line number and to _
           maintain compatibility with UNIX line-endings which use Line-Feed only
    ElseIf AscChar = ASC_CR Then
        GoTo SkipChar
    
    ElseIf AscChar = ASC_TAB Then
        'Tab goes to the next tab stop boundary:
        Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH)
    
    'Non-printable characters are skipped and do not add to the column number
    ElseIf Not Source.CharIsVisible(SourcePos) Then
        GoTo SkipChar
    Else
        Let Col = Col + 1
    End If
    
    'Comments: ........................................................................
    If IsComment Then
        'For single-line comments:
        If Not IsCommentMulti Then
            'The end of a line ends a single-line comment:
            If AscChar = ASC_LF Then Let IsComment = False: GoTo ReadWord
            'Is this a multi-line comment? (two backticks)
            If IsWord Then
                If AscChar = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
        
        'For multi-line comments, check for back-tick:
        ElseIf AscChar = SYNTAX_COMMENT Then
            'Check for a second back-tick following
            Let SourcePos = SourcePos + 1: Let Col = Col + 1
            If Source.Char(SourcePos) = SYNTAX_COMMENT Then
                Let IsComment = False: Let IsCommentMulti = False
                GoTo ReadWord
            End If
        End If
        'Skip to the next character in the comment
        GoTo SkipChar
    
    'Parsable Text: ...................................................................
    '(NOTE: Falsey checks are slightly faster)
    ElseIf Not IsQuote Then
        'Special handling for the beginning of a word
        If Not IsWord Then
            'Skip whitespace before the word begins
            If AscChar = ASC_TAB Then GoTo SkipChar
            If AscChar = ASC_SPC Then GoTo SkipChar
            If AscChar = ASC_LF Then GoTo SkipChar
            
            'Punctuation characters that are only one letter long: _
             When we hit one of these we need to end the word immediately, but parsing _
             normally ends a word on the letter after, so we deal with the edge case
            If AscChar = SYNTAX_NEXT Then GoTo Punctuation
            If AscChar = SYNTAX_Z80MEM_OPEN Then GoTo Punctuation
            If AscChar = SYNTAX_Z80MEM_CLOSE Then GoTo Punctuation
            If AscChar = SYNTAX_CHUNK_OPEN Then GoTo Punctuation
            If AscChar = SYNTAX_CHUNK_CLOSE Then GoTo Punctuation
            
            If AscChar = SYNTAX_OPERATOR_ADD Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_SUB Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_MUL Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_DIV Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_POW Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_MOD Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_OR Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_AND Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_NOT Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_XOR Then GoTo Punctuation
        Else
            'A word is underway, check for end of line
            If AscChar = ASC_LF Then
                'The end of a line ends a single-line comment:
                If Not IsCommentMulti Then
                    If IsComment Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo EndWord
            End If
            
            'End the word on a space
            If AscChar = ASC_SPC Then GoTo EndWord
            If AscChar = ASC_TAB Then GoTo EndWord
            
            'Symbols that can break a word in the middle without spaces to separate
            If AscChar = SYNTAX_NEXT Then GoTo WordBreak
            If AscChar = SYNTAX_Z80MEM_OPEN Then GoTo WordBreak
            If AscChar = SYNTAX_Z80MEM_CLOSE Then GoTo WordBreak
            If AscChar = SYNTAX_CHUNK_OPEN Then GoTo WordBreak
            If AscChar = SYNTAX_CHUNK_CLOSE Then GoTo WordBreak
            
            If AscChar = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_POW Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_OR Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_AND Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_NOT Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_XOR Then GoTo WordBreak
        End If
        
        'Is this a quote? (in which case, ignore wordbreaks until quote end). _
         the opening quote is included so that we can recognise it for tokenising
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = True: GoTo NextChar
        
        'Check for comment marker
        If AscChar = SYNTAX_COMMENT Then
            Let IsComment = True
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If Not IsWord Then GoTo NextChar
            'Split the word away from the comment
            GoTo WordBreak
            
        End If
        
    'Strings: .........................................................................
    Else
        'Skip over the closing quote so that it doesn't get tokenised
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = False: GoTo EndWord
        
    End If
    
NextChar:
    Dim WordBegin As Long, WordEnd As Long
    
    'Is this the first character of the word? If so, record the starting position _
     of the word and flag as word in-progress
    If Not IsWord Then
        Let WordBegin = SourcePos: Let LineBegin = Line: Let ColBegin = Col
        Call Word.Clear: Let IsWord = True
    End If
    'Since this is a valid letter, capture it
    Call Word.CharPush(AscChar)
    GoTo SkipChar

    '----------------------------------------------------------------------------------
Punctuation:
    Let WordBegin = SourcePos
    Let WordEnd = SourcePos
    Call Word.Clear.CharPush(AscChar)
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let SourcePos = SourcePos - 1: Let Col = Col - 1
    Let WordEnd = SourcePos
    GoTo TokenWord
    
EndWord:
    Let WordEnd = SourcePos - 1
    
    'Tokenise the word: _
     ----------------------------------------------------------------------------------
TokenWord:
    '*** TIMER END ***
    Call oz80.Profiler.EventEnd(EVENT_TOKENISE_READWORD)
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_TOKENISE_TOKENWORD)
    
    'Create an upper-case copy of the word for case-insensitive character comparisons
    Dim WordUCase As bluString: Set WordUCase = Nothing
    Set WordUCase = Word.Clone.UCase
    'Extract the 1st, 2nd, 3rd & 4th letters of the word as these make up most _
     comparisons and can be done quickly without resorting to strings
    Dim Asc1 As Long, Asc2 As Long, Asc3 As Long, Asc4 As Long
    Let Asc1 = WordUCase.Char(0)
    If Word.Length > 3 Then Let Asc4 = WordUCase.Char(3) Else Let Asc4 = ASC_NUL
    If Word.Length > 2 Then Let Asc3 = WordUCase.Char(2) Else Let Asc3 = ASC_NUL
    If Word.Length > 1 Then Let Asc2 = WordUCase.Char(1) Else Let Asc2 = ASC_NUL
    
    'Check the first letter of the word to find the right kind of token:
    
    'Comma (very common!) _
     .......................................................................... `,` ...
    If Asc1 = SYNTAX_NEXT Then
        'NOTE: At this point we can only see one word at a time (and not look ahead), _
               so we cannot know when a list has started, and even if the grammer is _
               valid at all (could have a number in completely the wrong place)
        GoTo NextWord
    
    'A decimal number? _
     ........................................................................ `0-9` ...
    ElseIf Word.CharIsNumeric(0, True) Then
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = Word.Length - 1 To 0 Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If Not Word.CharIsNumeric(ii, True) Then
                Debug.Print Word.Text
                Stop
                
                'ERROR: Invalid number!
                Call RaiseError(OZ80_ERROR_INVALID_NUMBER_DEC)
                GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (Word.Char(ii) - ASC_0) * Multiplier
                'Don't allow the multiplier to overflow when the number doesn't
                If ii > 0 Then Let Multiplier = Multiplier * 10
                On Error GoTo 0
            End If
        Next ii
        'Add the token with the number value
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
        
    'Binary number? _
     .......................................................................... `%` ...
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then
        'ERROR: Cannot be longer than 32-bits
        If Word.Length > 33 Then GoTo ErrorOverflow
        Let Number = 0: Let Multiplier = 0
        For ii = Word.Length - 1 To 1 Step -1
            If Word.Char(ii) = ASC_1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf Word.Char(ii) <> ASC_0 Then
                'Error - not a valid binary number
                Stop
            End If
            Let Multiplier = Multiplier + 1
        Next ii
        'Tokenise number
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
    
    'RAM name, e.g. `$.gamevar` _
     ......................................................................... `$.` ...
    ElseIf Asc1 = SYNTAX_NUMBER_HEX _
       And Asc2 = SYNTAX_PROPERTY Then
        'ERROR: Invalid RAM name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_RAM)
            GoTo Finish
        End If
        
        'TODO: Define RAM constant
        
        GoTo NextWord
        
    'Hexadecimal number _
     .......................................................................... `$` ...
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then
        'Any more than 8-digits will be an overflow (32-bit number maximum)
        If Word.Length > 9 Then GoTo ErrorOverflow
        'Check the digits are 0-9, A-F
        For ii = 1 To Word.Length - 1
            'ERROR: Invalid hexadecimal number!
            If Not Word.CharIsHex(ii) Then Call RaiseError( _
                OZ80_ERROR_INVALID_NUMBER_HEX _
            ): GoTo Finish
        Next ii
        
        'Now cast the string into a real number _
         (Note: The digits are valid, but it could be an overflow)
        On Error GoTo ErrorOverflow
        Let Number = Val("&H" & Mid$(Word.Text, 2) & "&")
        On Error GoTo 0
        
        'Add the token for the number
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
    
    '.......................................................................... `:` ...
    ElseIf Asc1 = SYNTAX_LABEL Then
        'A label can be `:` only, without a name (when used within a PROC)
        If Asc2 = ASC_NUL Then
            'This is tokenised as having a label index of 0
            Call Add(TOKEN_LABEL)
            GoTo NextWord
        
        'A label cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf Not IsValidName() Then
            'TODO: Label specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
        Else
            'Is this a section name (`::section`) or regular label (`:label`)?
            If Asc2 = SYNTAX_LABEL Then _
                Call AddString(TOKEN_SECTION) Else _
                Call AddString(TOKEN_LABEL)
            GoTo NextWord
        End If
    
    '.......................................................................... `.` ...
    ElseIf Asc1 = SYNTAX_PROPERTY Then
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        
        'A property can be `.` only, without a name (when used within a PROC)
        '-- currently undecided about this
        If Asc2 = ASC_NUL Then
            'TODO: Property specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
            GoTo NextWord
        
        'A property cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf Not IsValidName() Then
            'TODO: Property specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
        Else
            'If the property appears on a new line, it's a label-property definition _
             i.e. it defines a new sub-label of a chunk
            If LineBegin <> PrevLine Then _
                Call AddString(TOKEN_PROPERTY_NEW) Else _
                Call AddString(TOKEN_PROPERTY_USE)
            GoTo NextWord
        End If
        
    'String (quote) _
     .......................................................................... `"` ...
    ElseIf Asc1 = SYNTAX_QUOTE Then
        'TODO: Break the string into individual number tokens so it's an inline list?
        Call AddString(TOKEN_QUOTE)
        GoTo NextWord
        
    ElseIf Asc1 = ASC_A Then '................................................. `a` ...
        'Z80 instructions
        If Word.Equals("adc") Then Call Add(TOKEN_Z80_ADC): GoTo NextWord
        If Word.Equals("add") Then Call Add(TOKEN_Z80_ADD): GoTo NextWord
        If Word.Equals("and") Then Call Add(TOKEN_Z80_AND): GoTo NextWord
         
    ElseIf Asc1 = ASC_B Then '................................................. `b` ...
        'Z80 instructions
        If Word.Equals("bit") Then Call Add(TOKEN_Z80_BIT): GoTo NextWord
        
    ElseIf Asc1 = ASC_C Then '................................................. `c` ...
        'Z80 instructions
        If Word.Equals("call") Then Call Add(TOKEN_Z80_CALL): GoTo NextWord
        If Word.Equals("ccf") Then Call Add(TOKEN_Z80_CCF): GoTo NextWord
        If Word.Equals("cp") Then Call Add(TOKEN_Z80_CP): GoTo NextWord
        If Word.Equals("cpd") Then Call Add(TOKEN_Z80_CPD): GoTo NextWord
        If Word.Equals("cpdr") Then Call Add(TOKEN_Z80_CPDR): GoTo NextWord
        If Word.Equals("cpi") Then Call Add(TOKEN_Z80_CPI): GoTo NextWord
        If Word.Equals("cpir") Then Call Add(TOKEN_Z80_CPIR): GoTo NextWord
        If Word.Equals("cpl") Then Call Add(TOKEN_Z80_CPL): GoTo NextWord
        'Z80 flag condition
        If Asc2 = 0 Then Call Add(TOKEN_Z80_C): GoTo NextWord
        
    ElseIf Asc1 = ASC_D Then '................................................. `d` ...
        'Z80 instructions
        If Word.Equals("daa") Then Call Add(TOKEN_Z80_DAA): GoTo NextWord
        If Word.Equals("dec") Then Call Add(TOKEN_Z80_DEC): GoTo NextWord
        If Word.Equals("di") Then Call Add(TOKEN_Z80_DI): GoTo NextWord
        If Word.Equals("djnz") Then Call Add(TOKEN_Z80_DJNZ): GoTo NextWord
        
    ElseIf Asc1 = ASC_E Then '................................................. `e` ...
        'Z80 instructions
        If Word.Equals("ei") Then Call Add(TOKEN_Z80_EI): GoTo NextWord
        If Word.Equals("ex") Then Call Add(TOKEN_Z80_EX): GoTo NextWord
        If Word.Equals("exx") Then Call Add(TOKEN_Z80_EXX): GoTo NextWord
        
    ElseIf Asc1 = ASC_F Then '................................................. `f` ...
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = ASC_G Then '................................................. `g` ...
        'There's nothing that begins with G
        GoTo WordError
        
    ElseIf Asc1 = ASC_H Then '................................................. `h` ...
        'Z80 instructions
        If Word.Equals("hlt") Then Call Add(TOKEN_Z80_HALT): GoTo NextWord
        If Word.Equals("halt") Then Call Add(TOKEN_Z80_HALT): GoTo NextWord
        'Keywords
        If Word.Equals("help") Then Call Add(TOKEN_KEYWORD_HELP): GoTo NextWord
        
    ElseIf Asc1 = ASC_I Then '................................................. `i` ...
        'Z80 instructions
        If Word.Equals("im") Then Call Add(TOKEN_Z80_IM): GoTo NextWord
        If Word.Equals("in") Then Call Add(TOKEN_Z80_IN): GoTo NextWord
        If Word.Equals("inc") Then Call Add(TOKEN_Z80_INC): GoTo NextWord
        If Word.Equals("ind") Then Call Add(TOKEN_Z80_IND): GoTo NextWord
        If Word.Equals("indr") Then Call Add(TOKEN_Z80_INDR): GoTo NextWord
        If Word.Equals("ini") Then Call Add(TOKEN_Z80_INI): GoTo NextWord
        If Word.Equals("inir") Then Call Add(TOKEN_Z80_INIR): GoTo NextWord
        'keywords
        If Word.Equals("interrupt") Then Call Add(TOKEN_KEYWORD_INTERRUPT): GoTo NextWord
        
    ElseIf Asc1 = ASC_J Then '................................................. `j` ...
        'Z80 instructions
        If Word.Equals("jp") Then Call Add(TOKEN_Z80_JP): GoTo NextWord
        If Word.Equals("jr") Then Call Add(TOKEN_Z80_JR): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_K Then '................................................. `k` ...
        'Number suffixes K, KB & Kbit
        If Asc2 = ASC_NUL Then
            'Just `K`
            Call Add(TOKEN_PREFIX_K): GoTo NextWord
        
        'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = ASC_B Then
            'If only two letters, then it's `KB`
            If Asc3 = ASC_NUL Then Call Add(TOKEN_PREFIX_KB): GoTo NextWord
            'verify ...`Kbit`, we don't have Asc5 so avoid things like `Kbitty`
            If Word.Equals("kbit") Then Call Add(TOKEN_PREFIX_KBIT): GoTo NextWord
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_L Then '................................................. `l` ...
        'Z80 instructions
        If Word.Equals("ld") Then Call Add(TOKEN_Z80_LD): GoTo NextWord
        If Word.Equals("ldd") Then Call Add(TOKEN_Z80_LDD): GoTo NextWord
        If Word.Equals("lddr") Then Call Add(TOKEN_Z80_LDDR): GoTo NextWord
        If Word.Equals("ldi") Then Call Add(TOKEN_Z80_LDI): GoTo NextWord
        If Word.Equals("ldir") Then Call Add(TOKEN_Z80_LDIR): GoTo NextWord
    
    ElseIf Asc1 = ASC_M Then '................................................. `m` ...
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_M): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_N Then '................................................. `n` ...
        'Z80 instructions
        If Word.Equals("neg") Then Call Add(TOKEN_Z80_NEG): GoTo NextWord
        If Word.Equals("nop") Then Call Add(TOKEN_Z80_NOP): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_C Then Call Add(TOKEN_Z80_NC): GoTo NextWord
            If Asc2 = ASC_Z Then Call Add(TOKEN_Z80_NZ): GoTo NextWord
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_O Then '................................................. `o` ...
        'Z80 instructions
        If Word.Equals("or") Then Call Add(TOKEN_Z80_OR): GoTo NextWord
        If Word.Equals("out") Then Call Add(TOKEN_Z80_OUT): GoTo NextWord
        If Word.Equals("outd") Then Call Add(TOKEN_Z80_OUTD): GoTo NextWord
        If Word.Equals("otdr") Then Call Add(TOKEN_Z80_OTDR): GoTo NextWord
        If Word.Equals("outi") Then Call Add(TOKEN_Z80_OUTI): GoTo NextWord
        If Word.Equals("otir") Then Call Add(TOKEN_Z80_OTIR): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_P Then '................................................. `p` ...
        'Z80 instructions
        If Word.Equals("pop") Then Call Add(TOKEN_Z80_POP): GoTo NextWord
        If Word.Equals("push") Then Call Add(TOKEN_Z80_PUSH): GoTo NextWord
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_P): GoTo NextWord
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_E Then Call Add(TOKEN_Z80_PE): GoTo NextWord
            If Asc2 = ASC_O Then Call Add(TOKEN_Z80_PO): GoTo NextWord
        End If
        'Keywords
        If Word.Equals("params") Then Call Add(TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If Word.Equals("proc") Then Call Add(TOKEN_KEYWORD_PROC): GoTo NextWord
    
    ElseIf Asc1 = ASC_Q Then '................................................. `q` ...
        'There's nothing that begins with Q
        GoTo WordError
        
    ElseIf Asc1 = ASC_R Then '................................................. `r` ...
        'Z80 instructions
        If Word.Equals("res") Then Call Add(TOKEN_Z80_RES): GoTo NextWord
        If Word.Equals("ret") Then Call Add(TOKEN_Z80_RET): GoTo NextWord
        If Word.Equals("reti") Then Call Add(TOKEN_Z80_RETI): GoTo NextWord
        If Word.Equals("retn") Then Call Add(TOKEN_Z80_RETN): GoTo NextWord
        If Word.Equals("rla") Then Call Add(TOKEN_Z80_RLA): GoTo NextWord
        If Word.Equals("rl") Then Call Add(TOKEN_Z80_RL): GoTo NextWord
        If Word.Equals("rlca") Then Call Add(TOKEN_Z80_RLCA): GoTo NextWord
        If Word.Equals("rlc") Then Call Add(TOKEN_Z80_RLC): GoTo NextWord
        If Word.Equals("rld") Then Call Add(TOKEN_Z80_RLD): GoTo NextWord
        If Word.Equals("rra") Then Call Add(TOKEN_Z80_RRA): GoTo NextWord
        If Word.Equals("rr") Then Call Add(TOKEN_Z80_RR): GoTo NextWord
        If Word.Equals("rrca") Then Call Add(TOKEN_Z80_RRCA): GoTo NextWord
        If Word.Equals("rrc") Then Call Add(TOKEN_Z80_RRC): GoTo NextWord
        If Word.Equals("rrd") Then Call Add(TOKEN_Z80_RRD): GoTo NextWord
        If Word.Equals("rst") Then Call Add(TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If Word.Equals("return") Then Call Add(TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = ASC_S Then '................................................. `s` ...
        'Z80 instructions
        If Word.Equals("sbc") Then Call Add(TOKEN_Z80_SBC): GoTo NextWord
        If Word.Equals("scf") Then Call Add(TOKEN_Z80_SCF): GoTo NextWord
        If Word.Equals("set") Then Call Add(TOKEN_Z80_SET): GoTo NextWord
        If Word.Equals("sla") Then Call Add(TOKEN_Z80_SLA): GoTo NextWord
        If Word.Equals("sra") Then Call Add(TOKEN_Z80_SRA): GoTo NextWord
        If Word.Equals("sll") Then Call Add(TOKEN_Z80_SLL): GoTo NextWord
        If Word.Equals("srl") Then Call Add(TOKEN_Z80_SRL): GoTo NextWord
        If Word.Equals("sub") Then Call Add(TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If Word.Equals("section") Then Call Add(TOKEN_KEYWORD_SECTION): GoTo NextWord
        If Word.Equals("slot") Then Call Add(TOKEN_KEYWORD_SLOT): GoTo NextWord
    
    ElseIf Asc1 = ASC_T Then '................................................. `t` ...
        'There is no `t` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_U Then '................................................. `u` ...
        'There's nothing that begins with U
        GoTo WordError
    
    ElseIf Asc1 = ASC_V Then '................................................. `v` ...
        'There's nothing that begins with V
        GoTo WordError
        
    ElseIf Asc1 = ASC_W Then '................................................. `w` ...
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_X Then '................................................. `x` ...
        'If only `x`, then it's the repeat operator
        If Word.IsNull Then Call Add(TOKEN_OPERATOR_REP): GoTo NextWord
        'Z80 instructions
        If Word.Equals("xor") Then Call Add(TOKEN_Z80_XOR): GoTo NextWord
        'There is no `x` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_Y Then '................................................. `y` ...
        'There's nothing that begins with Y
        GoTo WordError
    
    ElseIf Asc1 = ASC_Z Then '................................................. `z` ...
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_Z): GoTo NextWord
        'There is no 'z' register
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_CHUNK_OPEN Then '..................................... `{` ...
        'TODO: Handle nesting
        Call Add(TOKEN_CHUNK_OPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_CHUNK_CLOSE Then  '................................... `}` ...
        'TODO: Handle nesting
        Call Add(TOKEN_CHUNK_CLOSE): GoTo NextWord
    ElseIf Asc1 = SYNTAX_Z80MEM_OPEN Then '.................................... `(` ...
        'TODO: Handle nesting
        Call Add(TOKEN_Z80MEM_OPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_Z80MEM_CLOSE Then '................................... `)` ...
        'TODO: Handle nesting
        Call Add(TOKEN_Z80MEM_CLOSE): GoTo NextWord
    ElseIf Asc1 = SYNTAX_HASH_OPEN Then '...................................... `[` ...
        'TODO: Handle nesting
        Call Add(TOKEN_HASH_OPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_HASH_CLOSE Then '..................................... `]` ...
        'TODO: Handle nesting
        Call Add(TOKEN_HASH_CLOSE): GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_OPERATOR_ADD Then '................................... `+` ...
        Call Add(TOKEN_OPERATOR_ADD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '................................... `-` ...
        Call Add(TOKEN_OPERATOR_SUB): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '................................... `*` ...
        Call Add(TOKEN_OPERATOR_MUL): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '................................... `/` ...
        Call Add(TOKEN_OPERATOR_DIV): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '................................... `^` ...
        Call Add(TOKEN_OPERATOR_POW): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '................................... `\` ...
        Call Add(TOKEN_OPERATOR_MOD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_OR Then  '................................... `|` ...
        Call Add(TOKEN_OPERATOR_OR): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_AND Then '................................... `&` ...
        Call Add(TOKEN_OPERATOR_AND): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_NOT Then '................................... `!` ...
        Call Add(TOKEN_OPERATOR_NOT): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_XOR Then '................................... `~` ...
        Call Add(TOKEN_OPERATOR_XOR): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < ASC_A Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > ASC_S Then GoTo WordError
    'If there's only 1 letter total, then it's obviously a 1-letter register
    If Asc2 = ASC_NUL Then GoTo OneLetter
    'If a semi-colon occurs at the second letter, then the register must be 1 letter
    If Asc2 = SYNTAX_HINT1 Then GoTo OneLetter
    'Also an apostrophe (shadow register) is the same
    If Asc2 = SYNTAX_HINT2 Then GoTo OneLetter
    
    'If there's only 2 letters in total, it has to be a two-letter register
    If Asc3 = ASC_NUL Then GoTo TwoLetters
    'If a semi-colon or apostrophe occurs at the third letter, _
     it could be a two-letter register
    If Asc3 = SYNTAX_HINT1 Then GoTo TwoLetters
    If Asc3 = SYNTAX_HINT2 Then GoTo TwoLetters
    
    'with all else failing, is it three letters?
    If Asc4 = ASC_NUL Then GoTo ThreeLetters
    If Asc4 = SYNTAX_HINT1 Then GoTo ThreeLetters
    If Asc4 = SYNTAX_HINT2 Then GoTo ThreeLetters
    
    GoTo WordError

OneLetter:
    If Asc1 = ASC_A Then Call Add(TOKEN_Z80_A): GoTo NextWord
    If Asc1 = ASC_B Then Call Add(TOKEN_Z80_B): GoTo NextWord
    If Asc1 = ASC_C Then Call Add(TOKEN_Z80_C): GoTo NextWord
    If Asc1 = ASC_D Then Call Add(TOKEN_Z80_D): GoTo NextWord
    If Asc1 = ASC_E Then Call Add(TOKEN_Z80_E): GoTo NextWord
    If Asc1 = ASC_H Then Call Add(TOKEN_Z80_H): GoTo NextWord
    If Asc1 = ASC_I Then Call Add(TOKEN_Z80_I): GoTo NextWord
    If Asc1 = ASC_L Then Call Add(TOKEN_Z80_L): GoTo NextWord
    If Asc1 = ASC_R Then Call Add(TOKEN_Z80_R): GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = ASC_A Then                '"af"
        If Asc2 = ASC_F Then Call Add(TOKEN_Z80_AF): GoTo NextWord
    ElseIf Asc1 = ASC_B Then            '"bc"
        If Asc2 = ASC_C Then Call Add(TOKEN_Z80_BC): GoTo NextWord
    ElseIf Asc1 = ASC_D Then            '"de"
        If Asc2 = ASC_E Then Call Add(TOKEN_Z80_DE): GoTo NextWord
    ElseIf Asc1 = ASC_H Then            '"hl"
        If Asc2 = ASC_L Then Call Add(TOKEN_Z80_HL): GoTo NextWord
    ElseIf Asc1 = ASC_I Then            '"ix"/"iy"
        If Asc2 = ASC_X Then Call Add(TOKEN_Z80_IX): GoTo NextWord
        If Asc2 = ASC_Y Then Call Add(TOKEN_Z80_IY): GoTo NextWord
    ElseIf Asc1 = ASC_P Then            '"pc"
        If Asc2 = ASC_C Then Call Add(TOKEN_Z80_PC): GoTo NextWord
    ElseIf Asc1 = ASC_S Then            '"sp"
        If Asc2 = ASC_P Then Call Add(TOKEN_Z80_SP): GoTo NextWord
    End If
    GoTo WordError
    
ThreeLetters:
    If Asc1 = ASC_I Then
        If Asc2 = ASC_X Then
            If Asc3 = ASC_H Then Call Add(TOKEN_Z80_IXH): GoTo NextWord
            If Asc3 = ASC_L Then Call Add(TOKEN_Z80_IXL): GoTo NextWord
        ElseIf Asc2 = ASC_Y Then
            If Asc3 = ASC_H Then Call Add(TOKEN_Z80_IYH): GoTo NextWord
            If Asc3 = ASC_L Then Call Add(TOKEN_Z80_IYL): GoTo NextWord
        End If
    End If
    GoTo WordError
    
    '----------------------------------------------------------------------------------
NextWord:
    Call oz80.Profiler.EventEnd(EVENT_TOKENISE_TOKENWORD)
    '*** TIMER END ***
    
    Dim WordCount As Long
    Let WordCount = WordCount + 1
    GoTo ReadWord

ErrorOverflow: '.......................................................................
    On Error GoTo 0
    Call RaiseError(OZ80_ERROR_OVERFLOW)
    GoTo Finish
    
WordError: '...........................................................................
    Call RaiseError(OZ80_ERROR_INVALID_WORD)
    GoTo Finish
    
    '----------------------------------------------------------------------------------
Finish:
    If My_Error = OZ80_ERROR_NONE Then
        'Reduce the buffer to the final size
        Let TokenLimit = TokenCount - 1
        ReDim Preserve Tokens(0 To TokenLimit) As oz80Token
        'Move the cursor back to the start of the stream
        Call Me.Rewind
        
        'When the file parses correctly, we can now expose the original file path. _
         This will be used to resolve relative paths with file includes
        Let Me.SourceFile = FilePath
        
        RaiseEvent Message(OZ80_LOG_INFO, "Words  " & Format$(WordCount, "#,#"))
        RaiseEvent Message(OZ80_LOG_INFO, "Tokens " & Format$(TokenCount, "#,#"))
        RaiseEvent Message(OZ80_LOG_INFO, "Time   " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    Else
        'clean up!
        'TODO: This leaves the class in a broken state _
         (no tokens array anymore and no call to `cacheToken`)
        Erase Tokens
        Let TokenCount = 0
    End If
    
    'Free up the memory used for the text file _
     (we won't be referring to the original text again)
    Set Source = Nothing
    'Other miscellaneous clean-up
    Set WordUCase = Nothing
    
    'Return the error number to the assembler
    Let Tokenise = My_Error
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Add : Add a token (any) to the stream
'======================================================================================
Private Sub Add( _
             ByRef Token As OZ80_TOKEN, _
    Optional ByRef Value As Double = 0 _
)
    'The `Token` param MUST be 0-255, but is given as long for speed reasons; _
     the rest of this class will ensure that this is never out of bounds
    With Tokens(TokenIndex)
        Let .Kind = Token And &HFF
        Let .Value = Value
        'Note that the line and column numbers are global to this class to avoid _
         having to use them in calls hundreds of times over
        Let .Line = LineBegin
        Let .Col = ColBegin
    End With
    
    'Generate a debug message
    Call CacheToken
    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken().Text)
    
    'Remember the line number used for this token to check for properties occuring _
     on a new line -- this constitutes a property definition in a chunk
    Let PrevLine = LineBegin
    
    'Move to the next token number
    Let TokenIndex = TokenIndex + 1
    'This will also likely increase the total number of tokens
    If TokenIndex > TokenCount Then Let TokenCount = TokenIndex
    
    'Note that the tokens array is given a large size to begin with as ReDimming it _
     with every token added is *incredibly* slow. If we go over this buffer size, _
     we need to increase it by a large chunk. Once parsing is complete the buffer _
     will be trimmed down to the final size
    If TokenCount > TokenLimit Then
        Let TokenLimit = TokenLimit + TOKEN_BUFFER
        ReDim Preserve Tokens(0 To TokenLimit) As oz80Token
    End If
    
    Call CacheToken
End Sub

'AddString : Strings are stored outside of the token stream and referenced by number
'======================================================================================
Private Sub AddString( _
    ByRef Token As OZ80_TOKEN _
)
    'Note Line & Col numbers are global variables to this class to save having to _
     use them in calls hundreds of times over. Again, speed over properness
    
    'Ignore the quote mark that is left on the front of OZ80 strings
    If Token = TOKEN_QUOTE Then Call Word.Mid(1)
    
    'Hash the string to quickly identify if it has already been added
    If Not Strings.Exists(Word.CRC) Then Call Strings.Add(Word.CRC, Word.Text)
    
    'Add the index key to the token so that we can retrieve the string later
    Call Add(Token, CDbl(Word.CRC))
End Sub

'CacheToken : Cache the current token info
'======================================================================================
Private Sub CacheToken()
    If TokenIndex > UBound(Tokens) Then
        Let TokenKind = 0
        Let TokenValue = 0
        Let TokenLine = 0
        Let TokenCol = 0
        Exit Sub
    End If
    With Tokens(TokenIndex)
        Let TokenKind = .Kind
        Let TokenValue = .Value
        Let TokenLine = .Line
        Let TokenCol = .Col
    End With
End Sub

'FileExists : See if a file exists or not
'======================================================================================
'<cuinl.tripod.com/Tips/fileexist.htm>
Private Function FileExists( _
    ByRef Path As String _
) As Boolean
    Let FileExists = CBool(Dir$(Path) <> vbNullString)
End Function

'FormatLong : Format a 32-bit hexadecimal number for logging
'======================================================================================
Private Function FormatLong( _
    ByRef Value As Long _
) As String
    Let FormatLong = "$" & Right$("0000000" & Hex(Value), 8)
    Let FormatLong = Left$(FormatLong, 5) & "-" & Right$(FormatLong, 4)
End Function

'FormatToken : Format the details of the current token for logging
'======================================================================================
Private Function FormatToken() As bluString
    'Set the template we'll be putting the data into
    Dim Template As New bluString
    Let Template.Text = FORMAT_TOKEN
    
    'Format the Token number
    Dim Format_TokenIndex As New bluString
    Let Format_TokenIndex.Text = Format$(TokenIndex + 1, "#,#")
    
    'TODO: need to check how the 64-bit number gets clipped to 32-bits here?
    Dim Format_TokenValue As New bluString
    Let Format_TokenValue.Text = FormatLong(CLng(TokenValue))
    
    Dim Format_Msg As New bluString
    Dim Format_TokenType As String
    
    If TokenKind = TOKEN_Z80MEM_OPEN _
    Or TokenKind = TOKEN_CHUNK_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_CHUNK
        
    ElseIf My_IsKeyword(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = UCase(oz80.TokenName(TokenKind))
        Let Format_TokenType = FORMAT_TOKEN_KEYWORD
        
    ElseIf TokenKind = TOKEN_LABEL Then
        '..............................................................................
        Let Format_Msg.Text = Chr(34) & Strings(TokenValue) & Chr(34)
        Let Format_TokenType = FORMAT_TOKEN_LABEL
        
    ElseIf TokenKind = TOKEN_NUMBER Then
        '..............................................................................
        Let Format_Msg.Text = Format$(Value, "#,#")
        Let Format_TokenType = FORMAT_TOKEN_NUMBER
        
    ElseIf My_IsOperator(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_OPERATOR
        
    ElseIf TokenKind = TOKEN_PREFIX_K _
        Or TokenKind = TOKEN_PREFIX_KB _
        Or TokenKind = TOKEN_PREFIX_KBIT Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_PREFIX
        
    ElseIf TokenKind = TOKEN_PROPERTY_NEW _
        Or TokenKind = TOKEN_PROPERTY_USE Then
        '..............................................................................
        Let Format_Msg.Text = Chr(34) & Strings(TokenValue) & Chr(34)
        Let Format_TokenType = FORMAT_TOKEN_PROPERTY
        
    ElseIf TokenKind = TOKEN_QUOTE Then
        '..............................................................................
        Let Format_Msg.Text = Chr(34) & Strings(TokenValue) & Chr(34)
        Let Format_TokenType = FORMAT_TOKEN_QUOTE
        
    ElseIf TokenKind = TOKEN_RAM Then
        '..............................................................................
        Let Format_Msg.Text = Chr(34) & Strings(TokenValue) & Chr(34)
        Let Format_TokenType = FORMAT_TOKEN_RAM
        
    ElseIf My_IsRegister(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_REGISTER
        
    ElseIf TokenKind = TOKEN_SECTION Then
        '..............................................................................
        Let Format_Msg.Text = Chr(34) & Strings(TokenValue) & Chr(34)
        Let Format_TokenType = FORMAT_TOKEN_SECTION
        
    ElseIf My_IsInstruction(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_Z80
        
    ElseIf TokenKind = TOKEN_Z80MEM_OPEN _
        Or TokenKind = TOKEN_Z80MEM_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_MEMZ80
    End If
    
    Set FormatToken = Template.Format( _
        Format_TokenIndex, _
        Right$("    " & TokenLine, 5), Left$(TokenCol & "  ", 3), _
        Format_TokenValue, Format_TokenType, _
        Format_Msg _
    )
End Function

'IsValidName : Check validity for named items
'======================================================================================
Private Function IsValidName() As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    
    'The first letter must be the type identifier; _
     `:label`, `?function`, `@macro`, `.property`, `#object`, `$.ram`
    If Word.Char(0) <> SYNTAX_LABEL And _
       Word.Char(0) <> SYNTAX_FUNCT And _
       Word.Char(0) <> SYNTAX_MACRO And _
       Word.Char(0) <> SYNTAX_PROPERTY And _
       Word.Char(0) <> SYNTAX_OBJECT And _
       Word.Char(0) <> SYNTAX_NUMBER_HEX _
    Then
        Exit Function
    End If
    
    'Check for special names with two character prefixes, _
     e.g. "::section" and "$.RAMconstant"
    Dim Begin As Long
    'TODO: Watch out for word-length! Could be a 1-letter label
    If (Word.Char(0) = SYNTAX_LABEL And Word.Char(1) = SYNTAX_LABEL) _
    Or (Word.Char(0) = SYNTAX_NUMBER_HEX And Word.Char(1) = SYNTAX_PROPERTY) Then
        Let Begin = 2
    Else
        Let Begin = 1
    End If
    
    'The name cannot begin with a number or a dot
    If Word.Char(Begin) = SYNTAX_PROPERTY Then Exit Function
    If Word.CharIsNumeric(Begin, True) Then Exit Function
    
    'Cannot end in a dot
    Dim Last As Long
    Let Last = Word.Length - 1
    If Word.Char(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim i As Long
    For i = Begin To Last
        'Not a dot?
        If Word.Char(i) <> SYNTAX_PROPERTY Then
            'Is letter not alphanumeric?
            If Not Word.CharIsAlphaNumeric(i, True) Then
                'Finally, not an underscore?
                If Word.Char(i) <> ASC_USC Then Exit Function
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If Word.CharIsNumeric(i + 1, True) Then Exit Function
            'Two dots in a row not allowed
            If Word.Char(i + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next i
    
    Let IsValidName = True
End Function

'RiaseError : Provide error details through an event
'======================================================================================
Private Sub RaiseError( _
    ByRef Number As OZ80_ERROR _
)
    'Set the error number on this class
    Let My_Error = Number
    
    'Send the error to the controller (if it sinks events)
    RaiseEvent Error(Number, LineBegin, ColBegin)
End Sub
