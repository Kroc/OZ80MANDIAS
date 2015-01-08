VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS:: oz80_File

'A token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. This class handles _
 parsing the source code text into tokens; the assembler only works with tokens
 
'This class only represents one source code file. The assembler handles the specifics _
 of including additional source code files and spawns more oz80_File instances

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'A lot of the text parsing will be letter-by-letter, so this enum provides the _
 common character codes we'll be referring to, as well as the particular punctuation _
 symbols used for OZ80 syntax. Note that `UCS2` is defined by 'bluString.bas'
Public Enum OZ80_SYNTAX
    ASC_NUL
    ASC_TAB = UCS2.C0_HT
    ASC_LF = UCS2.C0_LF
    ASC_CR = UCS2.C0_CR
    
    ASC_SPC = UCS2.SPACE_
    ASC_USC = UCS2.LOW_LINE
    
    ASC_A = UCS2.LATIN_CAPITAL_LETTER_A
:   ASC_B:      ASC_C:      ASC_D:      ASC_E:      ASC_F:      ASC_G:      ASC_H
:   ASC_I:      ASC_J:      ASC_K:      ASC_L:      ASC_M:      ASC_N:      ASC_O
:   ASC_P:      ASC_Q:      ASC_R:      ASC_S:      ASC_T:      ASC_U:      ASC_V
:   ASC_W:      ASC_X:      ASC_Y:      ASC_Z

    ASC_0 = UCS2.DIGIT_ZERO
:   ASC_1:      ASC_2:      ASC_3:      ASC_4:      ASC_5:      ASC_6:      ASC_7
:   ASC_8:      ASC_9
    
    ASC_NEG = UCS2.HYPHEN_MINUS
    
    SYNTAX_COMMENT = UCS2.GRAVE_ACCENT              ' ` - Comment marker
    SYNTAX_HINT1 = UCS2.SEMICOLON                   ' ; - Register hint, `a;index`
    SYNTAX_HINT2 = UCS2.APOSTROPHE                  ' ' - Shadow register, `ex af 'af`
    SYNTAX_TEXT = UCS2.QUOTATION_MARK               ' " - String identifier
    SYNTAX_CONSTANT = UCS2.EXCLAMATION_MARK         ' ! - Constant identifier
    SYNTAX_LABEL = UCS2.COLON                       ' : - Label identifier
    SYNTAX_PROPERTY = UCS2.FULL_STOP                ' . - Property identifier
    SYNTAX_HASH = UCS2.NUMBER_SIGN                  ' # - Hash identifier
    SYNTAX_RAM = UCS2.DOLLAR_SIGN                   ' $ - RAM constant, "$.abc"
    SYNTAX_MACRO = UCS2.COMMERCIAL_AT               ' @ - Macro identifier
    SYNTAX_FUNCT = UCS2.QUESTION_MARK               ' ? - Function identifier
    SYNTAX_NUMBER_HEX = UCS2.DOLLAR_SIGN            ' $ - Hexadecimal number, `$FFFF`
    SYNTAX_NUMBER_BIN = UCS2.PERCENT_SIGN           ' % - Binary number, `%10101011`
    SYNTAX_NEXT = UCS2.COMMA                        ' , - Item seperator, optional
    SYNTAX_Z80MEM_OPEN = UCS2.LEFT_PARENTHESIS      ' ( - Memory reference begin
    SYNTAX_Z80MEM_CLOSE = UCS2.RIGHT_PARENHESIS     ' ) - Memory reference end
    SYNTAX_CHUNK_OPEN = UCS2.LEFT_CURLY_BRACKET     ' { - Open brace
    SYNTAX_CHUNK_CLOSE = UCS2.RIGHT_CURLY_BRACKET   ' } - Close brace
    SYNTAX_HASH_OPEN = UCS2.LEFT_SQUARE_BRACKET     ' [ - Hash array open bracket
    SYNTAX_HASH_CLOSE = UCS2.RIGHT_SQUARE_BRACKET   ' ] - Hash array close bracket
    SYNTAX_OPERATOR_ADD = UCS2.PLUS_SIGN            ' + - Add
    SYNTAX_OPERATOR_SUB = UCS2.HYPHEN_MINUS         ' - - Subtract
    SYNTAX_OPERATOR_MUL = UCS2.ASTERISK             ' * - Multiply
    SYNTAX_OPERATOR_DIV = UCS2.SOLIDUS              ' / - Divide
    SYNTAX_OPERATOR_POW = UCS2.CIRCUMFLEX_ACCENT    ' ^ - Power
    SYNTAX_OPERATOR_MOD = UCS2.REVERSE_SOLIDUS      ' \ - Modulus
    SYNTAX_OPERATOR_OR = UCS2.VERTICAL_LINE         ' | - Bitwise OR
    SYNTAX_OPERATOR_AND = UCS2.AMPERSAND            ' & - Bitwise AND
    SYNTAX_OPERATOR_XOR = UCS2.TILDE                ' ~ - Bitwise XOR
End Enum

'The width of a tab stop (this affects the column number given)
Private Const TAB_WIDTH As Long = 8

'--------------------------------------------------------------------------------------

'String constants at class level so that these texts do not have to be temporarily _
 created and thrown away every time we use them. _
 This is a list of the token words recognised
 
Private Const STR_ADC               As String = "adc"
Private Const STR_ADD               As String = "add"
Private Const STR_AND               As String = "and"
Private Const STR_BIT               As String = "bit"
Private Const STR_BOOL              As String = "BOOL"
Private Const STR_BYTE              As String = "BYTE"
Private Const STR_CALL              As String = "call"
Private Const STR_CCF               As String = "ccf"
Private Const STR_CP                As String = "cp"
Private Const STR_CPD               As String = "cpd"
Private Const STR_CPDR              As String = "cpdr"
Private Const STR_CPI               As String = "cpi"
Private Const STR_CPIR              As String = "cpir"
Private Const STR_CPL               As String = "cpl"
Private Const STR_DAA               As String = "daa"
Private Const STR_DEC               As String = "dec"
Private Const STR_DI                As String = "di"
Private Const STR_DJNZ              As String = "djnz"
Private Const STR_DEF               As String = "DEF"
Private Const STR_EI                As String = "ei"
Private Const STR_EX                As String = "ex"
Private Const STR_EXX               As String = "exx"
Private Const STR_HLT               As String = "hlt"
Private Const STR_HALT              As String = "halt"
Private Const STR_HASH              As String = "HASH"
Private Const STR_HELP              As String = "HELP"
Private Const STR_IM                As String = "im"
Private Const STR_IN                As String = "in"
Private Const STR_INC               As String = "inc"
Private Const STR_IND               As String = "ind"
Private Const STR_INDR              As String = "indr"
Private Const STR_INI               As String = "ini"
Private Const STR_INIR              As String = "inir"
Private Const STR_INCLUDE           As String = "INCLUDE"
Private Const STR_INTERRUPT         As String = "INTERRUPT"
Private Const STR_JP                As String = "jp"
Private Const STR_JR                As String = "jr"
Private Const STR_LD                As String = "ld"
Private Const STR_LDD               As String = "ldd"
Private Const STR_LDDR              As String = "lddr"
Private Const STR_LDI               As String = "ldi"
Private Const STR_LDIR              As String = "ldir"
Private Const STR_LONG              As String = "LONG"
Private Const STR_NEG               As String = "neg"
Private Const STR_NOP               As String = "nop"
Private Const STR_NYBL              As String = "NYBL"
Private Const STR_OR                As String = "or"
Private Const STR_OUT               As String = "out"
Private Const STR_OUTD              As String = "outd"
Private Const STR_OTDR              As String = "otdr"
Private Const STR_OUTI              As String = "outi"
Private Const STR_OTIR              As String = "otir"
Private Const STR_POP               As String = "pop"
Private Const STR_PUSH              As String = "push"
Private Const STR_PARAMS            As String = "PARAMS"
Private Const STR_PROC              As String = "PROC"
Private Const STR_RAM               As String = "RAM"
Private Const STR_RES               As String = "res"
Private Const STR_RET               As String = "ret"
Private Const STR_RETI              As String = "reti"
Private Const STR_RETN              As String = "retn"
Private Const STR_RL                As String = "rl"
Private Const STR_RLA               As String = "rla"
Private Const STR_RLC               As String = "rlc"
Private Const STR_RLCA              As String = "rlca"
Private Const STR_RLD               As String = "rld"
Private Const STR_RR                As String = "rr"
Private Const STR_RRA               As String = "rra"
Private Const STR_RRC               As String = "rrc"
Private Const STR_RRCA              As String = "rrca"
Private Const STR_RRD               As String = "rrd"
Private Const STR_RST               As String = "rst"
Private Const STR_RETURN            As String = "RETURN"
Private Const STR_SBC               As String = "sbc"
Private Const STR_SCF               As String = "scf"
Private Const STR_SET               As String = "set"
Private Const STR_SLA               As String = "sla"
Private Const STR_SLL               As String = "sll"
Private Const STR_SRA               As String = "sra"
Private Const STR_SRL               As String = "srl"
Private Const STR_SUB               As String = "sub"
Private Const STR_SECTION           As String = "SECTION"
Private Const STR_SLOT              As String = "SLOT"
Private Const STR_SYSTEM            As String = "SYSTEM"
Private Const STR_TRIP              As String = "TRIP"
Private Const STR_WORD              As String = "WORD"
Private Const STR_XOR               As String = "xor"

'--------------------------------------------------------------------------------------

'For diaplying a disassembly
Private Const FORMAT_QUOTE          As String = """"
Private Const FORMAT_TOKEN          As String = "<tkn> {$} ({$},{$}) [{$}] {$}"
Private Const FORMAT_TOKEN_CHUNK    As String = "   CHUNK"
Private Const FORMAT_TOKEN_CONST    As String = "   CONST"
Private Const FORMAT_TOKEN_HASH     As String = "    HASH"
Private Const FORMAT_TOKEN_KEYWORD  As String = " KEYWORD"
Private Const FORMAT_TOKEN_LABEL    As String = "   LABEL"
Private Const FORMAT_TOKEN_MEMZ80   As String = "  MEMZ80"
Private Const FORMAT_TOKEN_NUMBER   As String = "  NUMBER"
Private Const FORMAT_TOKEN_OPERATOR As String = "OPERATOR"
Private Const FORMAT_TOKEN_PREFIX   As String = "  PREFIX"
Private Const FORMAT_TOKEN_PROPERTY As String = "PROPERTY"
Private Const FORMAT_TOKEN_RAM      As String = "     RAM"
Private Const FORMAT_TOKEN_REGISTER As String = "REGISTER"
Private Const FORMAT_TOKEN_SECTION  As String = " SECTION"
Private Const FORMAT_TOKEN_TEXT     As String = "    TEXT"
Private Const FORMAT_TOKEN_Z80      As String = "     Z80"

Private Const FORMAT_FILE_LENGTH    As String = "Length: {$} Characters"
Private Const FORMAT_FILE_TOKENS    As String = "Tokens: {$}"

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'The error number!
Private My_Error As OZ80_ERROR

'To save having to create and destroy short strings all the time, _
 we'll keep this one around for formatting log messages
Private LogMsg As bluString

'Remember where the original source file is as the assembler will need to change to _
 its directory when handling relative references with includes
Private My_Path As String

'--------------------------------------------------------------------------------------

'Use a 100'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 100000

Private Tokens() As oz80Token           'The token stream will live here
Private TokenLimit As Long              'The length of the buffer used for the above
Private My_Count As Long                'The current amount of space used in the array
Private My_Index As Long                'The current place within the tokenStream

Private Type oz80Token
    Kind As Byte                        '=OZ80_TOKEN, but use 1-byte instead of 4
    Value As Double                     'Associated value, e.g. for number tokens
    Line As Long                        'Line number in the original source text
    Col As Long                         'Column number in the original source text
End Type

'Caches of the current token to make multiple calls to this class a little quicker, _
 it also helps with a lot of internal referencing, mainly with logging
Private My_TokenKind As OZ80_TOKEN
Private My_TokenValue As Double
Private My_TokenLine As Long
Private My_TokenCol As Long

'Look up tables to determine what category a Token belongs to; _
 these are exposed publicly through properties so other classes can do quick checks
Private My_IsExpression(0 To [_TOKEN_LAST]) As Boolean
Private My_IsFlag(0 To [_TOKEN_LAST]) As Boolean
Private My_IsList(0 To [_TOKEN_LAST]) As Boolean
Private My_IsInstruction(0 To [_TOKEN_LAST]) As Boolean
Private My_IsKeyword(0 To [_TOKEN_LAST]) As Boolean
Private My_IsOperator(0 To [_TOKEN_LAST]) As Boolean
Private My_IsParameter(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister8Bit(0 To [_TOKEN_LAST]) As Boolean
Private My_IsRegister16Bit(0 To [_TOKEN_LAST]) As Boolean

'We could just unfold Text into a series of number tokens, but for quick access we _
 store them in their own database, referenced by the Token's Value
Public Strings As bluArrayStrings
'--------------------------------------------------------------------------------------

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

'We track the opening and closing of brackets to validate nesting of brackets _
 and to allow the assembler to skip over blocks
Private Nesting(0 To 255) As TokenNesting
Private NestingDepth As Long

Private Type TokenNesting
    Token As OZ80_TOKEN                 'Which type of bracket `{`, `[`, `(`
    Index As Long                       'The token index of the opening bracket
End Type

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output. _
 A bluString is provided to make manipulating the message before displaying very easy
Event Message( _
    ByRef LogLevel As OZ80_LOG, ByRef LogText As bluString _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    'Any kind of logging message will be built using this
    Set LogMsg = New bluString
    
    'Initialise the string database. _
     Since it's public, we want to ensure it's instantiated from the off
    Set Strings = New bluArrayStrings
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Clean up the look up tables
    Erase My_IsInstruction, My_IsRegister, My_IsRegister8Bit, My_IsRegister16Bit, _
          My_IsFlag, My_IsKeyword, My_IsList, My_IsOperator, My_IsExpression, _
          My_IsParameter
    'Empty the token stream
    Erase Tokens
    'Clear the string database
    Set Strings = Nothing
    
    Set LogMsg = Nothing
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'TokenCol : Column number in the source of the current token
'======================================================================================
Public Property Get TokenCol() As Long: Let TokenCol = My_TokenCol: End Property

'TokenCount : Number of tokens
'======================================================================================
Public Property Get TokenCount() As Long
    'Internally the token count is zero-based (easy to manage), _
     but externally it is one-based (expected behaviour)
    Let TokenCount = My_Count + 1
End Property

'EOF : End Of File?
'======================================================================================
Public Property Get EOF() As Boolean: Let EOF = (My_Index >= My_Count): End Property

'Error : Expose the internal error number
'======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property

'TokenIndex : Current location within the File's token stream
'======================================================================================
Public Property Get TokenIndex() As Long: Let TokenIndex = My_Index: End Property
Public Property Let TokenIndex( _
    ByVal Index As Long _
)   'Out of bounds?
    If (Index < 0) Or (Index > My_Count) Then
        'Raise a standard VB error for 'array out of bounds'
        Call Err.Raise(9)
    Else
        Let My_Index = Index
    End If
End Property

'TokenIsExpression : Is the current token considered a calculable value?
'======================================================================================
Public Property Get TokenIsExpression() As Boolean
    Let TokenIsExpression = My_IsExpression(My_TokenKind)
End Property

'TokenIsFlag : Is the current token a condition flag?
'======================================================================================
Public Property Get TokenIsFlag() As Boolean
    Let TokenIsFlag = My_IsFlag(My_TokenKind)
End Property

'TokenIsInstruction : Is the current token a Z80 instruction?
'======================================================================================
Public Property Get TokenIsInstruction() As Boolean
    Let TokenIsInstruction = My_IsInstruction(My_TokenKind)
End Property

'TokenIsKeyword : Is the current token an OZ80 Keyword?
'======================================================================================
Public Property Get TokenIsKeyword() As Boolean
    Let TokenIsKeyword = My_IsKeyword(My_TokenKind)
End Property

'TokenIsList : Is the current token considered a list item (an Expression or Text)
'======================================================================================
Public Property Get TokenIsList() As Boolean
    Let TokenIsList = My_IsList(My_TokenKind)
End Property

'TokenIsOperator : Is the current token an OZ80 operator?
'======================================================================================
Public Property Get TokenIsOperator() As Boolean
    Let TokenIsOperator = My_IsOperator(My_TokenKind)
End Property

'TokenIsParameter : Is the current token a Z80 instruction parameter?
'======================================================================================
Public Property Get TokenIsParameter() As Boolean
    Let TokenIsParameter = My_IsParameter(My_TokenKind)
End Property

'TokenIsRegister : Is the current token a Register (8 or 16-Bit)?
'======================================================================================
Public Property Get TokenIsRegister() As Boolean
    Let TokenIsRegister = My_IsRegister(My_TokenKind)
End Property

'TokenIsRegister8Bit : Is the current token an 8-bit Register?
'======================================================================================
Public Property Get TokenIsRegister8Bit() As Boolean
    Let TokenIsRegister8Bit = My_IsRegister8Bit(My_TokenKind)
End Property

'TokenIsRegister16Bit : Is the current token a 16-bit Register?
'======================================================================================
Public Property Get TokenIsRegister16Bit() As Boolean
    Let TokenIsRegister16Bit = My_IsRegister16Bit(My_TokenKind)
End Property

'TokenKind : Token Kind of current token
'======================================================================================
Public Property Get TokenKind() As OZ80_TOKEN: Let TokenKind = My_TokenKind: End Property

'TokenLine : Line number in the source of the current token
'======================================================================================
Public Property Get TokenLine() As Long: Let TokenLine = My_TokenLine: End Property

'Path : File path of the file that's been tokenised
'======================================================================================
Public Property Get Path() As String: Let Path = My_Path: End Property

'TokenText : Gets the string associated with a token (e.g. `TOKEN_QUOTE`)
'======================================================================================
Public Property Get TokenText() As String
    'Only valid for tokens that reference text _
     (null-string will be returned otherwise)
    If My_TokenKind = TOKEN_TEXT Then Let TokenText = Strings.Item(My_TokenValue)
End Property

'TokenValue : Token Value of current token
'======================================================================================
Public Property Get TokenValue() As Double: Let TokenValue = My_TokenValue: End Property

'NextToken : Move onto the next token
'======================================================================================
'Returns        | False if the token stream hit the end
'======================================================================================
Public Function NextToken() As Boolean
    'Generate a debug message with the details of the current token
    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken())
    
    Let My_Index = My_Index + 1
    Call CacheToken
    
    If My_Index >= My_Count Then Exit Function
    Let NextToken = True
End Function

'Rewind : Go back to the start of the token stream
'======================================================================================
Public Sub Rewind()
    Let My_Index = 0
    Call CacheToken
End Sub

'Tokenise : Break the source file into individual words and convert to machine tokens
'======================================================================================
'FilePath       | Path to the source file to tokenise into this class instance
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Public Function Tokenise( _
    ByRef FilePath As String _
) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    Let My_Error = OZ80_ERROR_NONE
    
    '----------------------------------------------------------------------------------
    'Build the lookup tables
    '----------------------------------------------------------------------------------
    'This eats up a chunk of memory so we won't do this at class initialisation
    Dim t As OZ80_TOKEN
    For t = 0 To [_TOKEN_LAST]
        'Check what token category it belongs to
        '..............................................................................
        Let My_IsInstruction(t) = (t > [_TOKEN_INSTRUCTIONS_BEGIN]) And _
                                  (t < [_TOKEN_INSTRUCTIONS_END])
        Let My_IsOperator(t) = (t > [_TOKEN_OPERATORS_BEGIN]) And _
                               (t < [_TOKEN_OPERATORS_END])
        Let My_IsKeyword(t) = (t > [_TOKEN_KEYWORDS_BEGIN]) And _
                              (t < [_TOKEN_KEYWORDS_END])
        
        'Registers and flags are more complex as there's some overlap
        '..............................................................................
        If (t = TOKEN_Z80_C) Then
            'At the token level, we can't differentiate Register C and Flag C; _
             we need context for that and we can't gaurantee the TokenStream is _
             valid until we begin assembling, therefore we treat Register C and _
             Flag C as the same and the assembler will sort out the difference
            Let My_IsRegister8Bit(t) = True
            Let My_IsParameter(t) = True
            Let My_IsRegister(t) = True
            Let My_IsFlag(t) = True
        
        '8-Bit registers:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_A) Or (t = TOKEN_Z80_B) Or (t = TOKEN_Z80_C) Or _
               (t = TOKEN_Z80_D) Or (t = TOKEN_Z80_E) Or (t = TOKEN_Z80_H) Or _
               (t = TOKEN_Z80_L) Or (t = TOKEN_Z80_I) Or (t = TOKEN_Z80_R) Or _
               (t = TOKEN_Z80_IXL) Or (t = TOKEN_Z80_IXH) Or _
               (t = TOKEN_Z80_IYL) Or (t = TOKEN_Z80_IYH) _
        Then
            Let My_IsRegister8Bit(t) = True
            Let My_IsParameter(t) = True
            Let My_IsRegister(t) = True
        
        '16-Bit registers:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_AF) Or (t = TOKEN_Z80_BC) Or (t = TOKEN_Z80_DE) Or _
               (t = TOKEN_Z80_HL) Or (t = TOKEN_Z80_IX) Or (t = TOKEN_Z80_IY) Or _
               (t = TOKEN_Z80_SP) Or (t = TOKEN_Z80_PC) _
        Then
            Let My_IsRegister16Bit(t) = True
            Let My_IsParameter(t) = True
            Let My_IsRegister(t) = True
        
        'Condition flags:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_NC) Or (t = TOKEN_Z80_M) Or (t = TOKEN_Z80_P) Or _
               (t = TOKEN_Z80_PE) Or (t = TOKEN_Z80_PO) Or (t = TOKEN_Z80_Z) Or _
               (t = TOKEN_Z80_NZ) _
        Then
            Let My_IsParameter(t) = True
            Let My_IsFlag(t) = True
        
        'If the token consititutes an Expression (a calculable Value)
        '..............................................................................
        ElseIf (t = TOKEN_LABEL) Or (t = TOKEN_NUMBER) Or _
               (t = TOKEN_PROPERTY_USE) Or (t = TOKEN_OPERATOR_SUB) _
        Then
            Let My_IsExpression(t) = True
            Let My_IsParameter(t) = True
            Let My_IsList(t) = True
        
        'Text classes as a self-contained List; _
         all other List items are Expressions
        '..............................................................................
        ElseIf (t = TOKEN_TEXT) Then
            Let My_IsList(t) = True
        
        'Z80 parameters are Expressions, with the addition of Z80 register / memory _
         references, e.g. `(ix+$FF)`
        '..............................................................................
        ElseIf (t = TOKEN_Z80MEM_OPEN) Then
            Let My_IsParameter(t) = True
        End If
    Next t
    
    '----------------------------------------------------------------------------------
    'Read the whole file into a single string
    '----------------------------------------------------------------------------------
    'Check the file even exists first...
    If Not FileExists(FilePath) _
        Then Call RaiseError(OZ80_ERROR_FILE_NOTFOUND): GoTo Finish
    
    'Store the file path in the class. This is so that you can remember which file _
     is which after you've tokenised them
    Let My_Path = FilePath
    
    'Read the file into our fast string handler (though we won't be doing any string- _
     manipulation, we will be relying on the various methods available to us)
    Dim Source As bluString: Set Source = New bluString
    'Note the string will not be increasing in size, so we can trim the buffer _
     down to the same length as the content
    'TODO: Handle file reading errors
    Call Source.ReadFile(My_Path).ReduceBuffer
    
    'Cache this for speed, we'll be referring to it every loop
    Dim SourceLen As Long: Let SourceLen = Source.Length
    
    Call RaiseMessage(OZ80_LOG_ACTION, _
        "Tokenising ""{$}""", My_Path _
    )
    
    'Clear the string database, just to be extra safe
    Call Strings.Clear
    
    '----------------------------------------------------------------------------------
    'Begin parsing words: _
     ----------------------------------------------------------------------------------
    'Prepare a large buffer to accept the token stream (using `ReDim` for every token _
     added is *incredibly* slow). If the tokens overflow this buffer, we'll increase _
     it by another chunk there and then
    Let TokenLimit = TOKEN_BUFFER: Let My_Count = 0
    ReDim Tokens(0 To TokenLimit) As oz80Token
    '(call this after ReDimming, as it will cache the first token info)
    Call Me.Rewind
    
    'Set the buffer for capturing the word to 1024 chars to avoid lots of memory _
     re-allocations if larger strings in the source are encountered
    Set Word = New bluString
    Let Word.Buffer = BUFFER_1K
    
    'These will track the "mode" of the text as we go
    Dim IsQuote As Long
    Dim IsComment As Long, IsCommentMulti As Long
    
    'These are used whilst converting words into real numbers
    Dim Number As Double, Multiplier As Long
    
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
    
    'Reset the captured word
    Call Word.Clear
SkipChar:
    'Move on to the next character in the source
    Let SourcePos = SourcePos + 1
ReadChar:
    'If the file ends?
    If SourcePos >= SourceLen Then
        'If no word is currently underway then we are all done
        If Word.IsNull Then GoTo Finish
        'Otherwise process this final word
        GoTo TokenWord
    End If
    
    'Read a character:
    Dim Char As OZ80_SYNTAX
    Let Char = Source.Char(SourcePos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the tokenising stage)
    '..................................................................................
    
    'New line:
    If Char = C0_LF Then
        Let Line = Line + 1: Let Col = 0
    
    'NOTE: Carriage-Returns are ignored to avoid doubling line number and to _
           maintain compatibility with UNIX line-endings which use Line-Feed only
    ElseIf Char = C0_CR Then
        GoTo SkipChar
    
    'Tab:
    ElseIf Char = C0_HT Then
        'Tab goes to the next tab stop boundary:
        Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH)
    
    ElseIf Source.CharIsVisible(SourcePos) Then
        Let Col = Col + 1
        
    Else
        'Non-printable characters are skipped and do not add to the column number
        GoTo SkipChar
    End If
    
    'Comments: ........................................................................
    If IsComment Then
        'For single-line comments:
        If Not IsCommentMulti Then
            'The end of a line ends a single-line comment:
            If Char = C0_LF Then Let IsComment = False: GoTo ReadWord
            'Is this a multi-line comment? (two backticks)
            'TODO: What is going on here???
            If Not Word.IsNull Then
                If Char = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
        
        'For multi-line comments, check for back-tick:
        ElseIf Char = SYNTAX_COMMENT Then
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
    ElseIf IsQuote = 0 Then
        'Special handling for the beginning of a word
        If Word.IsNull Then
            'Skip whitespace before the word begins
            If Source.CharIsSpace(SourcePos) Then GoTo SkipChar
            
            'Punctuation characters that are only one letter long: _
             When we hit one of these we need to end the word immediately, but parsing _
             normally ends a word on the letter after, so we deal with the edge case
            If Char = SYNTAX_NEXT Then GoTo Punctuation
            If Char = SYNTAX_CHUNK_OPEN Then GoTo Punctuation
            If Char = SYNTAX_CHUNK_CLOSE Then GoTo Punctuation
            If Char = SYNTAX_HASH_OPEN Then GoTo Punctuation
            If Char = SYNTAX_HASH_CLOSE Then GoTo Punctuation
            If Char = SYNTAX_Z80MEM_OPEN Then GoTo Punctuation
            If Char = SYNTAX_Z80MEM_CLOSE Then GoTo Punctuation
            
            If Char = SYNTAX_OPERATOR_ADD Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_SUB Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_MUL Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_DIV Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_POW Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_MOD Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_OR Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_AND Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_XOR Then GoTo Punctuation
        Else
            'A word is underway, check for end of line
            If Char = ASC_LF Then
                'The end of a line ends a single-line comment:
                If Not IsCommentMulti Then
                    If IsComment Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo TokenWord
            End If
            
            'End the word on a space
            If Source.CharIsSpace(SourcePos) Then GoTo TokenWord
            
            'Symbols that can break a word in the middle without spaces to separate
            If Char = SYNTAX_NEXT Then GoTo WordBreak
            If Char = SYNTAX_Z80MEM_OPEN Then GoTo WordBreak
            If Char = SYNTAX_Z80MEM_CLOSE Then GoTo WordBreak
            If Char = SYNTAX_CHUNK_OPEN Then GoTo WordBreak
            If Char = SYNTAX_CHUNK_CLOSE Then GoTo WordBreak
            
            If Char = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_POW Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_OR Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_AND Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_XOR Then GoTo WordBreak
        End If
        
        'Is this text? (in which case, ignore wordbreaks until text end). _
         the opening quote is included so that we can recognise it for tokenising
        If Char = SYNTAX_TEXT Then Let IsQuote = True: GoTo NextChar
        
        'Check for comment marker
        If Char = SYNTAX_COMMENT Then
            Let IsComment = True
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If Word.IsNull Then GoTo NextChar
            'Split the word away from the comment
            GoTo WordBreak
            
        End If
        
    'Strings: .........................................................................
    Else
        'Skip over the closing quote so that it doesn't get tokenised
        If Char = SYNTAX_TEXT Then Let IsQuote = False: GoTo TokenWord
    End If
    
NextChar:
    'Is this the first character of the word? _
     If so, record the starting position of the word
    If Word.IsNull Then Let LineBegin = Line: Let ColBegin = Col
    'Since this is a valid letter, capture it
    Call Word.CharPush(Char)
    GoTo SkipChar

    '----------------------------------------------------------------------------------
Punctuation:
    Call Word.Clear.CharPush(Char)
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let SourcePos = SourcePos - 1: Let Col = Col - 1
    
    '----------------------------------------------------------------------------------
    'Tokenise the word:
    '----------------------------------------------------------------------------------
TokenWord:
    '*** TIMER END ***
    Call oz80.Profiler.EventEnd(EVENT_TOKENISE_READWORD)
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_TOKENISE_TOKENWORD)
    
    'Create an upper-case copy of the word for easy ASCII-code comparisons
    Dim WordUCase As bluString: Set WordUCase = Nothing
    Set WordUCase = Word.Clone.UCase
    'Extract the 1st, 2nd, 3rd & 4th letters of the word as these make up most _
     comparisons and can be done quickly without resorting to Strings
    Dim Asc1 As OZ80_SYNTAX, Asc2 As OZ80_SYNTAX, _
        Asc3 As OZ80_SYNTAX, Asc4 As OZ80_SYNTAX
    Let Asc1 = WordUCase.Char(0)
    If Word.Length > 3 Then Let Asc4 = WordUCase.Char(3) Else Let Asc4 = ASC_NUL
    If Word.Length > 2 Then Let Asc3 = WordUCase.Char(2) Else Let Asc3 = ASC_NUL
    If Word.Length > 1 Then Let Asc2 = WordUCase.Char(1) Else Let Asc2 = ASC_NUL
    
    'Just a loop counter
    Dim i As Long
    
    'Comma (very common!) _
     .......................................................................... `,` ...
    If Asc1 = SYNTAX_NEXT Then
        'NOTE: At this point we can only see one word at a time (and not look ahead), _
         so we cannot know when a list has started, and even if the grammer is valid _
         at all (could have a number in completely the wrong place). For that reason, _
         all commas are optional and simply skipped by the tokeniser
        GoTo NextWord
        
    'A decimal number? _
     ........................................................................ `0-9` ...
    ElseIf Word.CharIsNumeric(0, True) Then
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For i = Word.Length - 1 To 0 Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If Not Word.CharIsNumeric(i, True) Then
                Debug.Print Word.Text
                Stop
                
                'ERROR: Invalid number!
                Call RaiseError(OZ80_ERROR_INVALID_NUMBER_DEC)
                GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (Word.Char(i) - ASC_0) * Multiplier
                'Don't allow the multiplier to overflow when the number doesn't
                If i > 0 Then Let Multiplier = Multiplier * 10
                On Error GoTo 0
            End If
        Next i
        'Add the token with the number value
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
    
    'Constant? _
     .......................................................................... `!` ...
    ElseIf Asc1 = SYNTAX_CONSTANT Then
        'ERROR: Invalid hash name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_HASH)
            GoTo Finish
        End If
        
        'TODO: Define constant name
        Call AddString(TOKEN_CONST)
        GoTo NextWord
    
    'Binary number? _
     .......................................................................... `%` ...
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then
        'ERROR: Cannot be longer than 32-bits
        If Word.Length > 33 Then GoTo ErrorOverflow
        Let Number = 0: Let Multiplier = 0
        For i = Word.Length - 1 To 1 Step -1
            If Word.Char(i) = ASC_1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf Word.Char(i) <> ASC_0 Then
                'Error - not a valid binary number
                Stop
            End If
            Let Multiplier = Multiplier + 1
        Next i
        'Tokenise number
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
        
    'Hash name? _
     .......................................................................... `#` ...
    ElseIf Asc1 = SYNTAX_HASH Then
        'ERROR: Invalid hash name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_HASH)
            GoTo Finish
        End If
        
        'TODO: Define hash name
        Call AddString(TOKEN_HASH)
        GoTo NextWord
    
    'RAM name, e.g. `$.gamevar`, `$#RAM_BANK.gamevar` _
     ......................................................................... `$.` ...
    ElseIf Asc1 = SYNTAX_RAM _
    And (Asc2 = SYNTAX_PROPERTY Or Asc2 = SYNTAX_HASH) Then
        'ERROR: Invalid RAM name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_RAM)
            GoTo Finish
        End If
        
        'TODO: Define RAM constant
        Call AddString(TOKEN_RAM)
        GoTo NextWord
        
    'Hexadecimal number _
     .......................................................................... `$` ...
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then
        'Any more than 8-digits will be an overflow (32-bit number maximum)
        If Word.Length > 9 Then GoTo ErrorOverflow
        'Check the digits are 0-9, A-F
        For i = 1 To Word.Length - 1
            'ERROR: Invalid hexadecimal number!
            If Not Word.CharIsHex(i) Then Call RaiseError( _
                OZ80_ERROR_INVALID_NUMBER_HEX _
            ): GoTo Finish
        Next i
        
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
        End If
        GoTo NextWord
    
    '.......................................................................... `.` ...
    ElseIf Asc1 = SYNTAX_PROPERTY Then
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        
        'A property can be `.` only, without a name
        If Asc2 = ASC_NUL Then
            Call AddString(TOKEN_PROPERTY_NEW)
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
        End If
        GoTo NextWord
        
    'Text _
     .......................................................................... `"` ...
    ElseIf Asc1 = SYNTAX_TEXT Then
        Call AddString(TOKEN_TEXT)
        GoTo NextWord
        
    ElseIf Asc1 = ASC_A Then '................................................. `a` ...
        'Z80 instructions
        If Word.Equals(STR_ADC) Then Call Add(TOKEN_Z80_ADC): GoTo NextWord
        If Word.Equals(STR_ADD) Then Call Add(TOKEN_Z80_ADD): GoTo NextWord
        If Word.Equals(STR_AND) Then Call Add(TOKEN_Z80_AND): GoTo NextWord
         
    ElseIf Asc1 = ASC_B Then '................................................. `b` ...
        'Z80 instructions
        If Word.Equals(STR_BIT) Then Call Add(TOKEN_Z80_BIT): GoTo NextWord
        'Keywords
        If Word.Equals(STR_BOOL) Then Call Add(TOKEN_KEYWORD_BOOL): GoTo NextWord
        If Word.Equals(STR_BYTE) Then Call Add(TOKEN_KEYWORD_BYTE): GoTo NextWord
        
    ElseIf Asc1 = ASC_C Then '................................................. `c` ...
        'Z80 instructions
        If Word.Equals(STR_CALL) Then Call Add(TOKEN_Z80_CALL): GoTo NextWord
        If Word.Equals(STR_CCF) Then Call Add(TOKEN_Z80_CCF): GoTo NextWord
        If Word.Equals(STR_CP) Then Call Add(TOKEN_Z80_CP): GoTo NextWord
        If Word.Equals(STR_CPD) Then Call Add(TOKEN_Z80_CPD): GoTo NextWord
        If Word.Equals(STR_CPDR) Then Call Add(TOKEN_Z80_CPDR): GoTo NextWord
        If Word.Equals(STR_CPI) Then Call Add(TOKEN_Z80_CPI): GoTo NextWord
        If Word.Equals(STR_CPIR) Then Call Add(TOKEN_Z80_CPIR): GoTo NextWord
        If Word.Equals(STR_CPL) Then Call Add(TOKEN_Z80_CPL): GoTo NextWord
        'Z80 flag condition
        If Asc2 = 0 Then Call Add(TOKEN_Z80_C): GoTo NextWord
        
    ElseIf Asc1 = ASC_D Then '................................................. `d` ...
        'Z80 instructions
        If Word.Equals(STR_DAA) Then Call Add(TOKEN_Z80_DAA): GoTo NextWord
        If Word.Equals(STR_DEC) Then Call Add(TOKEN_Z80_DEC): GoTo NextWord
        If Word.Equals(STR_DI) Then Call Add(TOKEN_Z80_DI): GoTo NextWord
        If Word.Equals(STR_DJNZ) Then Call Add(TOKEN_Z80_DJNZ): GoTo NextWord
        'Keywords
        If Word.Equals(STR_DEF) Then Call Add(TOKEN_KEYWORD_DEF): GoTo NextWord
        
    ElseIf Asc1 = ASC_E Then '................................................. `e` ...
        'Z80 instructions
        If Word.Equals(STR_EI) Then Call Add(TOKEN_Z80_EI): GoTo NextWord
        If Word.Equals(STR_EX) Then Call Add(TOKEN_Z80_EX): GoTo NextWord
        If Word.Equals(STR_EXX) Then Call Add(TOKEN_Z80_EXX): GoTo NextWord
        
    ElseIf Asc1 = ASC_F Then '................................................. `f` ...
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = ASC_G Then '................................................. `g` ...
        'There's nothing that begins with G
        GoTo WordError
        
    ElseIf Asc1 = ASC_H Then '................................................. `h` ...
        'Z80 instructions
        If Word.Equals(STR_HLT) Then Call Add(TOKEN_Z80_HALT): GoTo NextWord
        If Word.Equals(STR_HALT) Then Call Add(TOKEN_Z80_HALT): GoTo NextWord
        'Keywords
        If Word.Equals(STR_HASH) Then Call Add(TOKEN_KEYWORD_HASH): GoTo NextWord
        If Word.Equals(STR_HELP) Then Call Add(TOKEN_KEYWORD_HELP): GoTo NextWord
        
    ElseIf Asc1 = ASC_I Then '................................................. `i` ...
        'Z80 instructions
        If Word.Equals(STR_IM) Then Call Add(TOKEN_Z80_IM): GoTo NextWord
        If Word.Equals(STR_IN) Then Call Add(TOKEN_Z80_IN): GoTo NextWord
        If Word.Equals(STR_INC) Then Call Add(TOKEN_Z80_INC): GoTo NextWord
        If Word.Equals(STR_IND) Then Call Add(TOKEN_Z80_IND): GoTo NextWord
        If Word.Equals(STR_INDR) Then Call Add(TOKEN_Z80_INDR): GoTo NextWord
        If Word.Equals(STR_INI) Then Call Add(TOKEN_Z80_INI): GoTo NextWord
        If Word.Equals(STR_INIR) Then Call Add(TOKEN_Z80_INIR): GoTo NextWord
        'keywords
        If Word.Equals(STR_INCLUDE) Then Call Add(TOKEN_KEYWORD_INCLUDE): GoTo NextWord
        If Word.Equals(STR_INTERRUPT) Then Call Add(TOKEN_KEYWORD_INTERRUPT): GoTo NextWord
        
    ElseIf Asc1 = ASC_J Then '................................................. `j` ...
        'Z80 instructions
        If Word.Equals(STR_JP) Then Call Add(TOKEN_Z80_JP): GoTo NextWord
        If Word.Equals(STR_JR) Then Call Add(TOKEN_Z80_JR): GoTo NextWord
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
        If Word.Equals(STR_LD) Then Call Add(TOKEN_Z80_LD): GoTo NextWord
        If Word.Equals(STR_LDD) Then Call Add(TOKEN_Z80_LDD): GoTo NextWord
        If Word.Equals(STR_LDDR) Then Call Add(TOKEN_Z80_LDDR): GoTo NextWord
        If Word.Equals(STR_LDI) Then Call Add(TOKEN_Z80_LDI): GoTo NextWord
        If Word.Equals(STR_LDIR) Then Call Add(TOKEN_Z80_LDIR): GoTo NextWord
        'Keywords
        If Word.Equals(STR_LONG) Then Call Add(TOKEN_KEYWORD_LONG): GoTo NextWord
    
    ElseIf Asc1 = ASC_M Then '................................................. `m` ...
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_M): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_N Then '................................................. `n` ...
        'Z80 instructions
        If Word.Equals(STR_NEG) Then Call Add(TOKEN_Z80_NEG): GoTo NextWord
        If Word.Equals(STR_NOP) Then Call Add(TOKEN_Z80_NOP): GoTo NextWord
        'Keywords
        If Word.Equals(STR_NYBL) Then Call Add(TOKEN_KEYWORD_NYBL): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_C Then Call Add(TOKEN_Z80_NC): GoTo NextWord
            If Asc2 = ASC_Z Then Call Add(TOKEN_Z80_NZ): GoTo NextWord
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_O Then '................................................. `o` ...
        'Z80 instructions
        If Word.Equals(STR_OR) Then Call Add(TOKEN_Z80_OR): GoTo NextWord
        If Word.Equals(STR_OUT) Then Call Add(TOKEN_Z80_OUT): GoTo NextWord
        If Word.Equals(STR_OUTD) Then Call Add(TOKEN_Z80_OUTD): GoTo NextWord
        If Word.Equals(STR_OTDR) Then Call Add(TOKEN_Z80_OTDR): GoTo NextWord
        If Word.Equals(STR_OUTI) Then Call Add(TOKEN_Z80_OUTI): GoTo NextWord
        If Word.Equals(STR_OTIR) Then Call Add(TOKEN_Z80_OTIR): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_P Then '................................................. `p` ...
        'Z80 instructions
        If Word.Equals(STR_POP) Then Call Add(TOKEN_Z80_POP): GoTo NextWord
        If Word.Equals(STR_PUSH) Then Call Add(TOKEN_Z80_PUSH): GoTo NextWord
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_P): GoTo NextWord
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_E Then Call Add(TOKEN_Z80_PE): GoTo NextWord
            If Asc2 = ASC_O Then Call Add(TOKEN_Z80_PO): GoTo NextWord
        End If
        'Keywords
        If Word.Equals(STR_PARAMS) Then Call Add(TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If Word.Equals(STR_PROC) Then Call Add(TOKEN_KEYWORD_PROC): GoTo NextWord
    
    ElseIf Asc1 = ASC_Q Then '................................................. `q` ...
        'There's nothing that begins with Q
        GoTo WordError
        
    ElseIf Asc1 = ASC_R Then '................................................. `r` ...
        'Z80 instructions
        If Word.Equals(STR_RES) Then Call Add(TOKEN_Z80_RES): GoTo NextWord
        If Word.Equals(STR_RET) Then Call Add(TOKEN_Z80_RET): GoTo NextWord
        If Word.Equals(STR_RETI) Then Call Add(TOKEN_Z80_RETI): GoTo NextWord
        If Word.Equals(STR_RETN) Then Call Add(TOKEN_Z80_RETN): GoTo NextWord
        If Word.Equals(STR_RL) Then Call Add(TOKEN_Z80_RL): GoTo NextWord
        If Word.Equals(STR_RLA) Then Call Add(TOKEN_Z80_RLA): GoTo NextWord
        If Word.Equals(STR_RLC) Then Call Add(TOKEN_Z80_RLC): GoTo NextWord
        If Word.Equals(STR_RLCA) Then Call Add(TOKEN_Z80_RLCA): GoTo NextWord
        If Word.Equals(STR_RLD) Then Call Add(TOKEN_Z80_RLD): GoTo NextWord
        If Word.Equals(STR_RR) Then Call Add(TOKEN_Z80_RR): GoTo NextWord
        If Word.Equals(STR_RRA) Then Call Add(TOKEN_Z80_RRA): GoTo NextWord
        If Word.Equals(STR_RRC) Then Call Add(TOKEN_Z80_RRC): GoTo NextWord
        If Word.Equals(STR_RRCA) Then Call Add(TOKEN_Z80_RRCA): GoTo NextWord
        If Word.Equals(STR_RRD) Then Call Add(TOKEN_Z80_RRD): GoTo NextWord
        If Word.Equals(STR_RST) Then Call Add(TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If Word.Equals(STR_RAM) Then Call Add(TOKEN_KEYWORD_RAM): GoTo NextWord
        If Word.Equals(STR_RETURN) Then Call Add(TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = ASC_S Then '................................................. `s` ...
        'Z80 instructions
        If Word.Equals(STR_SBC) Then Call Add(TOKEN_Z80_SBC): GoTo NextWord
        If Word.Equals(STR_SCF) Then Call Add(TOKEN_Z80_SCF): GoTo NextWord
        If Word.Equals(STR_SET) Then Call Add(TOKEN_Z80_SET): GoTo NextWord
        If Word.Equals(STR_SLA) Then Call Add(TOKEN_Z80_SLA): GoTo NextWord
        If Word.Equals(STR_SLL) Then Call Add(TOKEN_Z80_SLL): GoTo NextWord
        If Word.Equals(STR_SRA) Then Call Add(TOKEN_Z80_SRA): GoTo NextWord
        If Word.Equals(STR_SRL) Then Call Add(TOKEN_Z80_SRL): GoTo NextWord
        If Word.Equals(STR_SUB) Then Call Add(TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If Word.Equals(STR_SECTION) Then Call Add(TOKEN_KEYWORD_SECTION): GoTo NextWord
        If Word.Equals(STR_SLOT) Then Call Add(TOKEN_KEYWORD_SLOT): GoTo NextWord
        If Word.Equals(STR_SYSTEM) Then Call Add(TOKEN_KEYWORD_SYSTEM): GoTo NextWord
        
    ElseIf Asc1 = ASC_T Then '................................................. `t` ...
        'Keywords
        If Word.Equals(STR_TRIP) Then Call Add(TOKEN_KEYWORD_TRIP): GoTo NextWord
        'There is no `t` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_U Then '................................................. `u` ...
        'There's nothing that begins with U
        GoTo WordError
    
    ElseIf Asc1 = ASC_V Then '................................................. `v` ...
        'There's nothing that begins with V
        GoTo WordError
        
    ElseIf Asc1 = ASC_W Then '................................................. `w` ...
        'Keywords
        If Word.Equals(STR_WORD) Then Call Add(TOKEN_KEYWORD_WORD): GoTo NextWord
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_X Then '................................................. `x` ...
        'If only `x`, then it's the repeat operator
        If Asc2 = ASC_NUL Then Call Add(TOKEN_OPERATOR_REP): GoTo NextWord
        'Z80 instructions
        If Word.Equals(STR_XOR) Then Call Add(TOKEN_Z80_XOR): GoTo NextWord
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
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_CHUNK_OPEN) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_CHUNK_CLOSE Then  '................................... `}` ...
        'Close the current nesting level, error if incorrect bracket-matching
        If NestOut(TOKEN_CHUNK_OPEN, TOKEN_CHUNK_CLOSE) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_Z80MEM_OPEN Then '.................................... `(` ...
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_Z80MEM_OPEN) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_Z80MEM_CLOSE Then '................................... `)` ...
        'Close the current nesting level, error if incorrect bracket-matching
        If NestOut(TOKEN_Z80MEM_OPEN, TOKEN_Z80MEM_CLOSE) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_HASH_OPEN Then '...................................... `[` ...
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_HASH_OPEN) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_HASH_CLOSE Then '..................................... `]` ...
        'Close the current nesting level, error if incorrect bracket-matching
        If NestOut(TOKEN_HASH_OPEN, TOKEN_HASH_CLOSE) Then GoTo Finish
        GoTo NextWord
    
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
    ElseIf Asc1 = SYNTAX_OPERATOR_XOR Then '................................... `~` ...
        Call Add(TOKEN_OPERATOR_XOR): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
Registers:
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
        Let TokenLimit = My_Count - 1
        ReDim Preserve Tokens(0 To TokenLimit) As oz80Token
        'Move the cursor back to the start of the stream
        Call Me.Rewind
        
        Call RaiseMessage(OZ80_LOG_INFO, _
            FORMAT_FILE_LENGTH, Format$(SourceLen, "#,0") _
        )
        Call RaiseMessage(OZ80_LOG_INFO, _
            FORMAT_FILE_TOKENS, Format$(My_Count, "#,0") _
        )
    Else
        'clean up!
        'TODO: This leaves the class in a broken state _
         (no tokens array anymore and no call to `cacheToken`)
        Erase Tokens: Let My_Count = 0
    End If
    
    'Free up the memory used for the text file _
     (we won't be referring to the original text again)
    Set Source = Nothing
    'Clear the nesting information
    Erase Nesting: Let NestingDepth = 0
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
    With Tokens(My_Index)
        Let .Kind = Token And &HFF
        Let .Value = Value
        'Note that the line and column numbers are global to this class to avoid _
         having to use them in calls hundreds of times over
        Let .Line = LineBegin
        Let .Col = ColBegin
    End With
    
    'Generate a debug message
    Call CacheToken
    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken())
    
    'Remember the line number used for this token to check for properties occuring _
     on a new line -- this constitutes a property definition in a chunk
    Let PrevLine = LineBegin
    
    'Move to the next token number
    Let My_Index = My_Index + 1
    'This will also likely increase the total number of tokens
    If My_Index > My_Count Then Let My_Count = My_Index
    
    'Note that the tokens array is given a large size to begin with as ReDimming it _
     with every token added is *incredibly* slow. If we go over this buffer size, _
     we need to increase it by a large chunk. Once parsing is complete the buffer _
     will be trimmed down to the final size
    If My_Count > TokenLimit Then
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
    If Token = TOKEN_TEXT Then Call Word.Mid(1)
    
    'Add the index key to the token so that we can retrieve the string later
    Call Add(Token, Strings.Add(Word.Text))
End Sub

'CacheToken : Cache the current token info
'======================================================================================
Private Sub CacheToken()
    If My_Index > UBound(Tokens) Then
        Let My_TokenKind = 0
        Let My_TokenValue = 0
        Let My_TokenLine = 0
        Let My_TokenCol = 0
        Exit Sub
    End If
    With Tokens(My_Index)
        Let My_TokenKind = .Kind
        Let My_TokenValue = .Value
        Let My_TokenLine = .Line
        Let My_TokenCol = .Col
    End With
End Sub

'FileExists : See if a file exists or not
'======================================================================================
'FilePath       | Path to the file to check for
'---------------+----------------------------------------------------------------------
'Returns        | True if the file exists, false if not
'======================================================================================
Private Function FileExists( _
    ByRef FilePath As String _
) As Boolean
    '<cuinl.tripod.com/Tips/fileexist.htm>
    Let FileExists = CBool(Dir$(FilePath) <> vbNullString)
End Function

'FormatToken : Format the details of the current token for logging
'======================================================================================
'Returns        | A new bluString containing the details of the current token
'======================================================================================
Private Function FormatToken( _
) As bluString
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_FORMATTOKEN)
    
    'Set the template we'll be putting the data into
    Dim Template As New bluString
    Let Template.Text = FORMAT_TOKEN
    
    'Format the Token number
    Dim Format_TokenIndex As New bluString
    Let Format_TokenIndex.Text = Format$(My_Index + 1, "#,#")
    
    Dim Format_Msg As New bluString
    Dim Format_TokenType As String
    
    If My_TokenKind = TOKEN_CHUNK_OPEN _
    Or My_TokenKind = TOKEN_CHUNK_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_CHUNK
        
    ElseIf My_TokenKind = TOKEN_CONST Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_CONST
        
    ElseIf My_TokenKind = TOKEN_HASH Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_HASH
        
    ElseIf My_TokenKind = TOKEN_HASH_OPEN _
        Or My_TokenKind = TOKEN_HASH_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_HASH
        
    ElseIf My_IsInstruction(My_TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_Z80
        
    ElseIf My_IsKeyword(My_TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Call Format_Msg.UCase
        Let Format_TokenType = FORMAT_TOKEN_KEYWORD
        
    ElseIf My_TokenKind = TOKEN_LABEL Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_LABEL
        
    ElseIf My_TokenKind = TOKEN_NUMBER Then
        '..............................................................................
        'Don't try and display an indefinite value
        If My_TokenValue = OZ80_INDEFINITE _
            Then Let Format_Msg.Text = "?" _
            Else Let Format_Msg.Text = Format$(My_TokenValue, "#,0")
        Let Format_TokenType = FORMAT_TOKEN_NUMBER
        
    ElseIf My_IsOperator(My_TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_OPERATOR
        
    ElseIf My_TokenKind = TOKEN_PREFIX_K _
        Or My_TokenKind = TOKEN_PREFIX_KB _
        Or My_TokenKind = TOKEN_PREFIX_KBIT Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_PREFIX
        
    ElseIf My_TokenKind = TOKEN_PROPERTY_NEW _
        Or My_TokenKind = TOKEN_PROPERTY_USE Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_PROPERTY
            
    ElseIf My_TokenKind = TOKEN_RAM Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_RAM
        
    ElseIf My_IsRegister(My_TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_REGISTER
        
    ElseIf My_TokenKind = TOKEN_SECTION Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_SECTION
        
    ElseIf My_TokenKind = TOKEN_TEXT Then
        '..............................................................................
        Let Format_Msg.Text = Strings.Item(My_TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_TEXT
            
    ElseIf My_TokenKind = TOKEN_Z80MEM_OPEN _
        Or My_TokenKind = TOKEN_Z80MEM_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = oz80.TokenName(My_TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_MEMZ80
    
    Else
        Debug.Print oz80.TokenName(My_TokenKind)
        Stop
    End If
    
    Set FormatToken = Template.Format( _
        Format_TokenIndex, _
        Right$("    " & My_TokenLine, 5), Left$(My_TokenCol & "  ", 3), _
        Format_TokenType, _
        Format_Msg _
    )
        
    '*** TIMER END ***
    Call oz80.Profiler.EventEnd(EVENT_FORMATTOKEN)
End Function

'IsValidName : Check validity for named items
'======================================================================================
'Returns        | True if the captured word conforms to name item syntax
'======================================================================================
Private Function IsValidName( _
) As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    Dim WordChar0 As OZ80_SYNTAX
    Let WordChar0 = Word.Char(0)
    
    'The first letter must be the type identifier; _
     `!CONST`, `:label`, `.property`, `#hash`, `$.RAM`, `?function`, `@macro`
    If WordChar0 <> SYNTAX_CONSTANT And _
       WordChar0 <> SYNTAX_LABEL And _
       WordChar0 <> SYNTAX_PROPERTY And _
       WordChar0 <> SYNTAX_HASH And _
       WordChar0 <> SYNTAX_RAM And _
       WordChar0 <> SYNTAX_FUNCT And _
       WordChar0 <> SYNTAX_MACRO _
    Then
        Exit Function
    End If
    
    'Check for special names with two character prefixes, _
     e.g. "::section", "$.RAM" and "$#RAM"
    Dim Begin As Long
    'TODO: Watch out for word-length! Could be a 1-letter label
    'TODO: Could change this function to use bluString's CharPull
    If (WordChar0 = SYNTAX_LABEL And Word.Char(1) = SYNTAX_LABEL) _
    Or (WordChar0 = SYNTAX_RAM And Word.Char(1) = SYNTAX_PROPERTY) _
    Or (WordChar0 = SYNTAX_RAM And Word.Char(1) = SYNTAX_HASH) _
    Then
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

'NestIn : Track an additional level of bracket nesting
'======================================================================================
'BracketType    | Token number to represent "{", "(" or "["
'---------------+----------------------------------------------------------------------
'Returns        | An error number, should the nesting be invalid
'======================================================================================
Private Function NestIn( _
    ByRef BracketType As OZ80_TOKEN _
) As OZ80_ERROR
    'Open another nesting level
    'TODO: Error above 255
    Let NestingDepth = NestingDepth + 1
    Let Nesting(NestingDepth).Token = BracketType
    Let Nesting(NestingDepth).Index = My_Index
    Call Add(BracketType, OZ80_INDEFINITE)
End Function

'NestOut : Finish a level of bracket nesting, folding up to the previous level
'======================================================================================
'BracketOpen    | Token number for the opening bracket "{", "(" or "["
'BracketClose   | Token number for the closing bracket "}", ")" or "]"
'---------------+----------------------------------------------------------------------
'Returns        | An error number, should the nesting be invalid
'======================================================================================
Private Function NestOut( _
    ByRef BracketOpen As OZ80_TOKEN, _
    ByRef BracketClose As OZ80_TOKEN _
) As OZ80_ERROR
    With Nesting(NestingDepth)
        'Check that the nesting is correct; you cannot close a bracket without an _
         open one before, nor can you close a bracket when a different one is open
        If .Token <> BracketOpen Then
            Let NestOut = RaiseError(OZ80_ERROR_EXPECTED_BRACKET)
        End If
        'Update the opening brace to reference the closing brace
        Let Tokens(.Index).Value = My_Index
        'And add the closing brace referencing the opening brace
        Call Add(BracketClose, CDbl(.Index))
    End With
    'Reduce the nesting level
    Let NestingDepth = NestingDepth - 1
End Function

'RaiseMessage : Provide a progress message
'======================================================================================
'LogLevel       | The type of log message, for filtering out lower-level messages
'FormatText     | The message to log, which can include "{$}" markers to be replaced
'               | with the values provided below
'FormatValues...| A list of Variant values (Strings, Longs &c.) to fill in
'======================================================================================
Private Sub RaiseMessage( _
    ByRef LogLevel As OZ80_LOG, _
    ByRef FormatText As String, _
    ParamArray FormatValues() As Variant _
)
    Let LogMsg.Text = FormatText
    
    'Are there any values to format?
    If UBound(FormatValues) < 0 Then
        'No, raise just the text given
        RaiseEvent Message(LogLevel, LogMsg)
    Else
        'Format the given log message before sending it up
        RaiseEvent Message(LogLevel, LogMsg.Format(FormatValues))
    End If
End Sub

'RaiseError : Provide error details through an event
'======================================================================================
'Number         | The error number to set on the Class and raise via Event
'---------------+----------------------------------------------------------------------
'Returns        | The same error number, for easy tail recursion
'======================================================================================
Private Function RaiseError( _
    ByRef Number As OZ80_ERROR _
) As OZ80_ERROR
    'Set the error number on this class
    Let My_Error = Number
    'Send the error to the controller (if it sinks events)
    RaiseEvent Error(Number, LineBegin, ColBegin)
    'Return the error number, for easy tail recursion
    Let RaiseError = Number
End Function
