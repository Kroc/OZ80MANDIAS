VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80Tokeniser"
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
'CLASS:: oz80TokenStream

'A Token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. This class handles _
 parsing the source code text into tokens; the assembler only works with tokens

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'<msdn.microsoft.com/en-us/library/windows/desktop/ms724408%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetTickCount Lib "kernel32" Alias "GetTickCount" ( _
) As Long

'API support for CompStr function _
 --------------------------------------------------------------------------------------
'<www.vbforums.com/showthread.php?729385-VB6-clsStrToIntArray-cls-Cast-String-To-Integer-Array>
Private Type SAFEARRAY1D    'Represents a safe array. (One Dimensional)
    cDims      As Integer   'The count of dimensions.
    fFeatures  As Integer   'Flags used by the SafeArray.
    cbElements As Long      'The size of an array element.
    cLocks     As Long      'The number of times the array has been locked without a corresponding unlock.
    pvData     As Long      'Pointer to the data.
    cElements  As Long      'The number of elements in the dimension.
    lLbound    As Long      'The lower bound of the dimension.
End Type                    'http://msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CompStrPtr  As Long
Private CompStrSA1D As SAFEARRAY1D

Dim CompStrArr() As Integer

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

'This makes life a whole lot easier
Private Enum ASCII
    UNDRSCR = 95
    
    NUM0 = 48
    NUM9 = 57
    
    A = 97
    B = 98
    C = 99
    D = 100
    E = 101
    F = 102
    G = 103
    H = 104
    I = 105
    J = 106
    K = 107
    L = 108
    M = 109
    N = 110
    O = 111
    P = 112
    Q = 113
    R = 114
    S = 115
    T = 116
    U = 117
    V = 118
    W = 119
    X = 120
    Y = 121
    Z = 122

    aSPC = 32
    aTAB = 9
    aCR = 10
    aLF = 13
End Enum

'These define the various punctiation marks (in ASCII codes) for the language syntax
Private Enum SYNTAX
    SYNTAX_COMMENT = 96         '`
    SYNTAX_HINT1 = 59           ';
    SYNTAX_HINT2 = 59           ''
    SYNTAX_QUOTE = 34           '"
    SYNTAX_LABEL = 58           ':
    SYNTAX_PROPERTY = 46        '.
    SYNTAX_VARIABLE = 35        '#
    SYNTAX_MACRO = 64           '@
    SYNTAX_FUNCT = 63           '?
    SYNTAX_NUMBER_HEX = 36      '$
    SYNTAX_NUMBER_BIN = 37      '%
    SYNTAX_NEXT = 44            ',
    SYNTAX_PAREN_OPEN = 40      '(
    SYNTAX_PAREN_CLOSE = 41     ')
    SYNTAX_BLOCK_OPEN = 123     '{
    SYNTAX_BLOCK_CLOSE = 125    '}
    SYNTAX_OPERATOR_ADD = 43    '+
    SYNTAX_OPERATOR_SUB = 45    '-
    SYNTAX_OPERATOR_MUL = 42    '*
    SYNTAX_OPERATOR_DIV = 47    '/
    SYNTAX_OPERATOR_POW = 94    '^
    SYNTAX_OPERATOR_MOD = 92    '\
    SYNTAX_OPERATOR_OR = 124    '|
    SYNTAX_OPERATOR_AND = 38    '&
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_TOKEN
    'These are just the mnemonic tokens -- the assembler itself checks the
     'parameters and determines which opcode should be used
    TOKEN_Z80_ADC = &H1                 'Add with Carry
    TOKEN_Z80_ADD = &H2                 'Add
    TOKEN_Z80_AND = &H3                 'Bitwise AND
    TOKEN_Z80_BIT = &H4                 'Bit test
    TOKEN_Z80_CALL = &H5                'Call routine
    TOKEN_Z80_CCF = &H6                 'Clear Carry Flag
    TOKEN_Z80_CP = &H7                  'Compare
    TOKEN_Z80_CPD = &H8                 'Compare and Decrement
    TOKEN_Z80_CPDR = &H9                'Compare, Decrement and Repeat
    TOKEN_Z80_CPI = &HA                 'Compare and Increment
    TOKEN_Z80_CPIR = &HB                'Compare, Increment and Repeat
    TOKEN_Z80_CPL = &HC                 'Complement (bitwise NOT)
    TOKEN_Z80_DAA = &HD                 'Decimal Adjust Accumulator
    TOKEN_Z80_DEC = &HE                 'Decrement
    TOKEN_Z80_DI = &HF                  'Disable Interrupts
    TOKEN_Z80_DJNZ = &H10               'Decrement and Jump if Not Zero
    TOKEN_Z80_EI = &H11                 'Enable Inettupts
    TOKEN_Z80_EX = &H12                 'Exchange
    TOKEN_Z80_EXX = &H13                'Exchange shadow registers
    TOKEN_Z80_HALT = &H14               'Stop CPU (wait for interrupt)
    TOKEN_Z80_IM = &H15                 'Interrupt Mode
    TOKEN_Z80_IN = &H16                 'Input from port
    TOKEN_Z80_INC = &H17                'Increment
    TOKEN_Z80_IND = &H18                'Input and Decrement
    TOKEN_Z80_INDR = &H19               'Input, Decrement and Repeat
    TOKEN_Z80_INI = &H1A                'Input and Increment
    TOKEN_Z80_INIR = &H1B               'Input, Increment and Repeat
    TOKEN_Z80_JP = &H1C                 'Jump
    TOKEN_Z80_JR = &H1D                 'Jump Relative
    TOKEN_Z80_LD = &H1E                 'Load
    TOKEN_Z80_LDD = &H1F                'Load and Decrement
    TOKEN_Z80_LDDR = &H20               'Load, Decrement and Repeat
    TOKEN_Z80_LDI = &H21                'Load and Increment
    TOKEN_Z80_LDIR = &H22               'Load, Increment and Repeat
    TOKEN_Z80_NEG = &H23                'Negate (flip the sign)
    TOKEN_Z80_NOP = &H24                'No Operation (do nothing)
    TOKEN_Z80_OR = &H25                 'Bitwise OR
    TOKEN_Z80_OUT = &H26                'Output to port
    TOKEN_Z80_OUTD = &H27               'Output and Decrement
    TOKEN_Z80_OUTDR = &H28              'Output, Decrement and Repeat
    TOKEN_Z80_OUTI = &H29               'Output and Increment
    TOKEN_Z80_OUTIR = &H2A              'Output, Increment and Repeat
    TOKEN_Z80_POP = &H2B                'Pull from stack
    TOKEN_Z80_PUSH = &H2C               'Push onto stack
    TOKEN_Z80_RES = &H2D                'Reset bit
    TOKEN_Z80_RET = &H2E                'Return from routine
    TOKEN_Z80_RETI = &H2F               'Return from Interrupt
    TOKEN_Z80_RETN = &H30               'Return from NMI
    TOKEN_Z80_RLA = &H31                'Rotate Left (Accumulator)
    TOKEN_Z80_RL = &H32                 'Rotate Left
    TOKEN_Z80_RLCA = &H33               'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLC = &H34                'Rotate Left Circular
    TOKEN_Z80_RLD = &H35                'Rotate Left 4-bits
    TOKEN_Z80_RRA = &H36                'Rotate Right (Accumulator)
    TOKEN_Z80_RR = &H37                 'Rotate Right
    TOKEN_Z80_RRCA = &H38               'Rotate Right Circular (Accumulator)
    TOKEN_Z80_RRC = &H39                'Rotate Right Circular
    TOKEN_Z80_RRD = &H3A                'Rotate Right 4-bits
    TOKEN_Z80_RST = &H3B                '"Restart" -- Call a page 0 routine
    TOKEN_Z80_SBC = &H3C                'Subtract with Carry
    TOKEN_Z80_SCF = &H3D                'Set Carry Flag
    TOKEN_Z80_SET = &H3E                'Set bit
    TOKEN_Z80_SLA = &H3F                'Shift Left Arithmetic
    TOKEN_Z80_SRA = &H40                'Shift Right Arithmetic
    TOKEN_Z80_SLL = &H41                'Shift Left Logical
    TOKEN_Z80_SRL = &H42                'Shift Right Logical
    TOKEN_Z80_SUB = &H43                'Subtract
    TOKEN_Z80_XOR = &H44                'Bitwise XOR
    
    'Z80 registers
    TOKEN_REGISTER_A = &H50             'Accumulator
    TOKEN_REGISTER_B = &H51
    TOKEN_REGISTER_C = &H52
    TOKEN_REGISTER_D = &H53
    TOKEN_REGISTER_E = &H54
    TOKEN_REGISTER_F = &H55             'Flags register
    TOKEN_REGISTER_H = &H56
    TOKEN_REGISTER_I = &H57             'Interrupt - not to be confused with IX & IY
    TOKEN_REGISTER_L = &H58
    TOKEN_REGISTER_R = &H59             'Refresh register (pseudo-random)

    TOKEN_REGISTER_AF = &H5A
    TOKEN_REGISTER_BC = &H5B
    TOKEN_REGISTER_DE = &H5C
    TOKEN_REGISTER_HL = &H5D
    TOKEN_REGISTER_IX = &H5E
    TOKEN_REGISTER_IY = &H5F
    
    TOKEN_REGISTER_SP = &H60            'Stack pointer
    TOKEN_REGISTER_PC = &H61            'Program counter
    
    'Undocumented Z80 instructions can access the 8-bit halves of IX & IY
    TOKEN_REGISTER_IXL = &H62
    TOKEN_REGISTER_IXH = &H63
    TOKEN_REGISTER_IYL = &H64
    TOKEN_REGISTER_IYH = &H65
    
    'Z80 flag conditions used on JP, CALL &  RET
    TOKEN_FLAG_C = &H68                 'Carry set
    TOKEN_FLAG_NC = &H69                'Carry not set
    TOKEN_FLAG_Z = &H6A                 'Zero set
    TOKEN_FLAG_NZ = &H6B                'Zero not set
    TOKEN_FLAG_M = &H6C                 'Sign is set
    TOKEN_FLAG_P = &H6D                 'Sign is not set
    TOKEN_FLAG_PE = &H6E                'Parity/Overflow is set
    TOKEN_FLAG_PO = &H6F                'Parity/Overflow is not set
    
    'OZ80's keywords
    TOKEN_KEYWORD_AT = &H70
    TOKEN_KEYWORD_AS = &H71
    TOKEN_KEYWORD_BANK = &H72
    TOKEN_KEYWORD_BINARY = &H73
    TOKEN_KEYWORD_BYTE = &H74
    TOKEN_KEYWORD_DATA = &H75
    TOKEN_KEYWORD_DEF = &H76
    TOKEN_KEYWORD_DEFAULT = &H77
    TOKEN_KEYWORD_FILL = &H78
    TOKEN_KEYWORD_INCLUDE = &H79
    TOKEN_KEYWORD_LENGTH = &H7A
    TOKEN_KEYWORD_OBJECT = &H7B
    TOKEN_KEYWORD_PARAMS = &H7C
    TOKEN_KEYWORD_PROC = &H7D
    TOKEN_KEYWORD_RETURN = &H7E
    TOKEN_KEYWORD_SLOT = &H7F
    TOKEN_KEYWORD_START = &H80
    TOKEN_KEYWORD_STOP = &H81
    TOKEN_KEYWORD_STRUCT = &H82
    TOKEN_KEYWORD_WORD = &H83
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H90
    'Number prefixes ("K", "KB" & "Kbit")
    TOKEN_PREFIX_K = &H91               'x1000
    TOKEN_PREFIX_KB = &H92              'x1024
    TOKEN_PREFIX_KBIT = &H93            'x128 (1024 bits)
    
    'Specifies an operator
    TOKEN_OPERATOR_ADD = &HA0           'Add "+"
    TOKEN_OPERATOR_SUB = &HA1           'Subtract "-"
    TOKEN_OPERATOR_MUL = &HA2           'Multiply "*"
    TOKEN_OPERATOR_DIV = &HA3           'Divide "/"
    TOKEN_OPERATOR_POW = &HA4           'Power "^"
    TOKEN_OPERATOR_MOD = &HA5           'Modulus "\"
    TOKEN_OPERATOR_REP = &HA6           'Repeat "x"
    TOKEN_OPERATOR_OR = &HA7            'Bitwise Or "|"
    TOKEN_OPERATOR_AND = &HA8           'Bitwise And "&"
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN = &HB0
    TOKEN_PARENCLOSE = &HB1
    TOKEN_BLOCKOPEN = &HB2
    TOKEN_BLOCKCLOSE = &HB3
    
    'A list is a series of 1 or more expressions separated by commas,
     'i.e.parameter lists
    TOKEN_LIST = &HC0
    TOKEN_NEXT = &HC1                   'Sandwiched between list items, i.e. ","
    
    TOKEN_QUOTE = &HD0
End Enum

'Use a 512 KB initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 524287
'And a 120'000 item buffer for the meta data for each non-data token byte
Private Const META_BUFFER As Long = 120000

'/// DEFS /////////////////////////////////////////////////////////////////////////////

'Please note that the token stream is a combination of tokens (single number codes _
 that act as shorthand for the textual 'words' in the source code), and for some _
 tokens variable length data that follows -- for example a quote token with the length _
 of the quoted text as the next 4 bytes and the quoted text as bytes following that

Private Tokens() As Byte                'The token stream will live here
Private TokensLimit As Long             'The length of the buffer used for above
Private TokensLength As Long            'The current amount of space used in the array

Private Type TokenMeta
    Index As Long
    Line As Long
    Col As Long
End Type

Private Meta() As TokenMeta
'The number of tokens, exluding the data parts. This is what will set the size of the _
 meta array which is one entry for each token (excluding data)
Private TokenCount As Long
Private MetaLimit As Long

'So we can look at which tokens are most common, for optimising
Private TokenCounts(0 To 255) As Long

'A look up table of ASCII codes from upper-case to lower-case
Private LowCase(0 To 255) As Long

'The buffer used to parse the source text file
Private Text() As Byte

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Event Error(ByRef Err As ErrObject)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise the lookup table of upper-case to lower-case values
    Dim ii As Long
    For ii = 0 To 255
        If ii >= 65 And ii <= 90 Then
            Let LowCase(ii) = ii Or 32
        Else
            Let LowCase(ii) = ii
        End If
    Next ii
    
    With CompStrSA1D
        .cDims = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .cbElements = 2&
        .cLocks = 1&
        .lLbound = 0&
    End With
    
    Let CompStrPtr = msvbvm60_VarPtr(CompStrArr())
    Call msvbvm60_PutMem4(CompStrPtr, VarPtr(CompStrSA1D))
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the look up table
    Erase LowCase
    'The `Tokenise` method will clean this up, but just be doubly sure
    Erase Text
    'Empty the token stream
    Erase Tokens
    Erase TokenCounts
    Erase Meta
    
    Call msvbvm60_PutMem4(CompStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Tokenise : Break a file into all its individual words and convert to machine tokens _
 ======================================================================================
Public Function Tokenise(ByRef FilePath As String) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    
    'Measure the time this takes
    Dim Start As Long
    Let Start = kernel32_GetTickCount()
    
    'Just a loop counter
    Dim ii As Long
    
    'Prepare a large buffer (512K) to accept the token stream _
     (using `ReDim` for every token added is *incredibly* slow)
    'If the data overflows this buffer, we'll increase it by _
     another chunk there and then
    Let TokensLimit = TOKEN_BUFFER
    ReDim Tokens(0 To TokensLimit) As Byte
    Let TokensLength = 0
    Erase Text
    
    'Prepare a 120'000 item buffer for the meta data for each non-data token. _
     This contains the line number column number and index number into the token _
     stream where the token occurs. This is so that when errors occur we can easily _
     provide the location in the source code
    Let TokenCount = 0
    Let MetaLimit = META_BUFFER
    ReDim Meta(0 To MetaLimit) As TokenMeta
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    Dim FileNumber As String
    Let FileNumber = FreeFile
    
    'Check the file even exists first...
    If FileExists(FilePath) = False Then
        Let Tokenise = OZ80_ERROR_FILENOTFOUND
        Call Err.Raise( _
            Number:=vbObjectError + OZ80_ERROR_FILENOTFOUND, _
            Source:="oz80Tokeniser", _
            Description:="The file '" & FilePath & "' could not be found." _
        )
        Exit Function
    End If
    
    Debug.Print "* Parsing file: " & FilePath
    'Once the file is open, ensure it's closed upon error
    On Error GoTo ErrorOpenFile
    Open FilePath For Input Access Read Lock Write As #FileNumber
    'Read the file as a binary string with one-byte per character (ASCII). _
     We first normalise to Unicode to deal with locale specifics and then downsample _
     into ASCII to save memory and make stepping through easier
    Let Text = StrConv( _
        StrConv(InputB(LOF(FileNumber), FileNumber), vbUnicode), _
        vbFromUnicode _
    )
    
    'Cache this for speed, we'll be referring to it every loop
    Dim TextLen As Long: Let TextLen = UBound(Text)
    Debug.Print "- Length: " & Format$(TextLen, "#,#") & " bytes"
    
    'Continue error handling normally
    Close #FileNumber
    On Error GoTo 0
    
    Dim IsQuote As Boolean
    Dim IsComment As Boolean
    Dim IsCommentMulti As Boolean
    
    Dim Number As Long, Multiplier As Long
    
    Dim Word As String
    
    'Which character in the source code we're parsing. Since the way the routine is _
     written (we increase the position before each word begins) begin with -1 so that _
     we may really begin at 0
    Dim TextPos As Long
    Let TextPos = -1
    
    'Line and column number in the source text. Used to record the position in the _
     original file of each word for producing helpful error messages
    Dim Line As Long, Col As Long
    Let Line = 1
    
    'Begin parsing words: _
     ----------------------------------------------------------------------------------
ReadWord:
    'If the word has begun (first letter is read)
    Dim IsWord As Long
    Let IsWord = 0
SkipChar:
    'Move on to the next character in the source
    Let TextPos = TextPos + 1
ReadChar:
    'If the file ends?
    If TextPos >= TextLen Then
        'If the word is currently underway then we are all done
        If IsWord = 0 Then GoTo Finish
        'Otherwise process this final word
        GoTo EndWord
    End If
    
    'Read a character:
    Dim AscChar As Long
    Let AscChar = Text(TextPos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the assembling stage)
    If AscChar = aLF Then
        Let Line = Line + 1: Let Col = 0
    ElseIf AscChar = aTAB Then
        'Tab goes to the next 8 char boundary
        'TODO: Could do with a lookup here for speed?
        Let Col = Col + 8 - (Col Mod 8)
    ElseIf AscChar = aCR Then
        GoTo SkipChar
    Else
        Let Col = Col + 1
    End If
    
    'Comments: ........................................................................
    If IsComment Then
        'For single-line comments:
        If IsCommentMulti = False Then
            'Is this a multi-line comment? (two backticks)
            If IsWord Then
                If AscChar = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
            'The end of a line ends a single-line comment: _
             (we only mark as next line on LF so as to avoid double-lines on Windows _
              which uses CR+LF to mark end of line, whilst staying compatible with _
              the Unix-only LF end of line marker)
            If AscChar = aLF Then Let IsComment = False: GoTo ReadWord
        
        'For multi-line comments, check for backtick:
        ElseIf AscChar = SYNTAX_COMMENT Then
            'Check for a second back-tick following
            Let TextPos = TextPos + 1: Let Col = Col + 1
            If Text(TextPos) = SYNTAX_COMMENT Then
                Let IsComment = False: Let IsCommentMulti = False
                GoTo ReadWord
            End If
        End If
        'Skip to the next character
        GoTo SkipChar
    
    'Parsable Text: ...................................................................
    '(false checks are slightly faster)
    ElseIf IsQuote = False Then
        'Special handling for the beginning of a word
        If IsWord = 0 Then
            'Skip whitespace before the word begins
            If AscChar = aSPC Then GoTo SkipChar
            If AscChar = aTAB Then GoTo SkipChar
            If AscChar = aLF Then GoTo SkipChar
            
            'Punctuation characters that are only one letter long. _
             When we hit one of these we need to end the word immediately, but parsing _
             normally ends a word on the letter after, so we deal with the edge case
            If AscChar = SYNTAX_NEXT Then GoTo Punctuation
            If AscChar = SYNTAX_PAREN_OPEN Then GoTo Punctuation
            If AscChar = SYNTAX_PAREN_CLOSE Then GoTo Punctuation
            If AscChar = SYNTAX_BLOCK_OPEN Then GoTo Punctuation
            If AscChar = SYNTAX_BLOCK_CLOSE Then GoTo Punctuation
            
            If AscChar = SYNTAX_OPERATOR_ADD Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_SUB Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_MUL Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_DIV Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_POW Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_MOD Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_OR Then GoTo Punctuation
            If AscChar = SYNTAX_OPERATOR_AND Then GoTo Punctuation
        Else
            'A word is underway, check for end of line (note the `Else`, we don't _
             want to test line endings twice when the Word is currently empty)
            If AscChar = aLF Then
                'The end of a line ends a single-line comment:
                If IsCommentMulti = False Then
                    If IsComment = True Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo EndWord
            End If
            
            'End the word on a space
            If AscChar = aSPC Then GoTo EndWord
            If AscChar = aTAB Then GoTo EndWord
            
            'Symbols that can break a word in the middle without spaces to separate
            If AscChar = SYNTAX_NEXT Then GoTo WordBreak
            If AscChar = SYNTAX_PAREN_OPEN Then GoTo WordBreak
            If AscChar = SYNTAX_PAREN_CLOSE Then GoTo WordBreak
            If AscChar = SYNTAX_BLOCK_OPEN Then GoTo WordBreak
            If AscChar = SYNTAX_BLOCK_CLOSE Then GoTo WordBreak
            
            If AscChar = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_POW Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_OR Then GoTo WordBreak
            If AscChar = SYNTAX_OPERATOR_AND Then GoTo WordBreak
        End If
        
        'Is this a quote? (in which case, ignore wordbreaks until quote end). _
         the opening quote is included so that we can recognise it for tokenising
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = True: GoTo NextChar
        
        'Check for comment marker
        If AscChar = SYNTAX_COMMENT Then
            Let IsComment = True
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If IsWord = 0 Then GoTo NextChar
            GoTo WordBreak
            
        End If
        
    'Quotes: ..........................................................................
    Else
        'TODO: skip tabs and line breaks -- not possible yet due to the memcopy! _
               -- will need to copy into a second byte array? filter array?
        'Skip over the closing quote so that it doesn't get tokenised
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = False: GoTo EndWord
        
    End If
    
NextChar:
    'Is this the first character of the word? If so, record the starting position of _
     the word and flag as word in-progress
    If IsWord = 0 Then
        Dim WordBegin As Long, LineBegin As Long, ColBegin As Long
        Let WordBegin = TextPos: Let LineBegin = Line: Let ColBegin = Col
        Let IsWord = 1
    End If
    GoTo SkipChar

    '----------------------------------------------------------------------------------
    Dim WordEnd As Long, WordLen As Long
    
Punctuation:
    Let WordBegin = TextPos
    Let WordEnd = TextPos
    Let WordLen = 0
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let TextPos = TextPos - 1: Let Col = Col - 1
    Let WordEnd = TextPos
    Let WordLen = (WordEnd - WordBegin) 'Note that this is 0-based, so 0 = 1 char
    GoTo TokenWord
    
EndWord:
    'Remember the end point and length of the word, _
     we can use this to compare bytes (faster than string comparisons)
    
    Let WordEnd = TextPos - 1
    Let WordLen = (WordEnd - WordBegin) 'Note that this is 0-based, so 0 = 1 char

    'Tokenise the word: _
     ----------------------------------------------------------------------------------
TokenWord:
    'Extract the 1st, 2nd, 3rd & 4th letters of the word as these make up most _
     comparisons and can be done quickly without string comparisons
    Dim Asc1 As Long, Asc2 As Long, Asc3 As Long, Asc4 As Long
    Let Asc1 = LowCase(Text(WordBegin))
    Let Asc2 = 0: Let Asc3 = 0: Let Asc4 = 0
    If WordLen > 2 Then Let Asc4 = LowCase(Text(WordBegin + 3))
    If WordLen > 1 Then Let Asc3 = LowCase(Text(WordBegin + 2))
    If WordLen > 0 Then Let Asc2 = LowCase(Text(WordBegin + 1))
    
    GoSub GetWord
    Debug.Print "= " & Word
    
    'Check the first letter of the word to find the right kind of token:
    
    'Comma (very common!)
    If Asc1 = SYNTAX_NEXT Then '............................................... `,` ...
        Call Add(OZ80_TOKEN.TOKEN_NEXT, LineBegin, ColBegin)
        GoTo NextWord
    
    'Hexadecimal number?
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then '..................................... `$` ...
        'TODO: Tokenise number
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_LABEL Then '.......................................... `:` ...
        'A label can be `:` only, without a name (when used within a PROC)
        If Asc2 = 0 Then
            'TODO: Tokenise this
            GoTo NextWord
        
        'A label cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin + 1, WordBegin + WordLen) = False Then
            GoTo ErrorInvalidName
        Else
            'TODO: Tokenise
            GoTo NextWord
        End If
    
    ElseIf Asc1 = SYNTAX_PROPERTY Then '....................................... `.` ...
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        
        'A property can be `.` only, without a name (when used within a PROC)
        If Asc2 = 0 Then
            'TODO: Tokenise this
            GoTo NextWord
        
        'A property cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin + 1, WordBegin + WordLen) = False Then
            GoTo ErrorInvalidName
        Else
            'TODO: Tokenise
            GoTo NextWord
        End If
    
    ElseIf Asc1 = SYNTAX_VARIABLE Then '....................................... `#` ...
        'A variable name cannot be just '#'
        If Asc2 = 0 Then
            GoTo ErrorInvalidName
        
        'A variable cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin + 1, WordBegin + WordLen) = False Then
            GoTo ErrorInvalidName
        Else
            'TODO: Tokenise
            GoTo NextWord
        End If
        
    ElseIf Asc1 = A Then '..................................................... `a` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "adc") Then Call Add(TOKEN_Z80_ADC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "add") Then Call Add(TOKEN_Z80_ADD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "and") Then Call Add(TOKEN_Z80_AND, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "as") Then Call Add(TOKEN_KEYWORD_AS, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "at") Then Call Add(TOKEN_KEYWORD_AT, LineBegin, ColBegin): GoTo NextWord
         
    ElseIf Asc1 = B Then '..................................................... `b` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "bit") Then Call Add(TOKEN_Z80_BIT, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "bank") Then Call Add(TOKEN_KEYWORD_BANK, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "binary") Then Call Add(TOKEN_KEYWORD_BINARY, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "byte") Then Call Add(TOKEN_KEYWORD_BYTE, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = C Then '..................................................... `c` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "call") Then Call Add(TOKEN_Z80_CALL, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ccf") Then Call Add(TOKEN_Z80_CCF, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cp") Then Call Add(TOKEN_Z80_CP, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cpd") Then Call Add(TOKEN_Z80_CPD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cpdr") Then Call Add(TOKEN_Z80_CPDR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cpi") Then Call Add(TOKEN_Z80_CPI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cpir") Then Call Add(TOKEN_Z80_CPIR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "cpl") Then Call Add(TOKEN_Z80_CPL, LineBegin, ColBegin): GoTo NextWord
        'Z80 flag condition
        If Asc2 = 0 Then Call Add(TOKEN_FLAG_C, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = D Then '..................................................... `d` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "daa") Then Call Add(TOKEN_Z80_DAA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "dec") Then Call Add(TOKEN_Z80_DEC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "di") Then Call Add(TOKEN_Z80_DI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "djnz") Then Call Add(TOKEN_Z80_DJNZ, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "data") Then Call Add(TOKEN_KEYWORD_DATA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "def") Then Call Add(TOKEN_KEYWORD_DEF, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "default") Then Call Add(TOKEN_KEYWORD_DEFAULT, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = E Then '..................................................... `e` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "ei") Then Call Add(TOKEN_Z80_EI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ex") Then Call Add(TOKEN_Z80_EX, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "exx") Then Call Add(TOKEN_Z80_EXX, LineBegin, ColBegin): GoTo NextWord
    
    ElseIf Asc1 = F Then '..................................................... `f` ...
        'Keywords
        If CompStr(WordBegin, WordLen, "fill") Then Call Add(TOKEN_KEYWORD_FILL, LineBegin, ColBegin): GoTo NextWord
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = H Then '..................................................... `h` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "halt") Then Call Add(TOKEN_Z80_HALT, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = I Then '..................................................... `i` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "im") Then Call Add(TOKEN_Z80_IM, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "in") Then Call Add(TOKEN_Z80_IN, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "inc") Then Call Add(TOKEN_Z80_INC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ind") Then Call Add(TOKEN_Z80_IND, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "indr") Then Call Add(TOKEN_Z80_INDR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ini") Then Call Add(TOKEN_Z80_INI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "inir") Then Call Add(TOKEN_Z80_INIR, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "include") Then Call Add(TOKEN_KEYWORD_INCLUDE, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = J Then '..................................................... `j` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "jp") Then Call Add(TOKEN_Z80_JP, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "jr") Then Call Add(TOKEN_Z80_JR, LineBegin, ColBegin): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = K Then '..................................................... `k` ...
        'Number suffixes K, KB & Kbit
        If Asc2 = 0 Then    'Just `K`
            Call Add(TOKEN_PREFIX_K, LineBegin, ColBegin): GoTo NextWord
        'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = B Then
            'If only two letters, then it's `KB`
            If Asc3 = 0 Then Call Add(TOKEN_PREFIX_KB, LineBegin, ColBegin): GoTo NextWord
            'verify ...`Kbit`, we don't have Asc5 so avoid things like `Kbitty`
            If CompStr(WordBegin, WordLen, "kbit") Then Call Add(TOKEN_PREFIX_KBIT, LineBegin, ColBegin): GoTo NextWord
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = L Then '..................................................... `l` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "ld") Then Call Add(TOKEN_Z80_LD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ldd") Then Call Add(TOKEN_Z80_LDD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "lddr") Then Call Add(TOKEN_Z80_LDDR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ldi") Then Call Add(TOKEN_Z80_LDI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ldir") Then Call Add(TOKEN_Z80_LDIR, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "length") Then Call Add(TOKEN_KEYWORD_LENGTH, LineBegin, ColBegin): GoTo NextWord
    
    ElseIf Asc1 = M Then '..................................................... `m` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG_M, LineBegin, ColBegin): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = N Then '..................................................... `n` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "neg") Then Call Add(TOKEN_Z80_NEG, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "nop") Then Call Add(TOKEN_Z80_NOP, LineBegin, ColBegin): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = 0 Then
            If Asc2 = C Then Call Add(TOKEN_FLAG_NC, LineBegin, ColBegin)
            If Asc2 = Z Then Call Add(TOKEN_FLAG_NZ, LineBegin, ColBegin)
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = O Then '..................................................... `o` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "or") Then Call Add(TOKEN_Z80_OR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "out") Then Call Add(TOKEN_Z80_OUT, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "outd") Then Call Add(TOKEN_Z80_OUTD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "outdr") Then Call Add(TOKEN_Z80_OUTDR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "outi") Then Call Add(TOKEN_Z80_OUTI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "outir") Then Call Add(TOKEN_Z80_OUTIR, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "object") Then Call Add(TOKEN_KEYWORD_OBJECT, LineBegin, ColBegin): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = P Then '..................................................... `p` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "pop") Then Call Add(TOKEN_Z80_POP, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "push") Then Call Add(TOKEN_Z80_PUSH, LineBegin, ColBegin): GoTo NextWord
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG_P, LineBegin, ColBegin)
        If Asc3 = 0 Then
            If Asc2 = E Then Call Add(TOKEN_FLAG_PE, LineBegin, ColBegin)
            If Asc2 = O Then Call Add(TOKEN_FLAG_PO, LineBegin, ColBegin)
        End If
        'Keywords
        If CompStr(WordBegin, WordLen, "params") Then Call Add(TOKEN_KEYWORD_PARAMS, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "proc") Then Call Add(TOKEN_KEYWORD_PROC, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = R Then '..................................................... `r` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "res") Then Call Add(TOKEN_Z80_RES, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "ret") Then Call Add(TOKEN_Z80_RET, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "reti") Then Call Add(TOKEN_Z80_RETI, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "retn") Then Call Add(TOKEN_Z80_RETN, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rla") Then Call Add(TOKEN_Z80_RLA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rl") Then Call Add(TOKEN_Z80_RL, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rlca") Then Call Add(TOKEN_Z80_RLCA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rlc") Then Call Add(TOKEN_Z80_RLC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rld") Then Call Add(TOKEN_Z80_RLD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rra") Then Call Add(TOKEN_Z80_RRA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rr") Then Call Add(TOKEN_Z80_RR, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rrca") Then Call Add(TOKEN_Z80_RRCA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rrc") Then Call Add(TOKEN_Z80_RRC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rrd") Then Call Add(TOKEN_Z80_RRD, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "rst") Then Call Add(TOKEN_Z80_RST, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "return") Then Call Add(TOKEN_KEYWORD_RETURN, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = S Then '..................................................... `s` ...
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "sbc") Then Call Add(TOKEN_Z80_SBC, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "scf") Then Call Add(TOKEN_Z80_SCF, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "set") Then Call Add(TOKEN_Z80_SET, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "sla") Then Call Add(TOKEN_Z80_SLA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "sra") Then Call Add(TOKEN_Z80_SRA, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "sll") Then Call Add(TOKEN_Z80_SLL, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "srl") Then Call Add(TOKEN_Z80_SRL, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "sub") Then Call Add(TOKEN_Z80_SUB, LineBegin, ColBegin): GoTo NextWord
        'Keywords
        If CompStr(WordBegin, WordLen, "slot") Then Call Add(TOKEN_KEYWORD_SLOT, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "start") Then Call Add(TOKEN_KEYWORD_START, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "stop") Then Call Add(TOKEN_KEYWORD_STOP, LineBegin, ColBegin): GoTo NextWord
        If CompStr(WordBegin, WordLen, "struct") Then Call Add(TOKEN_KEYWORD_STRUCT, LineBegin, ColBegin): GoTo NextWord
        
    ElseIf Asc1 = W Then '..................................................... `w` ...
        'Keywords
        If CompStr(WordBegin, WordLen, "word") Then Call Add(TOKEN_KEYWORD_WORD, LineBegin, ColBegin): GoTo NextWord
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = X Then '..................................................... `x` ...
        'If only `x`, then it's the repeat operator
        If WordLen = 0 Then Call Add(TOKEN_OPERATOR_REP, LineBegin, ColBegin): GoTo NextWord
        'Z80 instructions
        If CompStr(WordBegin, WordLen, "xor") Then Call Add(TOKEN_Z80_XOR, LineBegin, ColBegin): GoTo NextWord
        'There is no `x` register
        GoTo WordError
    
    ElseIf Asc1 = Z Then '..................................................... `z` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG_Z, LineBegin, ColBegin): GoTo NextWord
        'There is no 'z' register
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_BLOCK_OPEN Then '..................................... `{` ...
        Call Add(TOKEN_BLOCKOPEN, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_BLOCK_CLOSE Then  '................................... `}` ...
        Call Add(TOKEN_BLOCKCLOSE, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_OPEN Then '..................................... `(` ...
        Call Add(TOKEN_PARENOPEN, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_CLOSE Then '.................................... `)` ...
        Call Add(TOKEN_PARENCLOSE, LineBegin, ColBegin): GoTo NextWord
    
    'A decimal number?
    ElseIf Asc1 >= NUM0 And Asc1 <= NUM9 Then '.............................. `0-9` ...
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordBegin + WordLen To WordBegin Step -1
            If Text(ii) < NUM0 Or Text(ii) > NUM9 Then
                'TODO: Error, invalid number
                Stop
            Else
                'Convert the number from ASCII codes into a value
                'TODO: Handle an overflow!!!
                'TODO: Optimise this into LUTs
                Let Number = Number + (Text(ii) - NUM0) * Multiplier
                Let Multiplier = Multiplier * 10
            End If
        Next ii
        'Add the token to signal a number
        Call Add(TOKEN_NUMBER, LineBegin, ColBegin)
        'Copy the long directly into the token stream
        Call kernel32_RtlMoveMemory( _
            Tokens(TokensLength), Number, 4 _
        )
        Let TokensLength = TokensLength + 4
        GoTo NextWord
        
    'Binary number?
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then '..................................... `%` ...
        'TODO: Tokenise number
        GoTo NextWord
        
    'String (quote)
    ElseIf Asc1 = SYNTAX_QUOTE Then '.......................................... `"` ...
        'Add quote token, the quote text follows
        Call Add(TOKEN_QUOTE, LineBegin, ColBegin)
        'Copy the bytes from the quote directly into the token stream
        Call kernel32_RtlMoveMemory( _
            Tokens(TokensLength), Text(WordBegin + 1), WordLen - 1 _
        )
        Let TokensLength = TokensLength + WordLen
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_OPERATOR_ADD Then '................................... `+` ...
        Call Add(TOKEN_OPERATOR_ADD, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '................................... `-` ...
        Call Add(TOKEN_OPERATOR_SUB, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '................................... `*` ...
        Call Add(TOKEN_OPERATOR_MUL, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '................................... `/` ...
        Call Add(TOKEN_OPERATOR_DIV, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '................................... `^` ...
        Call Add(TOKEN_OPERATOR_POW, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '................................... `\` ...
        Call Add(TOKEN_OPERATOR_MOD, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_OR Then  '................................... `|` ...
        Call Add(TOKEN_OPERATOR_OR, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_AND Then '................................... `&` ...
        Call Add(TOKEN_OPERATOR_AND, LineBegin, ColBegin): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < A Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > S Then GoTo WordError
    'If there's only 1 letter total, then it's obviously a 1-letter register
    If Asc2 = 0 Then GoTo OneLetter
    'If a semi-colon occurs at the second letter, then the register must be 1 letter
    If Asc2 = SYNTAX_HINT1 Then GoTo OneLetter
    'Also an apostrophe (shadow register) is the same
    If Asc2 = SYNTAX_HINT2 Then GoTo OneLetter
    
    'If there's only 2 letters in total, it has to be a two-letter register
    If Asc3 = 0 Then GoTo TwoLetters
    'If a semi-colon or apostrophe occurs at the third letter, _
     it could be a two-letter register
    If Asc3 = SYNTAX_HINT1 Then GoTo TwoLetters
    If Asc3 = SYNTAX_HINT2 Then GoTo TwoLetters
    
    GoTo WordError

OneLetter:
    If Asc1 = A Then Call Add(TOKEN_REGISTER_A, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = B Then Call Add(TOKEN_REGISTER_B, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = C Then Call Add(TOKEN_REGISTER_C, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = D Then Call Add(TOKEN_REGISTER_D, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = E Then Call Add(TOKEN_REGISTER_E, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = H Then Call Add(TOKEN_REGISTER_H, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = I Then Call Add(TOKEN_REGISTER_I, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = L Then Call Add(TOKEN_REGISTER_L, LineBegin, ColBegin): GoTo NextWord
    If Asc1 = R Then Call Add(TOKEN_REGISTER_R, LineBegin, ColBegin): GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = A Then       '"af"
        If Asc2 = F Then Call Add(TOKEN_REGISTER_AF, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = B Then   '"bc"
        If Asc2 = C Then Call Add(TOKEN_REGISTER_BC, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = D Then  '"de"
        If Asc2 = E Then Call Add(TOKEN_REGISTER_DE, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = H Then  '"hl"
        If Asc2 = L Then Call Add(TOKEN_REGISTER_HL, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = I Then  '"ix"/"iy"
        If Asc2 = X Then Call Add(TOKEN_REGISTER_IX, LineBegin, ColBegin): GoTo NextWord
        If Asc2 = Y Then Call Add(TOKEN_REGISTER_IY, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = P Then  '"pc"
        If Asc2 = C Then Call Add(TOKEN_REGISTER_PC, LineBegin, ColBegin): GoTo NextWord
    ElseIf Asc1 = S Then  '"sp"
        If Asc2 = P Then Call Add(TOKEN_REGISTER_SP, LineBegin, ColBegin): GoTo NextWord
    End If
    
NextWord:
    Dim WordCount As Long
    Let WordCount = WordCount + 1
    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
GetWord:
    'Build a string (for use in error messages) of the word from the byte array
    Let Word = vbNullString
    For ii = WordBegin To WordBegin + WordLen
        Let Word = Word & Chr$(Text(ii))
    Next ii
    Return

    '----------------------------------------------------------------------------------
ErrorOpenFile:
    'If an error occured during file reading, ensure the file handle is closed
    On Error GoTo 0
    Close #FileNumber
    Let Tokenise = OZ80_ERROR_FILEREAD
    Call RaiseError( _
        OZ80_ERROR_FILEREAD, _
        "A problem occured when reading the file '" & FilePath & "', " & _
        "the system says: " & vbCrLf & vbCrLf & _
        "Err #" & Err.Number & ": " & Err.Description _
    )
    GoTo Finish

ErrorInvalidName:
    On Error GoTo 0
    GoSub GetWord
    Let Tokenise = OZ80_ERROR_INVALIDNAME
    Call RaiseError( _
        OZ80_ERROR_INVALIDNAME, _
        "Failure parsing '" & Word & "'; " & _
        "Variable, label and property names can contain A-Z, 0-9 underscore and dot " & _
        "with the exceptions that the first letter cannot be a number or a dot, " & _
        "two dots cannot occur in a row nor can a number follow a dot and the name " & _
        "cannot end in a dot", LineBegin, ColBegin _
    )
    GoTo Finish
    
WordError:
    GoSub GetWord
    Let Tokenise = OZ80_ERROR_BADWORD
    Call RaiseError( _
        OZ80_ERROR_BADWORD, _
        "Faliure parsing '" & Word & "'; " & _
        "Not a known keyword, Z80 instruction, register or other known word. " & _
        "Typo, perhaps?", LineBegin, ColBegin _
    )
    GoTo Finish
    
ErrorOut:
Finish:
    If Tokenise = OZ80_ERROR_NONE Then
       'Reduce the buffers to the final size
        ReDim Preserve Tokens(0 To TokensLength - 1) As Byte
        ReDim Preserve Meta(0 To TokenCount - 1) As TokenMeta
        
        Erase Text
'       Debug.Print "- Tokens: "
'       Debug.Print "  NEXT: " & Format$(TokenCounts(OZ80_TOKEN.TOKEN_NEXT), "#,#")
'       Debug.Print "  (): " & Format$(TokenCounts(OZ80_TOKEN.TOKEN_PARENOPEN) + TokenCounts(OZ80_TOKEN.TOKEN_PARENCLOSE), "#,#")
'       Debug.Print "  NUMBER: " & Format$(TokenCounts(OZ80_TOKEN.TOKEN_NUMBER), "#,#")
'       Debug.Print "  {}: " & Format$(TokenCounts(OZ80_TOKEN.TOKEN_BLOCKOPEN) + TokenCounts(OZ80_TOKEN.TOKEN_BLOCKCLOSE), "#,#")
        Debug.Print "- Words: " & Format$(WordCount, "#,#")
        Debug.Print "- Time: " & Format$((kernel32_GetTickCount() - Start) / 1000, "0.000") & "s"
    Else
        Erase Text, Tokens, Meta
        Let TokensLength = 0
        Let TokenCount = 0
    End If
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Add : Add a token (any) to the stream _
 ======================================================================================
Private Sub Add(ByRef Token As OZ80_TOKEN, ByRef Line As Long, ByRef Col As Long)
    'The param MUST be 0-255, but is given as long for speed reasons; _
     the rest of this class will ensure that this is never out of bounds
    Let Tokens(TokensLength) = Token
    
    'Collect info on the number of each different type of token so we can get some _
     statistics for possible speed optimisations
    Let TokenCounts(Token) = TokenCounts(Token) + 1
    
    'Capture the token's position in the source file and token stream
    With Meta(TokenCount)
        Let .Index = TokensLength - 1
        Let .Line = Line
        Let .Col = Col
    End With
    
    'Move to the next token number
    Let TokenCount = TokenCount + 1
    'Increase the buffer of token information if needed
    If TokenCount > MetaLimit Then
        Let MetaLimit = MetaLimit + META_BUFFER
        ReDim Preserve Meta(0 To MetaLimit) As TokenMeta
    End If
    
    'Move to the next space in the byte stream
    Let TokensLength = TokensLength + 1
    
    'Note that the tokens array is given a large size to begin with as ReDimming it _
     with every token added is *incredibly* slow. If we go over this buffer size, _
     we need to increase it
    If TokensLength > TokensLimit Then
        Let TokensLimit = TokensLimit + TOKEN_BUFFER
        ReDim Preserve Tokens(0 To TokensLimit) As Byte
    End If
End Sub

'CompStr : Compare a word from the byte array to a string _
 ======================================================================================
Private Function CompStr( _
    ByRef Begin As Long, ByRef Length As Long, _
    ByRef CompText As String _
) As Boolean
    'WARNING: `Text` is assumed to be lower case!
    
    'This function is designed as a quick'n'dirty replacement for `StrComp`, _
     but using the byte array as the source so we no longer have to concatenate _
     strings in the parsing routine
    
    'If the lengths don't match then they can't be the same
    If Length <> Len(CompText) Then Exit Function

    With CompStrSA1D
       .pvData = StrPtr(CompText)
       .cElements = Len(CompText)
    End With
    
    Dim ii As Long
    For ii = 0 To Length - 1
        If LowCase(Text(Begin + ii)) <> CompStrArr(ii) Then Exit Function
    Next ii
    
    Let CompStr = True
End Function

'FileExists : See if a file exists or not _
 ======================================================================================
'<cuinl.tripod.com/Tips/fileexist.htm>
Private Function FileExists(ByRef Path As String) As Boolean
    Let FileExists = CBool(Dir$(Path) <> vbNullString)
End Function

'IsValidName : Check validity for variables, labels & properties _
 ======================================================================================
Public Function IsValidName(ByRef First As Long, ByRef Last As Long) As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    
    'NOTE: This function assumes that you have skipped over the prefix (`#`/`:`/`.`)
    
    'The name cannot begin with a number or a dot
    If Text(First) = SYNTAX_PROPERTY Then Exit Function
    If Text(First) >= NUM0 And Text(First) <= NUM9 Then Exit Function
    'Cannot end in a dot
    If Text(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim ii As Long
    For ii = First To Last
        'Not a dot?
        If Text(ii) <> SYNTAX_PROPERTY Then
            'Is letter not a number?
            If Text(ii) < NUM0 Or Text(ii) > NUM9 Then
                'Is letter outside of A-Z (lower case and upper case)?
                If Text(ii) < A Or Text(ii) > Z Then
                    If Text(ii) < 65 Or Text(ii) > 90 Then
                        'Finally, not an underscore?
                        If Text(ii) <> UNDRSCR Then Exit Function
                    End If
                End If
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If Text(ii + 1) >= NUM0 And Text(ii + 1) <= NUM9 Then Exit Function
            'Two dots in a row not allowed
            If Text(ii + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next ii
    
    Let IsValidName = True
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Sub RaiseError( _
    ByRef Number As OZ80_ERROR, _
    ByRef Description As String, _
    Optional ByRef Line As Long = 0, Optional ByRef Col As Long = 0 _
)
    Debug.Print
    Debug.Print "! ERROR: #" & Number
    If Line > 0 And Col > 0 Then
        Debug.Print "- Line: " & Format$(Line, "#,#") & " Col: " & Col
    End If
    Debug.Print "- " & Description
    Debug.Print
    
    Call Err.Clear
    Let Err.Number = vbObjectError + Number
    Let Err.Description = Description
    Let Err.Source = "Line #" & Line & " Col #" & Col
    
    RaiseEvent Error(Err)
End Sub
