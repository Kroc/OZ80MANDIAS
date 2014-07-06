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

'/// API //////////////////////////////////////////////////////////////////////////////

'Cast String to Array: _
 --------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB string directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for VB arrays _
 <msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx>
Private Type SAFEARRAY1D
    cDims      As Integer   'The count of dimensions
    fFeatures  As Integer   'Flags used by the SafeArray
    cbElements As Long      'The size of an array element
    cLocks     As Long      'Lock count
    pvData     As Long      'Pointer to the data
    cElements  As Long      'The number of elements in the dimension
    lLbound    As Long      'The lower bound of the dimension
End Type

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

'Undocumented VB6 API to get at the array pointer
Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Integer         'Empty array that will be hacked

'/// PUBLIC CONSTANTS /////////////////////////////////////////////////////////////////
'These are defined here (instead of in 'oz80_Assembler') to avoid circular references _
 due to the way the assembler sinks events from this class

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_FILENOTFOUND             'Requested file does not exist
    OZ80_ERROR_FILEREAD                 'Some kind of problem with file handle open
    OZ80_ERROR_INVALIDNAME              'Invalid label/property/variable name
    OZ80_ERROR_BADWORD                  'Couldn't parse a word
    OZ80_ERROR_BADNUMBER_DEC            'Not a valid decimal number
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_Z80_PARAMETER            'An unexpected parameter for a z80 instruction
    OZ80_ERROR_OPERAND                  'Not a valid operand for an expression
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_UNEXPECTED               'Incorrect content at the current scope
    OZ80_ERROR_ENDOFFILE                'Unexpected end of file
    OZ80_ERROR_INDEFINITEVALUE          'Indefinite value cannot be used here
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_TOKEN
    TOKEN_Z80 = &H1                     'Z80 instruction
    TOKEN_REGISTER = &H2                'Z80 register
    TOKEN_FLAG = &H2                    'Z80 flag condition (used on JP, CALL &  RET)
                                         '(part of registers, due to shared "C")
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H10
    'Number prefixes ("K", "KB" & "Kbit")
    TOKEN_PREFIX_K = &H11               'x1000
    TOKEN_PREFIX_KB = &H12              'x1024
    TOKEN_PREFIX_KBIT = &H13            'x128 (1024 bits)
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN = &H20
    TOKEN_PARENCLOSE = &H21
    TOKEN_BLOCKOPEN = &H22
    TOKEN_BLOCKCLOSE = &H23
    
    TOKEN_OPERATOR = &H30               'Operator (e.g. "+ - * /")
    
    TOKEN_KEYWORD = &H41                'Keyword (IF/DATA/ECHO &c.)
    TOKEN_QUOTE = &H42
    TOKEN_LABEL = &H43
    TOKEN_PROPERTY = &H44
    TOKEN_VARIABLE = &H45
End Enum

Public Enum OZ80_TOKEN_DATA
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
    TOKEN_Z80_OTDR = &H28               'Output, Decrement and Repeat
    TOKEN_Z80_OUTI = &H29               'Output and Increment
    TOKEN_Z80_OTIR = &H2A               'Output, Increment and Repeat
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
    
    'Z80 Registers ....................................................................
    TOKEN_REGISTER_A = &H1              'Accumulator
    TOKEN_REGISTER_F = &H2              'Flags register
    TOKEN_REGISTER_B = &H4
    TOKEN_REGISTER_C = &H8
    TOKEN_REGISTER_D = &H10
    TOKEN_REGISTER_E = &H20
    TOKEN_REGISTER_H = &H40
    TOKEN_REGISTER_L = &H80
    
    TOKEN_REGISTER_I = &H100            'Interrupt - not to be confused with IX & IY
    TOKEN_REGISTER_R = &H101            'Refresh register (pseudo-random)
    
    TOKEN_REGISTER_SP = &H104           'Stack pointer
    TOKEN_REGISTER_PC = &H108           'Program counter
    
    TOKEN_REGISTER_AF = TOKEN_REGISTER_A Or TOKEN_REGISTER_F
    TOKEN_REGISTER_BC = TOKEN_REGISTER_B Or TOKEN_REGISTER_C
    TOKEN_REGISTER_DE = TOKEN_REGISTER_D Or TOKEN_REGISTER_E
    TOKEN_REGISTER_HL = TOKEN_REGISTER_H Or TOKEN_REGISTER_L
    
    'Undocumented Z80 instructions can access the 8-bit halves of IX & IY
    TOKEN_REGISTER_IXL = &H201
    TOKEN_REGISTER_IXH = &H202
    TOKEN_REGISTER_IYL = &H211
    TOKEN_REGISTER_IYH = &H212
    TOKEN_REGISTER_IX = TOKEN_REGISTER_IXL Or TOKEN_REGISTER_IXH
    TOKEN_REGISTER_IY = TOKEN_REGISTER_IYL Or TOKEN_REGISTER_IYH
    
    'Z80 Flag Conditions ..............................................................
    'The flags share the same space as the registers since they share the "C"
     'register/flag and it's not possible to deterime which is implied early on
    TOKEN_FLAG_C = &H8                  'Carry set
    TOKEN_FLAG_NC = &H301               'Carry not set
    TOKEN_FLAG_Z = &H302                'Zero set
    TOKEN_FLAG_NZ = &H303               'Zero not set
    TOKEN_FLAG_M = &H304                'Sign is set
    TOKEN_FLAG_P = &H305                'Sign is not set
    TOKEN_FLAG_PE = &H306               'Parity/Overflow is set
    TOKEN_FLAG_PO = &H307               'Parity/Overflow is not set
    
    'Operators ........................................................................
    TOKEN_OPERATOR_ADD = &H1            'Add "+"
    TOKEN_OPERATOR_SUB = &H2            'Subtract "-"
    TOKEN_OPERATOR_MUL = &H3            'Multiply "*"
    TOKEN_OPERATOR_DIV = &H4            'Divide "/"
    TOKEN_OPERATOR_POW = &H5            'Power "^"
    TOKEN_OPERATOR_MOD = &H6            'Modulus "\"
    TOKEN_OPERATOR_REP = &H7            'Repeat "x"
    TOKEN_OPERATOR_OR = &H8             'Bitwise Or "|"
    TOKEN_OPERATOR_AND = &H9            'Bitwise And "&"
    
    'Keywords .........................................................................
    TOKEN_KEYWORD_AT = &H1
    TOKEN_KEYWORD_AS = &H2
    TOKEN_KEYWORD_BANK = &H3
    TOKEN_KEYWORD_BINARY = &H4
    TOKEN_KEYWORD_BYTE = &H5
    TOKEN_KEYWORD_DATA = &H6
    TOKEN_KEYWORD_DEFAULT = &H7
    TOKEN_KEYWORD_ECHO = &H8
    TOKEN_KEYWORD_ELSE = &H9
    TOKEN_KEYWORD_EXISTS = &HA
    TOKEN_KEYWORD_FAIL = &HB
    TOKEN_KEYWORD_FILL = &HC
    TOKEN_KEYWORD_IF = &HD
    TOKEN_KEYWORD_INCLUDE = &HE
    TOKEN_KEYWORD_LENGTH = &HF
    TOKEN_KEYWORD_OBJECT = &H10
    TOKEN_KEYWORD_PARAMS = &H11
    TOKEN_KEYWORD_PROC = &H12
    TOKEN_KEYWORD_RETURN = &H13
    TOKEN_KEYWORD_SECTION = &H14
    TOKEN_KEYWORD_SLOT = &H15
    TOKEN_KEYWORD_START = &H16
    TOKEN_KEYWORD_STOP = &H17
    TOKEN_KEYWORD_STRUCT = &H18
    TOKEN_KEYWORD_TABLE = &H19
    TOKEN_KEYWORD_VAR = &H1A
    TOKEN_KEYWORD_WORD = &H1B
End Enum

'/// PRIVATE DEFS /////////////////////////////////////////////////////////////////////

'This makes life a whole lot easier
Private Enum ASCII
    UNDRSCR = 95
    
    NUM0 = 48
    NUM1 = 49
    NUM9 = 57
    
    a = 97:     B = 98:     C = 99:     D = 100:    E = 101:    F = 102
    G = 103:    H = 104:    i = 105:    J = 106:    K = 107:    L = 108
    M = 109:    N = 110:    O = 111:    p = 112:    q = 113:    R = 114
    S = 115:    T = 116:    U = 117:    V = 118:    W = 119:    x = 120
    y = 121:    z = 122
    
    aSPC = 32
    aTAB = 9
    aCR = 10
    aLF = 13
End Enum

'--------------------------------------------------------------------------------------

'These define the various punctiation marks (in ASCII codes) for the language syntax
Private Enum SYNTAX
    SYNTAX_COMMENT = 96         ' ` - Comment marker. "``" for multi-line comment
    SYNTAX_HINT1 = 59           ' ; - register hint, e.g. `a;index`
    SYNTAX_HINT2 = 39           ' ' - shadow register hint, e.g. `ex af 'af`
    SYNTAX_QUOTE = 34           ' " - string identifier
    SYNTAX_LABEL = 58           ' : - label identifier
    SYNTAX_PROPERTY = 46        ' . - property identifier
    SYNTAX_VARIABLE = 35        ' # - variable identifier
    SYNTAX_MACRO = 64           ' @ - macro identifier
    SYNTAX_FUNCT = 63           ' ? - function identifier
    SYNTAX_NUMBER_HEX = 36      ' $ - hexadecimal number, e.g. `$FFFF`
    SYNTAX_NUMBER_BIN = 37      ' % - binary number, e.g. `%10101011`
    SYNTAX_NEXT = 44            ' , - item seperator, optional
    SYNTAX_PAREN_OPEN = 40      ' (
    SYNTAX_PAREN_CLOSE = 41     ' )
    SYNTAX_BLOCK_OPEN = 123     ' {
    SYNTAX_BLOCK_CLOSE = 125    ' }
    SYNTAX_OPERATOR_ADD = 43    ' +
    SYNTAX_OPERATOR_SUB = 45    ' -
    SYNTAX_OPERATOR_MUL = 42    ' *
    SYNTAX_OPERATOR_DIV = 47    ' /
    SYNTAX_OPERATOR_POW = 94    ' ^
    SYNTAX_OPERATOR_MOD = 92    ' \
    SYNTAX_OPERATOR_OR = 124    ' |
    SYNTAX_OPERATOR_AND = 38    ' &
End Enum

'Use a 100'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 100000

'Used for time measurement
Private Stopwatch As PerformanceCounter

'Look-up tables to validate ASCII codes:
'TODO: Add an IsValid array to test for unprintable ASCII codes like null, bell &c.
Private LowCase(0 To 255) As Long       'From upper case to lower case
Private IsNumber(0 To 255) As Boolean   'Validates decimal digits
Private IsHex(0 To 255) As Boolean      'Validates hexadecimal digits (0-9, A-F)
Private IsAlpha(0 To 255) As Boolean    'Validate alpha characters (both cases)
Private IsAlphaNum(0 To 255) As Boolean 'Validate alphanumeric chars (both cases)

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'Object to calculate the CRC of the file during load _
 (will be used by the assembler to determine if a file has already been passed)
Private CRC32 As oz80_CRC32

'The source text file is held here as a byte-array whilst it gets parsed
Private Text() As Byte

Private Tokens() As oz80Token           'The token stream will live here
Private TokenLimit As Long              'The length of the buffer used for the above
Private TokenCount As Long              'The current amount of space used in the array
Private TokenIndex As Long              'The current place within the tokenStream

Private Type oz80Token
    Kind As Byte                        '=OZ80_TOKEN, but use 1-byte instead of 4
    Data As OZ80_TOKEN_DATA             'Associated value
    Line As Long                        'Line number in the original source text
    Col As Long                         'Column number in the original source text
End Type

'Remember where the original source file is as the assembler will need to change to _
 its directory when handling relative references with includes
Public SourceFile As String

'When parsing source text, we can save having to pass these through function calls _
 hundreds of times by making them global to this class
Private WordBegin As Long, WordLen As Long
Private LineBegin As Long, ColBegin As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message(ByRef Text As String)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)

'We raise an event each time we come across a string so that the assembler can pool _
 together all strings across all source files
Event StringAdded(ByRef Key As Long, ByRef StringData As String)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Set Stopwatch = New PerformanceCounter
    
    'Save time doing and/or comparisons by compiling look up tables of ASCII _
     codes that count as numeric / hex / alpha and alphanumeric
    Dim ii As Long
    For ii = 0 To 255
        If ii >= NUM0 And ii <= NUM9 Then
            Let LowCase(ii) = ii
            Let IsNumber(ii) = True
            Let IsHex(ii) = True
            Let IsAlphaNum(ii) = True
        ElseIf ii >= a And ii <= z Then
            Let LowCase(ii) = ii
            Let IsAlpha(ii) = True
            Let IsAlphaNum(ii) = True
            If ii <= F Then Let IsHex(ii) = True
        ElseIf ii >= 65 And ii <= 90 Then
            Let LowCase(ii) = ii Or 32
            Let IsAlpha(ii) = True
            Let IsAlphaNum(ii) = True
            If ii <= 70 Then Let IsHex(ii) = True
        Else
            Let LowCase(ii) = ii
        End If
    Next ii
    
    'Initialise the array used within the `CompStr` function
    With CastStrSA1D
        .cDims = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .cbElements = 2&
        .cLocks = 1&
        .lLbound = 0&
    End With
    
    Let CastStrPtr = msvbvm60_VarPtr(CastStrArr())
    Call msvbvm60_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the timing class
    Set Stopwatch = Nothing
    'Clean up the look up tables
    Erase LowCase, IsNumber, IsHex, IsAlphaNum
    'The `Tokenise` method will clean this up, but just be doubly sure
    Erase Text
    'Empty the token stream
    Erase Tokens
    'Clear the original source file path
    Let Me.SourceFile = vbNullString
    
    'Clean up our array hacking (for casting a string to an array) _
     otherwise VB will crash when it frees up the memory
    If CastStrPtr <> 0 Then Call msvbvm60_PutMem4(CastStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Count : Number of tokens _
 ======================================================================================
Public Property Get Count() As Long
    'Internally the token count is zero-based (easy to manage), _
     but externally it is one-based (expected behaviour)
    Let Count = TokenCount + 1
End Property

'PROPERTY EOF : End Of File? _
 ======================================================================================
Public Property Get EOF() As Boolean
    Let EOF = CBool(TokenIndex > TokenCount)
End Property

'PROPERTY Index : Current location within the tokenStream _
 ======================================================================================
Public Property Get Index() As Long: Let Index = TokenIndex: End Property
Public Property Let Index(ByVal Value As Long)
    'Out of bounds?
    If (Value < 0) Or (Value > TokenCount) Then
        'Raise a standard VB error for 'array out of bounds'
        Call Err.Raise(9)
    Else
        Let TokenIndex = Value
    End If
End Property

'PROPERTY Kind : Token Kind of current token _
 ======================================================================================
Public Property Get Kind() As OZ80_TOKEN
    If EOF = True Then Let Kind = 0 Else Let Kind = Tokens(TokenIndex).Kind
End Property

'PROPERTY Data : Data value of current token _
 ======================================================================================
Public Property Get Data() As OZ80_TOKEN_DATA
    If EOF = True Then Let Data = 0 Else Let Data = Tokens(TokenIndex).Data
End Property

'PROPERTY Line : Line number in the source of the current token _
 ======================================================================================
Public Property Get Line() As Long
    If EOF = True Then Let Line = 0 Else Let Line = Tokens(TokenIndex).Line
End Property

'PROPERTY Col : Column number in the source of the current token _
 ======================================================================================
Public Property Get Col() As Long
    If EOF = True Then Let Col = 0 Else Let Col = Tokens(TokenIndex).Col
End Property

'Forward : Move onto the next token, returns false if the TokenStream hit the end _
 ======================================================================================
Public Function Forward() As Boolean
    Let TokenIndex = TokenIndex + 1
    If TokenIndex > TokenCount Then Exit Function
    Let Forward = True
End Function

'Rewind : Go back to the start of the tokenStream _
 ======================================================================================
Public Sub Rewind()
    Let TokenIndex = 0
End Sub

'Tokenise : Break the source file into individual words and convert to machine tokens _
 ======================================================================================
Public Function Tokenise(ByRef FilePath As String) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    
    'Measure the time this takes
    Call Stopwatch.Reset
    
    'Just a loop counter
    Dim ii As Long
    
    'Prepare a large buffer to accept the token stream (using `ReDim` for every token _
     added is *incredibly* slow). If the tokens overflow this buffer, _
     we'll increase it by another chunk there and then
    Let TokenLimit = TOKEN_BUFFER: Let TokenCount = 0: Call Me.Rewind
    ReDim Tokens(0 To TokenLimit) As oz80Token
    Erase Text
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    Dim FileNumber As String
    Let FileNumber = FreeFile
    
    'Check the file even exists first...
    If FileExists(FilePath) = False Then Let Tokenise = RaiseError( _
        OZ80_ERROR_FILENOTFOUND, _
        "File not found", _
        "The file '" & FilePath & "' could not be found." _
    ): GoTo Finish
    
    RaiseEvent Message("* Parsing file: " & FilePath)
    
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
    RaiseEvent Message("- Length: " & Format$(TextLen, "#,#") & " bytes")
    
    'Continue error handling normally
    Close #FileNumber
    On Error GoTo 0
    
    Dim IsQuote As Boolean
    Dim IsComment As Boolean
    Dim IsCommentMulti As Boolean
    
    'These are used whilst converting words into real numbers
    Dim Number As Long, Multiplier As Long
    'This is used to convert the word from a byte array into a string
    Dim Word As String
    
    'Which character in the source code we're parsing. Since the way the routine is _
     written (we increase the position before each word begins) begin with -1 so that _
     we may really begin at 0
    Dim TextPos As Long
    Let TextPos = -1
    
    'Line and column number in the source text. Used to record the position in the _
     original file of each word for producing helpful error messages. 1-based
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
        
    'Strings: .........................................................................
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
        Let WordBegin = TextPos: Let LineBegin = Line: Let ColBegin = Col
        Let IsWord = 1
    End If
    GoTo SkipChar

    '----------------------------------------------------------------------------------
    Dim WordEnd As Long
    
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
    
    'Check the first letter of the word to find the right kind of token:
    
    'Comma (very common!)
    If Asc1 = SYNTAX_NEXT Then '............................................... `,` ...
        'NOTE: At this point we can only see one word at a time (and not look ahead), _
               so we cannot know when a list has started, and even if the grammer is _
               valid at all (could have a number in completely the wrong place)
        GoTo NextWord
    
    'Hexadecimal number?
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then '..................................... `$` ...
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordBegin + WordLen To WordBegin + 1 Step -1
            'Check the digits are 0-9, A-F
            If IsHex(Text(ii)) = False Then
                'ERROR: Invalid hexadecimal number!
                GoSub GetWord
                Let Tokenise = RaiseError(OZ80_ERROR_BADNUMBER_DEC, _
                    "Invalid number", _
                    "Not a valid hexadecimal number! '" & Word & "'", _
                    LineBegin, ColBegin _
                ): GoTo Finish
            End If
        Next ii
        'The digits are valid, but it could be an overflow
        On Error GoTo ErrorOverflow
        'TODO: Test this against another method for speed
        GoSub GetWord
        Let Number = Val("&H" & Mid$(Word, 2) & "&")
        On Error GoTo 0
        'Add the token for the number
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_LABEL Then '.......................................... `:` ...
        'A label can be `:` only, without a name (when used within a PROC)
        If Asc2 = 0 Then
            'This is tokenised as having a label index of 0
            Call Add(TOKEN_LABEL, 0)
            GoTo NextWord
        
        'A label cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin + 1, WordBegin + WordLen) = False Then
            GoTo ErrorInvalidName
        Else
            GoSub GetWord
            Call AddString(TOKEN_LABEL, Word)
            GoTo NextWord
        End If
    
    ElseIf Asc1 = SYNTAX_PROPERTY Then '....................................... `.` ...
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        
        'A property can be `.` only, without a name (when used within a PROC)
        '-- currently undecided about this
        If Asc2 = 0 Then
            GoTo ErrorInvalidName
            GoTo NextWord
        
        'A property cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin + 1, WordBegin + WordLen) = False Then
            GoTo ErrorInvalidName
        Else
            GoSub GetWord
            Call AddString(TOKEN_PROPERTY, Word)
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
            GoSub GetWord
            Call AddString(TOKEN_VARIABLE, Word)
            GoTo NextWord
        End If
        
    ElseIf Asc1 = a Then '..................................................... `a` ...
        'Z80 instructions
        If CompStr("adc") Then Call Add(TOKEN_Z80, TOKEN_Z80_ADC): GoTo NextWord
        If CompStr("add") Then Call Add(TOKEN_Z80, TOKEN_Z80_ADD): GoTo NextWord
        If CompStr("and") Then Call Add(TOKEN_Z80, TOKEN_Z80_AND): GoTo NextWord
        'Keywords
        If CompStr("as") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_AS): GoTo NextWord
        If CompStr("at") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_AT): GoTo NextWord
         
    ElseIf Asc1 = B Then '..................................................... `b` ...
        'Z80 instructions
        If CompStr("bit") Then Call Add(TOKEN_Z80, TOKEN_Z80_BIT): GoTo NextWord
        'Keywords
        If CompStr("bank") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_BANK): GoTo NextWord
        If CompStr("binary") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_BINARY): GoTo NextWord
        If CompStr("byte") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_BYTE): GoTo NextWord
        
    ElseIf Asc1 = C Then '..................................................... `c` ...
        'Z80 instructions
        If CompStr("call") Then Call Add(TOKEN_Z80, TOKEN_Z80_CALL): GoTo NextWord
        If CompStr("ccf") Then Call Add(TOKEN_Z80, TOKEN_Z80_CCF): GoTo NextWord
        If CompStr("cp") Then Call Add(TOKEN_Z80, TOKEN_Z80_CP): GoTo NextWord
        If CompStr("cpd") Then Call Add(TOKEN_Z80, TOKEN_Z80_CPD): GoTo NextWord
        If CompStr("cpdr") Then Call Add(TOKEN_Z80, TOKEN_Z80_CPDR): GoTo NextWord
        If CompStr("cpi") Then Call Add(TOKEN_Z80, TOKEN_Z80_CPI): GoTo NextWord
        If CompStr("cpir") Then Call Add(TOKEN_Z80, TOKEN_Z80_CPIR): GoTo NextWord
        If CompStr("cpl") Then Call Add(TOKEN_Z80, TOKEN_Z80_CPL): GoTo NextWord
        'Z80 flag condition
        If Asc2 = 0 Then Call Add(TOKEN_FLAG, TOKEN_FLAG_C): GoTo NextWord
        
    ElseIf Asc1 = D Then '..................................................... `d` ...
        'Z80 instructions
        If CompStr("daa") Then Call Add(TOKEN_Z80, TOKEN_Z80_DAA): GoTo NextWord
        If CompStr("dec") Then Call Add(TOKEN_Z80, TOKEN_Z80_DEC): GoTo NextWord
        If CompStr("di") Then Call Add(TOKEN_Z80, TOKEN_Z80_DI): GoTo NextWord
        If CompStr("djnz") Then Call Add(TOKEN_Z80, TOKEN_Z80_DJNZ): GoTo NextWord
        'Keywords
        If CompStr("data") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_DATA): GoTo NextWord
        If CompStr("default") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_DEFAULT): GoTo NextWord
        
    ElseIf Asc1 = E Then '..................................................... `e` ...
        'Z80 instructions
        If CompStr("ei") Then Call Add(TOKEN_Z80, TOKEN_Z80_EI): GoTo NextWord
        If CompStr("ex") Then Call Add(TOKEN_Z80, TOKEN_Z80_EX): GoTo NextWord
        If CompStr("exx") Then Call Add(TOKEN_Z80, TOKEN_Z80_EXX): GoTo NextWord
        'Keywords
        If CompStr("echo") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_ECHO): GoTo NextWord
        If CompStr("else") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_ELSE): GoTo NextWord
        If CompStr("exists") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_EXISTS): GoTo NextWord
        
    ElseIf Asc1 = F Then '..................................................... `f` ...
        'Keywords
        If CompStr("fail") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_FAIL): GoTo NextWord
        If CompStr("fill") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_FILL): GoTo NextWord
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = H Then '..................................................... `h` ...
        'Z80 instructions
        If CompStr("halt") Then Call Add(TOKEN_Z80, TOKEN_Z80_HALT): GoTo NextWord
        
    ElseIf Asc1 = i Then '..................................................... `i` ...
        'Z80 instructions
        If CompStr("im") Then Call Add(TOKEN_Z80, TOKEN_Z80_IM): GoTo NextWord
        If CompStr("in") Then Call Add(TOKEN_Z80, TOKEN_Z80_IN): GoTo NextWord
        If CompStr("inc") Then Call Add(TOKEN_Z80, TOKEN_Z80_INC): GoTo NextWord
        If CompStr("ind") Then Call Add(TOKEN_Z80, TOKEN_Z80_IND): GoTo NextWord
        If CompStr("indr") Then Call Add(TOKEN_Z80, TOKEN_Z80_INDR): GoTo NextWord
        If CompStr("ini") Then Call Add(TOKEN_Z80, TOKEN_Z80_INI): GoTo NextWord
        If CompStr("inir") Then Call Add(TOKEN_Z80, TOKEN_Z80_INIR): GoTo NextWord
        'Keywords
        If CompStr("if") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_IF): GoTo NextWord
        If CompStr("include") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_INCLUDE): GoTo NextWord
        
    ElseIf Asc1 = J Then '..................................................... `j` ...
        'Z80 instructions
        If CompStr("jp") Then Call Add(TOKEN_Z80, TOKEN_Z80_JP): GoTo NextWord
        If CompStr("jr") Then Call Add(TOKEN_Z80, TOKEN_Z80_JR): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = K Then '..................................................... `k` ...
        'Number suffixes K, KB & Kbit
        If Asc2 = 0 Then    'Just `K`
            Call Add(TOKEN_PREFIX_K, 0): GoTo NextWord
            'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = B Then
            'If only two letters, then it's `KB`
            If Asc3 = 0 Then Call Add(TOKEN_PREFIX_KB, 0): GoTo NextWord
            'verify ...`Kbit`, we don't have Asc5 so avoid things like `Kbitty`
            If CompStr("kbit") Then Call Add(TOKEN_PREFIX_KBIT, 0): GoTo NextWord
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = L Then '..................................................... `l` ...
        'Z80 instructions
        If CompStr("ld") Then Call Add(TOKEN_Z80, TOKEN_Z80_LD): GoTo NextWord
        If CompStr("ldd") Then Call Add(TOKEN_Z80, TOKEN_Z80_LDD): GoTo NextWord
        If CompStr("lddr") Then Call Add(TOKEN_Z80, TOKEN_Z80_LDDR): GoTo NextWord
        If CompStr("ldi") Then Call Add(TOKEN_Z80, TOKEN_Z80_LDI): GoTo NextWord
        If CompStr("ldir") Then Call Add(TOKEN_Z80, TOKEN_Z80_LDIR): GoTo NextWord
        'Keywords
        If CompStr("length") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_LENGTH): GoTo NextWord
    
    ElseIf Asc1 = M Then '..................................................... `m` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG, TOKEN_FLAG_M): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = N Then '..................................................... `n` ...
        'Z80 instructions
        If CompStr("neg") Then Call Add(TOKEN_Z80, TOKEN_Z80_NEG): GoTo NextWord
        If CompStr("nop") Then Call Add(TOKEN_Z80, TOKEN_Z80_NOP): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = 0 Then
            If Asc2 = C Then Call Add(TOKEN_FLAG, TOKEN_FLAG_NC): GoTo NextWord
            If Asc2 = z Then Call Add(TOKEN_FLAG, TOKEN_FLAG_NZ): GoTo NextWord
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = O Then '..................................................... `o` ...
        'Z80 instructions
        If CompStr("or") Then Call Add(TOKEN_Z80, TOKEN_Z80_OR): GoTo NextWord
        If CompStr("out") Then Call Add(TOKEN_Z80, TOKEN_Z80_OUT): GoTo NextWord
        If CompStr("outd") Then Call Add(TOKEN_Z80, TOKEN_Z80_OUTD): GoTo NextWord
        If CompStr("otdr") Then Call Add(TOKEN_Z80, TOKEN_Z80_OTDR): GoTo NextWord
        If CompStr("outi") Then Call Add(TOKEN_Z80, TOKEN_Z80_OUTI): GoTo NextWord
        If CompStr("otir") Then Call Add(TOKEN_Z80, TOKEN_Z80_OTIR): GoTo NextWord
        'Keywords
        If CompStr("object") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_OBJECT): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = p Then '..................................................... `p` ...
        'Z80 instructions
        If CompStr("pop") Then Call Add(TOKEN_Z80, TOKEN_Z80_POP): GoTo NextWord
        If CompStr("push") Then Call Add(TOKEN_Z80, TOKEN_Z80_PUSH): GoTo NextWord
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG, TOKEN_FLAG_P): GoTo NextWord
        If Asc3 = 0 Then
            If Asc2 = E Then Call Add(TOKEN_FLAG, TOKEN_FLAG_PE): GoTo NextWord
            If Asc2 = O Then Call Add(TOKEN_FLAG, TOKEN_FLAG_PO): GoTo NextWord
        End If
        'Keywords
        If CompStr("params") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If CompStr("proc") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_PROC): GoTo NextWord
        
    ElseIf Asc1 = R Then '..................................................... `r` ...
        'Z80 instructions
        If CompStr("res") Then Call Add(TOKEN_Z80, TOKEN_Z80_RES): GoTo NextWord
        If CompStr("ret") Then Call Add(TOKEN_Z80, TOKEN_Z80_RET): GoTo NextWord
        If CompStr("reti") Then Call Add(TOKEN_Z80, TOKEN_Z80_RETI): GoTo NextWord
        If CompStr("retn") Then Call Add(TOKEN_Z80, TOKEN_Z80_RETN): GoTo NextWord
        If CompStr("rla") Then Call Add(TOKEN_Z80, TOKEN_Z80_RLA): GoTo NextWord
        If CompStr("rl") Then Call Add(TOKEN_Z80, TOKEN_Z80_RL): GoTo NextWord
        If CompStr("rlca") Then Call Add(TOKEN_Z80, TOKEN_Z80_RLCA): GoTo NextWord
        If CompStr("rlc") Then Call Add(TOKEN_Z80, TOKEN_Z80_RLC): GoTo NextWord
        If CompStr("rld") Then Call Add(TOKEN_Z80, TOKEN_Z80_RLD): GoTo NextWord
        If CompStr("rra") Then Call Add(TOKEN_Z80, TOKEN_Z80_RRA): GoTo NextWord
        If CompStr("rr") Then Call Add(TOKEN_Z80, TOKEN_Z80_RR): GoTo NextWord
        If CompStr("rrca") Then Call Add(TOKEN_Z80, TOKEN_Z80_RRCA): GoTo NextWord
        If CompStr("rrc") Then Call Add(TOKEN_Z80, TOKEN_Z80_RRC): GoTo NextWord
        If CompStr("rrd") Then Call Add(TOKEN_Z80, TOKEN_Z80_RRD): GoTo NextWord
        If CompStr("rst") Then Call Add(TOKEN_Z80, TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If CompStr("return") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = S Then '..................................................... `s` ...
        'Z80 instructions
        If CompStr("sbc") Then Call Add(TOKEN_Z80, TOKEN_Z80_SBC): GoTo NextWord
        If CompStr("scf") Then Call Add(TOKEN_Z80, TOKEN_Z80_SCF): GoTo NextWord
        If CompStr("set") Then Call Add(TOKEN_Z80, TOKEN_Z80_SET): GoTo NextWord
        If CompStr("sla") Then Call Add(TOKEN_Z80, TOKEN_Z80_SLA): GoTo NextWord
        If CompStr("sra") Then Call Add(TOKEN_Z80, TOKEN_Z80_SRA): GoTo NextWord
        If CompStr("sll") Then Call Add(TOKEN_Z80, TOKEN_Z80_SLL): GoTo NextWord
        If CompStr("srl") Then Call Add(TOKEN_Z80, TOKEN_Z80_SRL): GoTo NextWord
        If CompStr("sub") Then Call Add(TOKEN_Z80, TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If CompStr("section") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_SECTION): GoTo NextWord
        If CompStr("slot") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_SLOT): GoTo NextWord
        If CompStr("start") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_START): GoTo NextWord
        If CompStr("stop") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_STOP): GoTo NextWord
        If CompStr("struct") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_STRUCT): GoTo NextWord
    
    ElseIf Asc1 = T Then '..................................................... `t` ...
        'Keywords
        If CompStr("table") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_TABLE): GoTo NextWord
        'There is no `t` register
        GoTo WordError
    
    ElseIf Asc1 = V Then '..................................................... `v` ...
        'Keywords
        If CompStr("var") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_VAR): GoTo NextWord
        GoTo WordError
        
    ElseIf Asc1 = W Then '..................................................... `w` ...
        'Keywords
        If CompStr("word") Then Call Add(TOKEN_KEYWORD, TOKEN_KEYWORD_WORD): GoTo NextWord
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = x Then '..................................................... `x` ...
        'If only `x`, then it's the repeat operator
        If WordLen = 0 Then Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_REP): GoTo NextWord
        'Z80 instructions
        If CompStr("xor") Then Call Add(TOKEN_Z80, TOKEN_Z80_XOR): GoTo NextWord
        'There is no `x` register
        GoTo WordError
    
    ElseIf Asc1 = z Then '..................................................... `z` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_FLAG, TOKEN_FLAG_Z): GoTo NextWord
        'There is no 'z' register
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_BLOCK_OPEN Then '..................................... `{` ...
        Call Add(TOKEN_BLOCKOPEN, 0): GoTo NextWord
    ElseIf Asc1 = SYNTAX_BLOCK_CLOSE Then  '................................... `}` ...
        Call Add(TOKEN_BLOCKCLOSE, 0): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_OPEN Then '..................................... `(` ...
        Call Add(TOKEN_PARENOPEN, 0): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_CLOSE Then '.................................... `)` ...
        Call Add(TOKEN_PARENCLOSE, 0): GoTo NextWord
    
    'A decimal number?
    ElseIf IsNumber(Asc1) Then '............................................. `0-9` ...
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordBegin + WordLen To WordBegin Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If IsNumber(Text(ii)) = False Then
                'ERROR: Invalid number!
                GoSub GetWord
                Let Tokenise = RaiseError(OZ80_ERROR_BADNUMBER_DEC, _
                    "Invalid number", _
                    "Not a valid decimal number! '" & Word & "'", _
                    LineBegin, ColBegin _
                ): GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (Text(ii) - NUM0) * Multiplier
                'Don't allow the multiplier to overflow when the number doesn't
                If ii > WordBegin Then Let Multiplier = Multiplier * 10
                On Error GoTo 0
            End If
        Next ii
        'Add the token with the number value
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
        
    'Binary number?
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then '..................................... `%` ...
        'TODO: Cannot be longer than 31-bits
        If WordLen > 31 Then GoTo ErrorOverflow
        Let Number = 0: Let Multiplier = 0
        For ii = WordBegin + WordLen To WordBegin + 1 Step -1
            If Text(ii) = NUM1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf Text(ii) <> NUM0 Then
                'Error - not a valid binary number
                Stop
            End If
            Let Multiplier = Multiplier + 1
        Next ii
        'Tokenise number
        Call Add(TOKEN_NUMBER, Number)
        GoTo NextWord
        
    'String (quote)
    ElseIf Asc1 = SYNTAX_QUOTE Then '.......................................... `"` ...
        GoSub GetWord
        'TODO: Break the string into individual number tokens so it's an inline list
        Call AddString(TOKEN_QUOTE, Mid$(Word, 2))
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_OPERATOR_ADD Then '................................... `+` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_ADD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '................................... `-` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_SUB): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '................................... `*` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_MUL): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '................................... `/` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_DIV): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '................................... `^` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_POW): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '................................... `\` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_MOD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_OR Then  '................................... `|` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_OR): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_AND Then '................................... `&` ...
        Call Add(TOKEN_OPERATOR, TOKEN_OPERATOR_AND): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < a Then GoTo WordError
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
    If Asc1 = a Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_A): GoTo NextWord
    If Asc1 = B Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_B): GoTo NextWord
    If Asc1 = C Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_C): GoTo NextWord
    If Asc1 = D Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_D): GoTo NextWord
    If Asc1 = E Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_E): GoTo NextWord
    If Asc1 = H Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_H): GoTo NextWord
    If Asc1 = i Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_I): GoTo NextWord
    If Asc1 = L Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_L): GoTo NextWord
    If Asc1 = R Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_R): GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = a Then       '"af"
        If Asc2 = F Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_AF): GoTo NextWord
    ElseIf Asc1 = B Then   '"bc"
        If Asc2 = C Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_BC): GoTo NextWord
    ElseIf Asc1 = D Then  '"de"
        If Asc2 = E Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_DE): GoTo NextWord
    ElseIf Asc1 = H Then  '"hl"
        If Asc2 = L Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_HL): GoTo NextWord
    ElseIf Asc1 = i Then  '"ix"/"iy"
        If Asc2 = x Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_IX): GoTo NextWord
        If Asc2 = y Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_IY): GoTo NextWord
    ElseIf Asc1 = p Then  '"pc"
        If Asc2 = C Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_PC): GoTo NextWord
    ElseIf Asc1 = S Then  '"sp"
        If Asc2 = p Then Call Add(TOKEN_REGISTER, TOKEN_REGISTER_SP): GoTo NextWord
    End If
    GoTo WordError
    
    '----------------------------------------------------------------------------------
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
    Let Tokenise = RaiseError(OZ80_ERROR_FILEREAD, _
        "Unknown file system error", _
        "A problem occured when reading the file '" & FilePath & "', " & _
            "the system says: " & vbCrLf & vbCrLf & _
            "Err #" & Err.Number & ": " & Err.Description _
    ): GoTo Finish

ErrorInvalidName: '....................................................................
    On Error GoTo 0
    GoSub GetWord
    Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME, _
        "Invalid item name", _
        "Failure parsing '" & Word & "'; " & _
            "Variable, label and property names can contain A-Z, 0-9 underscore and " & _
            "dot with the exceptions that the first letter cannot be a number or a " & _
            "dot, two dots cannot occur in a row nor can a number follow a dot and " & _
            "the name cannot end in a dot" _
    ): GoTo Finish

ErrorOverflow: '.......................................................................
    On Error GoTo 0
    GoSub GetWord
    Let Tokenise = RaiseError(OZ80_ERROR_OVERFLOW, _
        "Overflow", _
        "Number too large! ('" & Word & "') Numbers are limited to a maximum of " & _
            "31-bits (" & Format$((2 ^ 31) - 1, "#,#") & ")." _
    ): GoTo Finish
    
WordError: '...........................................................................
    GoSub GetWord
    Let Tokenise = RaiseError(OZ80_ERROR_BADWORD, _
        "Invalid content", _
        "Faliure parsing '" & Word & "'; " & _
            "Not a known keyword, Z80 instruction, register or other known word. " & _
            "Typo, perhaps?" _
    ): GoTo Finish
    
    '----------------------------------------------------------------------------------
Finish:
    If Tokenise = OZ80_ERROR_NONE Then
       'Reduce the buffers to the final size
        ReDim Preserve Tokens(0 To TokenCount - 1) As oz80Token
        'Move the cursor back to the start of the stream
        Call Me.Rewind
        
        'When the file parses correctly, we can now expose the original file path. _
         This will be used to resolve relative paths with file includes
        Let Me.SourceFile = FilePath
        
        Erase Text
        RaiseEvent Message("- Words: " & Format$(WordCount, "#,#"))
        RaiseEvent Message("- Tokens: " & Format$(TokenCount, "#,#"))
        RaiseEvent Message("- Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    Else
        'clean up!
        Erase Text, Tokens
        Let TokenCount = 0
        Call Me.Rewind
    End If
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Add : Add a token (any) to the stream _
 ======================================================================================
Private Sub Add( _
    ByRef Kind As OZ80_TOKEN, ByRef Data As OZ80_TOKEN_DATA _
)
    'The `Kind` param MUST be 0-255, but is given as long for speed reasons; _
     the rest of this class will ensure that this is never out of bounds
    With Tokens(TokenIndex)
        Let .Kind = Kind And &HFF
        Let .Data = Data
        'Note that the line and column numbers are global to this class to avoid _
         having to use them in calls hundreds of times over
        Let .Line = LineBegin
        Let .Col = ColBegin
    End With
    
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
End Sub

'AddString : Strings are stored outside of the token stream and referenced by number _
 ======================================================================================
Private Sub AddString( _
    ByRef TokenType As OZ80_TOKEN, ByRef Data As String _
)
    'Note Line & Col numbers are global variables to this class to save having to _
     use them in calls hundreds of times over
    Dim Key As Long
    
    'Pass the string back to the assembler. It'll want to pool together all of the _
     strings used across all files for the purpose of label lookups
    'Note that the assembler will produce the key for us and pass it back
    RaiseEvent StringAdded(Key, Data)
    
    'Add the index key to the token so that we can retrieve the string later
    Call Add(TokenType, Key)
End Sub

'CompStr : Compare a word from the byte array to a string _
 ======================================================================================
Private Function CompStr( _
    ByRef CompText As String _
) As Boolean
    'WARNING: `CompText` is assumed to be lower case!
    
    'This function is designed as a quick'n'dirty replacement for `StrComp`, _
     but using the byte array as the source so we no longer have to concatenate _
     strings in the parsing routine
    
    'If the lengths don't match then they can't be the same. _
     Note that Length is 0-based here
    If WordLen <> Len(CompText) - 1 Then Exit Function

    With CastStrSA1D
       .pvData = StrPtr(CompText)
       .cElements = Len(CompText)
    End With
    
    Dim ii As Long
    For ii = 0 To WordLen
        If LowCase(Text(WordBegin + ii)) <> CastStrArr(ii) Then Exit Function
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
Private Function IsValidName(ByRef First As Long, ByRef Last As Long) As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    
    'NOTE: This function assumes that you have skipped over the prefix (`#`/`:`/`.`)
    
    'The name cannot begin with a number or a dot
    If Text(First) = SYNTAX_PROPERTY Then Exit Function
    If IsNumber(Text(First)) Then Exit Function
    'Cannot end in a dot
    If Text(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim ii As Long
    For ii = First To Last
        'Not a dot?
        If Text(ii) <> SYNTAX_PROPERTY Then
            'Is letter not alphanumeric?
            If IsAlphaNum(Text(ii)) = False Then
                'Finally, not an underscore?
                If Text(ii) <> UNDRSCR Then Exit Function
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If IsNumber(Text(ii + 1)) Then Exit Function
            'Two dots in a row not allowed
            If Text(ii + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next ii
    
    Let IsValidName = True
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Function RaiseError( _
    ByRef Number As OZ80_ERROR, _
    ByRef Title As String, _
    ByRef Description As String, _
    Optional ByRef Line As Long = 0, Optional ByRef Col As Long = 0 _
) As OZ80_ERROR
    'Send the error to the controller (if it sinks events)
    RaiseEvent Error(Number, Title, Description, Line, Col)
    
    'We return the same error number we were given only to make the tail recursion _
     simpler, e.g. `let TokeniseFile = RaiseError(OZ80_ERROR_...`
    Let RaiseError = Number
End Function
