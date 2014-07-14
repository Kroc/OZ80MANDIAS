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

'This makes life a whole lot easier
Private Enum ASCII
    ASC_UNDRSCR = 95
    
    ASC_0 = 48
    ASC_1 = 49
    ASC_9 = 57
    
    ASC_A = 97
    ASC_B
    ASC_C
    ASC_D
    ASC_E
    ASC_F
    ASC_G
    ASC_H
    ASC_I
    ASC_J
    ASC_K
    ASC_L
    ASC_M
    ASC_N
    ASC_O
    ASC_P
    ASC_Q
    ASC_R
    ASC_S
    ASC_T
    ASC_U
    ASC_V
    ASC_W
    ASC_X
    ASC_Y
    ASC_Z
    
    ASC_SPC = 32
    ASC_TAB = 9
    ASC_CR = 10
    ASC_LF = 13
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
    SYNTAX_PAREN_OPEN = 40      ' ( - memory reference open parenthesis
    SYNTAX_PAREN_CLOSE = 41     ' ) - memory reference close parenthesis
    SYNTAX_BLOCK_OPEN = 123     ' { - block open brace
    SYNTAX_BLOCK_CLOSE = 125    ' } - block close brace
    SYNTAX_OPERATOR_ADD = 43    ' + - Add
    SYNTAX_OPERATOR_SUB = 45    ' - - Subtract
    SYNTAX_OPERATOR_MUL = 42    ' * - Multiply
    SYNTAX_OPERATOR_DIV = 47    ' / - Divide
    SYNTAX_OPERATOR_POW = 94    ' ^ - Power
    SYNTAX_OPERATOR_MOD = 92    ' \ - Modulus
    SYNTAX_OPERATOR_OR = 124    ' | - Bitwise OR
    SYNTAX_OPERATOR_AND = 38    ' & - Bitwise AND
    SYNTAX_OPERATOR_NOT = 33    ' ! - Bitwise NOT
End Enum

'Use a 100'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 100000

'Look-up tables to validate ASCII codes:
'TODO: Add an IsValid array to test for unprintable ASCII codes like null, bell &c.
Private LowCase(0 To 255) As Long       'From upper case to lower case
Private IsNumber(0 To 255) As Boolean   'Validates decimal digits
Private IsHex(0 To 255) As Boolean      'Validates hexadecimal digits (0-9, A-F)
Private IsAlpha(0 To 255) As Boolean    'Validate alpha characters (both cases)
Private IsAlphaNum(0 To 255) As Boolean 'Validate alphanumeric chars (both cases)

'Used for time measurement
Private Stopwatch As PerformanceCounter

Private TokenStrings(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'The source text file is held here as a byte-array whilst it gets parsed
Private Text() As Byte

Private Tokens() As oz80Token           'The token stream will live here
Private TokenLimit As Long              'The length of the buffer used for the above
Private TokenCount As Long              'The current amount of space used in the array
Private TokenIndex As Long              'The current place within the tokenStream

Private Type oz80Token
    Kind As Byte                        '=OZ80_TOKEN, but use 1-byte instead of 4
    Value As Long                       'Associated value, e.g. for number tokens
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

'We need to be able to tell the difference between a property being used as a _
 parameter on a Z80 instruction, and a label-property definition (within a block) _
 which occurs on a new line. We keep track of the previous line number used and _
 look for a property occuring on a new line
Private PrevLine As Long

'Look up tables to determine what category a Token belongs to, _
 these are exposed publicly through properties so other classes and do quick checks
Private My_IsInstruction(0 To 255) As Boolean
Private My_IsRegister(0 To 255) As Boolean
Private My_IsRegister8Bit(0 To 255) As Boolean
Private My_IsRegister16Bit(0 To 255) As Boolean
Private My_IsFlag(0 To 255) As Boolean
Private My_IsKeyword(0 To 255) As Boolean
Private My_IsOperator(0 To 255) As Boolean
Private My_IsExpression(0 To 255) As Boolean
Private My_IsParameter(0 To 255) As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message( _
    ByRef Depth As Long, _
    ByRef LogLevel As OZ80_LOG, _
    ByRef Text As String _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, ByVal Line As Long, ByVal Col As Long _
)

'We raise an event each time we come across a string so that the assembler can pool _
 together all strings across all source files
Event StringAdded(ByRef Key As Long, ByRef StringData As String)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Set Stopwatch = New PerformanceCounter
    
    Let TokenStrings(TOKEN_Z80_ADC) = "adc"
    Let TokenStrings(TOKEN_Z80_ADD) = "add"
    Let TokenStrings(TOKEN_Z80_AND) = "and"
    Let TokenStrings(TOKEN_Z80_BIT) = "bit"
    Let TokenStrings(TOKEN_Z80_CALL) = "call"
    Let TokenStrings(TOKEN_Z80_CCF) = "ccf"
    Let TokenStrings(TOKEN_Z80_CP) = "cp"
    Let TokenStrings(TOKEN_Z80_CPD) = "cpd"
    Let TokenStrings(TOKEN_Z80_CPDR) = "cpdr"
    Let TokenStrings(TOKEN_Z80_CPI) = "cpi"
    Let TokenStrings(TOKEN_Z80_CPIR) = "cpir"
    Let TokenStrings(TOKEN_Z80_CPL) = "cpl"
    Let TokenStrings(TOKEN_Z80_DAA) = "daa"
    Let TokenStrings(TOKEN_Z80_DEC) = "dec"
    Let TokenStrings(TOKEN_Z80_DI) = "di"
    Let TokenStrings(TOKEN_Z80_DJNZ) = "djnz"
    Let TokenStrings(TOKEN_Z80_EI) = "ei"
    Let TokenStrings(TOKEN_Z80_EX) = "ex"
    Let TokenStrings(TOKEN_Z80_EXX) = "exx"
    Let TokenStrings(TOKEN_Z80_HALT) = "halt"
    Let TokenStrings(TOKEN_Z80_IM) = "im"
    Let TokenStrings(TOKEN_Z80_IN) = "in"
    Let TokenStrings(TOKEN_Z80_INC) = "inc"
    Let TokenStrings(TOKEN_Z80_IND) = "ind"
    Let TokenStrings(TOKEN_Z80_INDR) = "indr"
    Let TokenStrings(TOKEN_Z80_INI) = "ini"
    Let TokenStrings(TOKEN_Z80_INIR) = "inir"
    Let TokenStrings(TOKEN_Z80_JP) = "jp"
    Let TokenStrings(TOKEN_Z80_JR) = "jr"
    Let TokenStrings(TOKEN_Z80_LD) = "ld"
    Let TokenStrings(TOKEN_Z80_LDD) = "ldd"
    Let TokenStrings(TOKEN_Z80_LDDR) = "lddr"
    Let TokenStrings(TOKEN_Z80_LDI) = "ldi"
    Let TokenStrings(TOKEN_Z80_LDIR) = "ldir"
    Let TokenStrings(TOKEN_Z80_NEG) = "neg"
    Let TokenStrings(TOKEN_Z80_NOP) = "nop"
    Let TokenStrings(TOKEN_Z80_OR) = "or"
    Let TokenStrings(TOKEN_Z80_OUT) = "out"
    Let TokenStrings(TOKEN_Z80_OUTD) = "outd"
    Let TokenStrings(TOKEN_Z80_OTDR) = "otdr"
    Let TokenStrings(TOKEN_Z80_OUTI) = "outi"
    Let TokenStrings(TOKEN_Z80_OTIR) = "otir"
    Let TokenStrings(TOKEN_Z80_POP) = "pop"
    Let TokenStrings(TOKEN_Z80_PUSH) = "push"
    Let TokenStrings(TOKEN_Z80_RES) = "res"
    Let TokenStrings(TOKEN_Z80_RET) = "ret"
    Let TokenStrings(TOKEN_Z80_RETI) = "reti"
    Let TokenStrings(TOKEN_Z80_RETN) = "retn"
    Let TokenStrings(TOKEN_Z80_RLA) = "rla"
    Let TokenStrings(TOKEN_Z80_RL) = "rl"
    Let TokenStrings(TOKEN_Z80_RLCA) = "rlca"
    Let TokenStrings(TOKEN_Z80_RLC) = "rlc"
    Let TokenStrings(TOKEN_Z80_RLD) = "rld"
    Let TokenStrings(TOKEN_Z80_RRA) = "rra"
    Let TokenStrings(TOKEN_Z80_RR) = "rr"
    Let TokenStrings(TOKEN_Z80_RRCA) = "rrca"
    Let TokenStrings(TOKEN_Z80_RRC) = "rrc"
    Let TokenStrings(TOKEN_Z80_RRD) = "rrd"
    Let TokenStrings(TOKEN_Z80_RST) = "rst"
    Let TokenStrings(TOKEN_Z80_SBC) = "sbc"
    Let TokenStrings(TOKEN_Z80_SCF) = "scf"
    Let TokenStrings(TOKEN_Z80_SET) = "set"
    Let TokenStrings(TOKEN_Z80_SLA) = "sla"
    Let TokenStrings(TOKEN_Z80_SRA) = "sra"
    Let TokenStrings(TOKEN_Z80_SLL) = "sll"
    Let TokenStrings(TOKEN_Z80_SRL) = "srl"
    Let TokenStrings(TOKEN_Z80_SUB) = "sub"
    Let TokenStrings(TOKEN_Z80_XOR) = "xor"
    
    'Z80 Registers ....................................................................
    Let TokenStrings(TOKEN_Z80_A) = "a"
    Let TokenStrings(TOKEN_Z80_AF) = "af"
    Let TokenStrings(TOKEN_Z80_B) = "b"
    Let TokenStrings(TOKEN_Z80_C) = "c"
    Let TokenStrings(TOKEN_Z80_NC) = "nc"
    Let TokenStrings(TOKEN_Z80_BC) = "bc"
    Let TokenStrings(TOKEN_Z80_D) = "d"
    Let TokenStrings(TOKEN_Z80_E) = "e"
    Let TokenStrings(TOKEN_Z80_DE) = "de"
    Let TokenStrings(TOKEN_Z80_H) = "h"
    Let TokenStrings(TOKEN_Z80_L) = "l"
    Let TokenStrings(TOKEN_Z80_HL) = "hl"
    Let TokenStrings(TOKEN_Z80_I) = "i"
    Let TokenStrings(TOKEN_Z80_IX) = "ix"
    Let TokenStrings(TOKEN_Z80_IXL) = "ixl"
    Let TokenStrings(TOKEN_Z80_IXH) = "ixh"
    Let TokenStrings(TOKEN_Z80_IY) = "iy"
    Let TokenStrings(TOKEN_Z80_IYL) = "iyl"
    Let TokenStrings(TOKEN_Z80_IYH) = "iyh"
    Let TokenStrings(TOKEN_Z80_M) = "m"
    Let TokenStrings(TOKEN_Z80_P) = "p"
    Let TokenStrings(TOKEN_Z80_PC) = "pc"
    Let TokenStrings(TOKEN_Z80_PE) = "pe"
    Let TokenStrings(TOKEN_Z80_PO) = "po"
    Let TokenStrings(TOKEN_Z80_R) = "r"
    Let TokenStrings(TOKEN_Z80_SP) = "sp"
    Let TokenStrings(TOKEN_Z80_Z) = "Z"
    Let TokenStrings(TOKEN_Z80_NZ) = "nz"
    
    'Operators ........................................................................
    Let TokenStrings(TOKEN_OPERATOR_ADD) = Chr$(SYNTAX.SYNTAX_OPERATOR_ADD)
    Let TokenStrings(TOKEN_OPERATOR_SUB) = Chr$(SYNTAX.SYNTAX_OPERATOR_SUB)
    Let TokenStrings(TOKEN_OPERATOR_MUL) = Chr$(SYNTAX.SYNTAX_OPERATOR_MUL)
    Let TokenStrings(TOKEN_OPERATOR_DIV) = Chr$(SYNTAX.SYNTAX_OPERATOR_DIV)
    Let TokenStrings(TOKEN_OPERATOR_POW) = Chr$(SYNTAX.SYNTAX_OPERATOR_POW)
    Let TokenStrings(TOKEN_OPERATOR_MOD) = Chr$(SYNTAX.SYNTAX_OPERATOR_MOD)
    Let TokenStrings(TOKEN_OPERATOR_REP) = "x"
    Let TokenStrings(TOKEN_OPERATOR_OR) = Chr$(SYNTAX.SYNTAX_OPERATOR_OR)
    Let TokenStrings(TOKEN_OPERATOR_AND) = Chr$(SYNTAX.SYNTAX_OPERATOR_AND)
    Let TokenStrings(TOKEN_OPERATOR_NOT) = Chr$(SYNTAX.SYNTAX_OPERATOR_NOT)
    
    'Keywords .........................................................................
    Let TokenStrings(TOKEN_KEYWORD_AT) = "at"
    Let TokenStrings(TOKEN_KEYWORD_AS) = "as"
    Let TokenStrings(TOKEN_KEYWORD_BANK) = "bank"
    Let TokenStrings(TOKEN_KEYWORD_BINARY) = "binary"
    Let TokenStrings(TOKEN_KEYWORD_BYTE) = "byte"
    Let TokenStrings(TOKEN_KEYWORD_DATA) = "data"
    Let TokenStrings(TOKEN_KEYWORD_DEFAULT) = "default"
    Let TokenStrings(TOKEN_KEYWORD_ECHO) = "echo"
    Let TokenStrings(TOKEN_KEYWORD_ELSE) = "else"
    Let TokenStrings(TOKEN_KEYWORD_EXISTS) = "exists"
    Let TokenStrings(TOKEN_KEYWORD_FAIL) = "fail"
    Let TokenStrings(TOKEN_KEYWORD_FILL) = "fill"
    Let TokenStrings(TOKEN_KEYWORD_IF) = "if"
    Let TokenStrings(TOKEN_KEYWORD_INCLUDE) = "include"
    Let TokenStrings(TOKEN_KEYWORD_LENGTH) = "length"
    Let TokenStrings(TOKEN_KEYWORD_OBJECT) = "object"
    Let TokenStrings(TOKEN_KEYWORD_PARAMS) = "params"
    Let TokenStrings(TOKEN_KEYWORD_PROC) = "proc"
    Let TokenStrings(TOKEN_KEYWORD_RAM) = "ram"
    Let TokenStrings(TOKEN_KEYWORD_RETURN) = "return"
    Let TokenStrings(TOKEN_KEYWORD_SECTION) = "section"
    Let TokenStrings(TOKEN_KEYWORD_SLOT) = "slot"
    Let TokenStrings(TOKEN_KEYWORD_START) = "start"
    Let TokenStrings(TOKEN_KEYWORD_STOP) = "stop"
    Let TokenStrings(TOKEN_KEYWORD_STRUCT) = "struct"
    Let TokenStrings(TOKEN_KEYWORD_TABLE) = "table"
    Let TokenStrings(TOKEN_KEYWORD_VAR) = "var"
    Let TokenStrings(TOKEN_KEYWORD_WORD) = "word"
    
    Let TokenStrings(TOKEN_PREFIX_K) = "k"
    Let TokenStrings(TOKEN_PREFIX_KB) = "kb"
    Let TokenStrings(TOKEN_PREFIX_KBIT) = "kbit"
    
    Let TokenStrings(TOKEN_PARENOPEN) = Chr$(SYNTAX.SYNTAX_PAREN_OPEN)
    Let TokenStrings(TOKEN_PARENCLOSE) = Chr$(SYNTAX.SYNTAX_PAREN_CLOSE)
    Let TokenStrings(TOKEN_BLOCKOPEN) = Chr$(SYNTAX.SYNTAX_BLOCK_OPEN)
    Let TokenStrings(TOKEN_BLOCKCLOSE) = Chr$(SYNTAX.SYNTAX_BLOCK_CLOSE)
    
    Let TokenStrings(TOKEN_QUOTE) = Chr$(SYNTAX.SYNTAX_QUOTE)
    Let TokenStrings(TOKEN_LABEL) = Chr$(SYNTAX.SYNTAX_LABEL)
    Let TokenStrings(TOKEN_PROPERTY_USE) = Chr$(SYNTAX.SYNTAX_PROPERTY)
    Let TokenStrings(TOKEN_PROPERTY_NEW) = Chr$(SYNTAX.SYNTAX_PROPERTY)
    Let TokenStrings(TOKEN_RAM) = Chr$(SYNTAX.SYNTAX_NUMBER_HEX) & Chr$(SYNTAX.SYNTAX_PROPERTY)
    Let TokenStrings(TOKEN_SECTION) = String(2, Chr$(SYNTAX.SYNTAX_LABEL))
    Let TokenStrings(TOKEN_VARIABLE) = Chr$(SYNTAX.SYNTAX_VARIABLE)
    
    '----------------------------------------------------------------------------------
    
    'Save time doing and/or comparisons by compiling look up tables of ASCII _
     codes that count as numeric / hex / alpha and alphanumeric
    Dim ii As OZ80_TOKEN
    For ii = 0 To 255
        'Is this a number?
        If ii >= ASC_0 And ii <= ASC_9 Then
            Let LowCase(ii) = ii        'Is already lower case
            Let IsNumber(ii) = True     'Is a valid number digit
            Let IsHex(ii) = True        'Is a valid hexadecimal digit
            Let IsAlphaNum(ii) = True   'Is alphanumeric
        
        'Is this lowercase a-z?
        ElseIf ii >= ASC_A And ii <= ASC_Z Then
            Let LowCase(ii) = ii        'Is already lower case
            Let IsAlpha(ii) = True      'Is alpha
            Let IsAlphaNum(ii) = True   'Is alphanumeric too
            'If a-f then also hexadecimal
            If ii <= ASC_F Then Let IsHex(ii) = True
        
        'Is uppercase A-Z?
        ElseIf ii >= 65 And ii <= 90 Then
            Let LowCase(ii) = ii Or 32  'Convert to lowercase
            Let IsAlpha(ii) = True      'Is alpha
            Let IsAlphaNum(ii) = True   'Is alphanumeric
            'If A-F then also hexadecimal
            If ii <= 70 Then Let IsHex(ii) = True
        
        'No special case
        Else
            Let LowCase(ii) = ii        'Is already lowercase
        End If
        
        'Using this number as a Token number, check what token category it belongs to
        Let My_IsInstruction(ii) = (ii > [_TOKEN_INSTRUCTIONS_BEGIN]) And _
                                   (ii < [_TOKEN_INSTRUCTIONS_END])
        Let My_IsOperator(ii) = (ii > [_TOKEN_OPERATORS_BEGIN]) And _
                                (ii < [_TOKEN_OPERATORS_END])
        Let My_IsKeyword(ii) = (ii > [_TOKEN_KEYWORDS_BEGIN]) And _
                               (ii < [_TOKEN_KEYWORDS_END])
        
        'Registers and flags are more complex as there's some overlap
        If (ii = TOKEN_Z80_C) Then
            'At the token level, we can't differentiate Register C and Flag C; _
             we need context for that and we can't gaurantee the Token Stream is _
             valid until we begin assembling, therefore we treat Register C and _
             Flag C as the same and the assembler will sort out the difference
            Let My_IsRegister8Bit(ii) = True
            Let My_IsRegister(ii) = True
            Let My_IsFlag(ii) = True
            Let My_IsParameter(ii) = True
        
        '8-Bit registers:
        ElseIf (ii = TOKEN_Z80_A) Or (ii = TOKEN_Z80_B) Or (ii = TOKEN_Z80_C) Or _
               (ii = TOKEN_Z80_D) Or (ii = TOKEN_Z80_E) Or (ii = TOKEN_Z80_H) Or _
               (ii = TOKEN_Z80_L) Or (ii = TOKEN_Z80_I) Or (ii = TOKEN_Z80_R) Or _
               (ii = TOKEN_Z80_IXL) Or (ii = TOKEN_Z80_IXH) Or _
               (ii = TOKEN_Z80_IYL) Or (ii = TOKEN_Z80_IYH) _
        Then
            Let My_IsRegister8Bit(ii) = True
            Let My_IsRegister(ii) = True
            Let My_IsParameter(ii) = True
        
        '16-Bit registers:
        ElseIf (ii = TOKEN_Z80_AF) Or (ii = TOKEN_Z80_BC) Or (ii = TOKEN_Z80_DE) Or _
               (ii = TOKEN_Z80_HL) Or (ii = TOKEN_Z80_IX) Or (ii = TOKEN_Z80_IY) Or _
               (ii = TOKEN_Z80_SP) Or (ii = TOKEN_Z80_PC) _
        Then
            Let My_IsRegister16Bit(ii) = True
            Let My_IsRegister(ii) = True
            Let My_IsParameter(ii) = True
        
        'Condition flags:
        ElseIf (ii = TOKEN_Z80_NC) Or (ii = TOKEN_Z80_M) Or (ii = TOKEN_Z80_P) Or _
               (ii = TOKEN_Z80_PE) Or (ii = TOKEN_Z80_PO) Or (ii = TOKEN_Z80_Z) Or _
               (ii = TOKEN_Z80_NZ) _
        Then
            Let My_IsFlag(ii) = True
            Let My_IsParameter(ii) = True
        
        'If the token consititutes an expression (a calculable value)
        ElseIf (ii = TOKEN_LABEL) Or (ii = TOKEN_NUMBER) Or _
               (ii = TOKEN_PROPERTY_USE) Or (ii = TOKEN_VARIABLE) _
        Then
            Let My_IsExpression(ii) = True
            Let My_IsParameter(ii) = True
        
        'Z80 parameters are expressions, with the addition of Z80 register / memory _
         references, e.g. `(ix+$FF)`
        ElseIf (ii = TOKEN_PARENOPEN) Then
            Let My_IsParameter(ii) = True
        End If
    Next ii
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the timing class
    Set Stopwatch = Nothing
    'Clean up the look up tables
    Erase LowCase, IsNumber, IsHex, IsAlpha, IsAlphaNum
    Erase My_IsInstruction, My_IsRegister, My_IsRegister8Bit, My_IsRegister16Bit, _
          My_IsFlag, My_IsKeyword, My_IsOperator, My_IsExpression, My_IsParameter
    'The `Tokenise` method will clean this up, but just be doubly sure
    Erase Text
    'Empty the token stream
    Erase Tokens
    'Clear the original source file path
    Let Me.SourceFile = vbNullString
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Col : Column number in the source of the current token _
 ======================================================================================
Public Property Get Col() As Long
    If EOF = True Then Let Col = 0 Else Let Col = Tokens(TokenIndex).Col
End Property

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

'PROPERTY IsExpression : Is this token number considered a calculable value? _
 ======================================================================================
Public Property Get IsExpression() As Boolean
    Let IsExpression = My_IsExpression(Me.Kind)
End Property

'PROPERTY IsFlag : Is this token number a condition flag? _
 ======================================================================================
Public Property Get IsFlag() As Boolean
    Let IsFlag = My_IsFlag(Me.Kind)
End Property

'PROPERTY IsInstruction : Is this token number a Z80 instruction? _
 ======================================================================================
Public Property Get IsInstruction() As Boolean
    Let IsInstruction = My_IsInstruction(Me.Kind)
End Property

'PROPERTY IsKeyword : Is this token number a OZ80 Keyword? _
 ======================================================================================
Public Property Get IsKeyword() As Boolean
    Let IsKeyword = My_IsKeyword(Me.Kind)
End Property

'PROPERTY IsOperator : Is this token number an OZ80 operator? _
 ======================================================================================
Public Property Get IsOperator() As Boolean
    Let IsOperator = My_IsOperator(Me.Kind)
End Property

'PROPERTY IsPrameter : Is this token number a Z80 instruction parameter? _
 ======================================================================================
Public Property Get IsParameter() As Boolean
    Let IsParameter = My_IsParameter(Me.Kind)
End Property

'PROPERTY IsRegister : Is this token number a Register (8 or 16-Bit)? _
 ======================================================================================
Public Property Get IsRegister() As Boolean
    Let IsRegister = My_IsRegister(Me.Kind)
End Property

'PROPERTY IsRegister8Bit : Is this token number an 8-bit Register? _
 ======================================================================================
Public Property Get IsRegister8Bit() As Boolean
    Let IsRegister8Bit = My_IsRegister8Bit(Me.Kind)
End Property

'PROPERTY IsRegister16Bit : Is this token number a 16-bit Register? _
 ======================================================================================
Public Property Get IsRegister16Bit() As Boolean
    Let IsRegister16Bit = My_IsRegister16Bit(Me.Kind)
End Property

'PROPERTY Kind : Token Kind of current token _
 ======================================================================================
Public Property Get Kind() As OZ80_TOKEN
    If EOF = True Then Let Kind = 0 Else Let Kind = Tokens(TokenIndex).Kind
End Property

'PROPERTY Line : Line number in the source of the current token _
 ======================================================================================
Public Property Get Line() As Long
    If EOF = True Then Let Line = 0 Else Let Line = Tokens(TokenIndex).Line
End Property

'PROPERTY TokenString : Get the textual representation of a token (for logging) _
 ======================================================================================
Public Property Get TokenString(ByRef Token As OZ80_TOKEN) As String
    Let TokenString = TokenStrings(Token)
End Property

'PROPERTY Value : Token Value of current token _
 ======================================================================================
Public Property Get Value() As Long
    If EOF = True Then Let Value = 0 Else Let Value = Tokens(TokenIndex).Value
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
        OZ80_ERROR_FILENOTFOUND _
    ): GoTo Finish
    
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
    
    'Continue error handling normally
    Close #FileNumber
    On Error GoTo 0
    
    RaiseEvent Message(1, OZ80_LOG_ACTION, _
        "Tokenising " & Chr(34) & FilePath & Chr(34) _
    )
    RaiseEvent Message(2, OZ80_LOG_INFO, _
        "Length: " & Format$(TextLen, "#,#") & " bytes" _
    )
    
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
    Let Line = 1: Let PrevLine = 1
    
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
    Dim AscChar As SYNTAX
    Let AscChar = Text(TextPos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the assembling stage)
    If AscChar = ASC_LF Then
        Let Line = Line + 1: Let Col = 0
    ElseIf AscChar = ASC_TAB Then
        'Tab goes to the next 8 char boundary
        'TODO: Could do with a lookup here for speed?
        Let Col = Col + 8 - (Col Mod 8)
    ElseIf AscChar = ASC_CR Then
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
            If AscChar = ASC_LF Then Let IsComment = False: GoTo ReadWord
        
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
            If AscChar = ASC_SPC Then GoTo SkipChar
            If AscChar = ASC_TAB Then GoTo SkipChar
            If AscChar = ASC_LF Then GoTo SkipChar
            
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
            If AscChar = SYNTAX_OPERATOR_NOT Then GoTo Punctuation
        Else
            'A word is underway, check for end of line (note the `Else`, we don't _
             want to test line endings twice when the Word is currently empty)
            If AscChar = ASC_LF Then
                'The end of a line ends a single-line comment:
                If IsCommentMulti = False Then
                    If IsComment = True Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo EndWord
            End If
            
            'End the word on a space
            If AscChar = ASC_SPC Then GoTo EndWord
            If AscChar = ASC_TAB Then GoTo EndWord
            
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
            If AscChar = SYNTAX_OPERATOR_NOT Then GoTo WordBreak
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
    
    'RAM name, e.g. `$.gamevar`
    ElseIf Asc1 = SYNTAX_NUMBER_HEX _
       And Asc2 = SYNTAX_PROPERTY Then '...................................... `$.` ...
        'ERROR: Invalid RAM name?
        If Not IsValidName(WordBegin, WordBegin + WordLen) Then
            Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME_RAM)
            GoTo Finish
        End If
        
        'TODO: Define RAM constant
        
        GoTo NextWord
        
    'Hexadecimal number
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then '..................................... `$` ...
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordBegin + WordLen To WordBegin + 1 Step -1
            'Check the digits are 0-9, A-F
            'ERROR: Invalid hexadecimal number!
            If IsHex(Text(ii)) = False Then Let Tokenise = RaiseError( _
                OZ80_ERROR_INVALIDNUMBER_HEX _
            ): GoTo Finish
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
            Call Add(TOKEN_LABEL)
            GoTo NextWord
        
        'A label cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf Not IsValidName(WordBegin, WordBegin + WordLen) Then
            'TODO: Label specific error message
            Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME)
            GoTo Finish
        Else
            GoSub GetWord
            'Is this a section name (`::section`) or regular label (`:label`)?
            If Asc2 = SYNTAX_LABEL Then
                Call AddString(TOKEN_SECTION, Word)
            Else
                Call AddString(TOKEN_LABEL, Word)
            End If
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
            'TODO: Property specific error message
            Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME)
            GoTo Finish
            GoTo NextWord
        
        'A property cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf IsValidName(WordBegin, WordBegin + WordLen) = False Then
            'TODO: Property specific error message
            Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME)
            GoTo Finish
        Else
            GoSub GetWord
            'If the property appears on a new line, it's a label-property definition _
             i.e. it defines a new sub-label of a block
            If LineBegin <> PrevLine Then
                Call AddString(TOKEN_PROPERTY_NEW, Word)
            Else
                Call AddString(TOKEN_PROPERTY_USE, Word)
            End If
            GoTo NextWord
        End If
    
    ElseIf Asc1 = SYNTAX_VARIABLE Then '....................................... `#` ...
        'A variable name cannot be just '#'
        'A variable cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        If (Asc2 = 0) Or (Not IsValidName(WordBegin, WordBegin + WordLen)) Then
            'TODO: Variable specific error message
            Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNAME)
            GoTo Finish
        Else
            GoSub GetWord
            Call AddString(TOKEN_VARIABLE, Word)
            GoTo NextWord
        End If
        
    ElseIf Asc1 = ASC_A Then '................................................. `a` ...
        'Z80 instructions
        If CompStr("adc") Then Call Add(TOKEN_Z80_ADC): GoTo NextWord
        If CompStr("add") Then Call Add(TOKEN_Z80_ADD): GoTo NextWord
        If CompStr("and") Then Call Add(TOKEN_Z80_AND): GoTo NextWord
        'Keywords
        If CompStr("as") Then Call Add(TOKEN_KEYWORD_AS): GoTo NextWord
        If CompStr("at") Then Call Add(TOKEN_KEYWORD_AT): GoTo NextWord
         
    ElseIf Asc1 = ASC_B Then '................................................. `b` ...
        'Z80 instructions
        If CompStr("bit") Then Call Add(TOKEN_Z80_BIT): GoTo NextWord
        'Keywords
        If CompStr("bank") Then Call Add(TOKEN_KEYWORD_BANK): GoTo NextWord
        If CompStr("binary") Then Call Add(TOKEN_KEYWORD_BINARY): GoTo NextWord
        If CompStr("byte") Then Call Add(TOKEN_KEYWORD_BYTE): GoTo NextWord
        
    ElseIf Asc1 = ASC_C Then '................................................. `c` ...
        'Z80 instructions
        If CompStr("call") Then Call Add(TOKEN_Z80_CALL): GoTo NextWord
        If CompStr("ccf") Then Call Add(TOKEN_Z80_CCF): GoTo NextWord
        If CompStr("cp") Then Call Add(TOKEN_Z80_CP): GoTo NextWord
        If CompStr("cpd") Then Call Add(TOKEN_Z80_CPD): GoTo NextWord
        If CompStr("cpdr") Then Call Add(TOKEN_Z80_CPDR): GoTo NextWord
        If CompStr("cpi") Then Call Add(TOKEN_Z80_CPI): GoTo NextWord
        If CompStr("cpir") Then Call Add(TOKEN_Z80_CPIR): GoTo NextWord
        If CompStr("cpl") Then Call Add(TOKEN_Z80_CPL): GoTo NextWord
        'Z80 flag condition
        If Asc2 = 0 Then Call Add(TOKEN_Z80_C): GoTo NextWord
        
    ElseIf Asc1 = ASC_D Then '................................................. `d` ...
        'Z80 instructions
        If CompStr("daa") Then Call Add(TOKEN_Z80_DAA): GoTo NextWord
        If CompStr("dec") Then Call Add(TOKEN_Z80_DEC): GoTo NextWord
        If CompStr("di") Then Call Add(TOKEN_Z80_DI): GoTo NextWord
        If CompStr("djnz") Then Call Add(TOKEN_Z80_DJNZ): GoTo NextWord
        'Keywords
        If CompStr("data") Then Call Add(TOKEN_KEYWORD_DATA): GoTo NextWord
        If CompStr("default") Then Call Add(TOKEN_KEYWORD_DEFAULT): GoTo NextWord
        
    ElseIf Asc1 = ASC_E Then '................................................. `e` ...
        'Z80 instructions
        If CompStr("ei") Then Call Add(TOKEN_Z80_EI): GoTo NextWord
        If CompStr("ex") Then Call Add(TOKEN_Z80_EX): GoTo NextWord
        If CompStr("exx") Then Call Add(TOKEN_Z80_EXX): GoTo NextWord
        'Keywords
        If CompStr("echo") Then Call Add(TOKEN_KEYWORD_ECHO): GoTo NextWord
        If CompStr("else") Then Call Add(TOKEN_KEYWORD_ELSE): GoTo NextWord
        If CompStr("exists") Then Call Add(TOKEN_KEYWORD_EXISTS): GoTo NextWord
        
    ElseIf Asc1 = ASC_F Then '................................................. `f` ...
        'Keywords
        If CompStr("fail") Then Call Add(TOKEN_KEYWORD_FAIL): GoTo NextWord
        If CompStr("fill") Then Call Add(TOKEN_KEYWORD_FILL): GoTo NextWord
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = ASC_G Then '................................................. `g` ...
        'There's nothing that begins with G
        GoTo WordError
        
    ElseIf Asc1 = ASC_H Then '................................................. `h` ...
        'Z80 instructions
        If CompStr("halt") Then Call Add(TOKEN_Z80_HALT): GoTo NextWord
        
    ElseIf Asc1 = ASC_I Then '................................................. `i` ...
        'Z80 instructions
        If CompStr("im") Then Call Add(TOKEN_Z80_IM): GoTo NextWord
        If CompStr("in") Then Call Add(TOKEN_Z80_IN): GoTo NextWord
        If CompStr("inc") Then Call Add(TOKEN_Z80_INC): GoTo NextWord
        If CompStr("ind") Then Call Add(TOKEN_Z80_IND): GoTo NextWord
        If CompStr("indr") Then Call Add(TOKEN_Z80_INDR): GoTo NextWord
        If CompStr("ini") Then Call Add(TOKEN_Z80_INI): GoTo NextWord
        If CompStr("inir") Then Call Add(TOKEN_Z80_INIR): GoTo NextWord
        'Keywords
        If CompStr("if") Then Call Add(TOKEN_KEYWORD_IF): GoTo NextWord
        If CompStr("include") Then Call Add(TOKEN_KEYWORD_INCLUDE): GoTo NextWord
        
    ElseIf Asc1 = ASC_J Then '................................................. `j` ...
        'Z80 instructions
        If CompStr("jp") Then Call Add(TOKEN_Z80_JP): GoTo NextWord
        If CompStr("jr") Then Call Add(TOKEN_Z80_JR): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_K Then '................................................. `k` ...
        'Number suffixes K, KB & Kbit
        If Asc2 = 0 Then    'Just `K`
            Call Add(TOKEN_PREFIX_K): GoTo NextWord
            'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = ASC_B Then
            'If only two letters, then it's `KB`
            If Asc3 = 0 Then Call Add(TOKEN_PREFIX_KB): GoTo NextWord
            'verify ...`Kbit`, we don't have Asc5 so avoid things like `Kbitty`
            If CompStr("kbit") Then Call Add(TOKEN_PREFIX_KBIT): GoTo NextWord
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_L Then '................................................. `l` ...
        'Z80 instructions
        If CompStr("ld") Then Call Add(TOKEN_Z80_LD): GoTo NextWord
        If CompStr("ldd") Then Call Add(TOKEN_Z80_LDD): GoTo NextWord
        If CompStr("lddr") Then Call Add(TOKEN_Z80_LDDR): GoTo NextWord
        If CompStr("ldi") Then Call Add(TOKEN_Z80_LDI): GoTo NextWord
        If CompStr("ldir") Then Call Add(TOKEN_Z80_LDIR): GoTo NextWord
        'Keywords
        If CompStr("length") Then Call Add(TOKEN_KEYWORD_LENGTH): GoTo NextWord
    
    ElseIf Asc1 = ASC_M Then '................................................. `m` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_Z80_M): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_N Then '................................................. `n` ...
        'Z80 instructions
        If CompStr("neg") Then Call Add(TOKEN_Z80_NEG): GoTo NextWord
        If CompStr("nop") Then Call Add(TOKEN_Z80_NOP): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = 0 Then
            If Asc2 = ASC_C Then Call Add(TOKEN_Z80_NC): GoTo NextWord
            If Asc2 = ASC_Z Then Call Add(TOKEN_Z80_NZ): GoTo NextWord
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_O Then '................................................. `o` ...
        'Z80 instructions
        If CompStr("or") Then Call Add(TOKEN_Z80_OR): GoTo NextWord
        If CompStr("out") Then Call Add(TOKEN_Z80_OUT): GoTo NextWord
        If CompStr("outd") Then Call Add(TOKEN_Z80_OUTD): GoTo NextWord
        If CompStr("otdr") Then Call Add(TOKEN_Z80_OTDR): GoTo NextWord
        If CompStr("outi") Then Call Add(TOKEN_Z80_OUTI): GoTo NextWord
        If CompStr("otir") Then Call Add(TOKEN_Z80_OTIR): GoTo NextWord
        'Keywords
        If CompStr("object") Then Call Add(TOKEN_KEYWORD_OBJECT): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_P Then '................................................. `p` ...
        'Z80 instructions
        If CompStr("pop") Then Call Add(TOKEN_Z80_POP): GoTo NextWord
        If CompStr("push") Then Call Add(TOKEN_Z80_PUSH): GoTo NextWord
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_Z80_P): GoTo NextWord
        If Asc3 = 0 Then
            If Asc2 = ASC_E Then Call Add(TOKEN_Z80_PE): GoTo NextWord
            If Asc2 = ASC_O Then Call Add(TOKEN_Z80_PO): GoTo NextWord
        End If
        'Keywords
        If CompStr("params") Then Call Add(TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If CompStr("proc") Then Call Add(TOKEN_KEYWORD_PROC): GoTo NextWord
    
    ElseIf Asc1 = ASC_Q Then '................................................. `q` ...
        'There's nothing that begins with Q
        GoTo WordError
        
    ElseIf Asc1 = ASC_R Then '................................................. `r` ...
        'Z80 instructions
        If CompStr("res") Then Call Add(TOKEN_Z80_RES): GoTo NextWord
        If CompStr("ret") Then Call Add(TOKEN_Z80_RET): GoTo NextWord
        If CompStr("reti") Then Call Add(TOKEN_Z80_RETI): GoTo NextWord
        If CompStr("retn") Then Call Add(TOKEN_Z80_RETN): GoTo NextWord
        If CompStr("rla") Then Call Add(TOKEN_Z80_RLA): GoTo NextWord
        If CompStr("rl") Then Call Add(TOKEN_Z80_RL): GoTo NextWord
        If CompStr("rlca") Then Call Add(TOKEN_Z80_RLCA): GoTo NextWord
        If CompStr("rlc") Then Call Add(TOKEN_Z80_RLC): GoTo NextWord
        If CompStr("rld") Then Call Add(TOKEN_Z80_RLD): GoTo NextWord
        If CompStr("rra") Then Call Add(TOKEN_Z80_RRA): GoTo NextWord
        If CompStr("rr") Then Call Add(TOKEN_Z80_RR): GoTo NextWord
        If CompStr("rrca") Then Call Add(TOKEN_Z80_RRCA): GoTo NextWord
        If CompStr("rrc") Then Call Add(TOKEN_Z80_RRC): GoTo NextWord
        If CompStr("rrd") Then Call Add(TOKEN_Z80_RRD): GoTo NextWord
        If CompStr("rst") Then Call Add(TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If CompStr("return") Then Call Add(TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = ASC_S Then '................................................. `s` ...
        'Z80 instructions
        If CompStr("sbc") Then Call Add(TOKEN_Z80_SBC): GoTo NextWord
        If CompStr("scf") Then Call Add(TOKEN_Z80_SCF): GoTo NextWord
        If CompStr("set") Then Call Add(TOKEN_Z80_SET): GoTo NextWord
        If CompStr("sla") Then Call Add(TOKEN_Z80_SLA): GoTo NextWord
        If CompStr("sra") Then Call Add(TOKEN_Z80_SRA): GoTo NextWord
        If CompStr("sll") Then Call Add(TOKEN_Z80_SLL): GoTo NextWord
        If CompStr("srl") Then Call Add(TOKEN_Z80_SRL): GoTo NextWord
        If CompStr("sub") Then Call Add(TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If CompStr("section") Then Call Add(TOKEN_KEYWORD_SECTION): GoTo NextWord
        If CompStr("slot") Then Call Add(TOKEN_KEYWORD_SLOT): GoTo NextWord
        If CompStr("start") Then Call Add(TOKEN_KEYWORD_START): GoTo NextWord
        If CompStr("stop") Then Call Add(TOKEN_KEYWORD_STOP): GoTo NextWord
        If CompStr("struct") Then Call Add(TOKEN_KEYWORD_STRUCT): GoTo NextWord
    
    ElseIf Asc1 = ASC_T Then '................................................. `t` ...
        'Keywords
        If CompStr("table") Then Call Add(TOKEN_KEYWORD_TABLE): GoTo NextWord
        'There is no `t` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_U Then '................................................. `u` ...
        'There's nothing that begins with U
        GoTo WordError
    
    ElseIf Asc1 = ASC_V Then '................................................. `v` ...
        'Keywords
        If CompStr("var") Then Call Add(TOKEN_KEYWORD_VAR): GoTo NextWord
        GoTo WordError
        
    ElseIf Asc1 = ASC_W Then '................................................. `w` ...
        'Keywords
        If CompStr("word") Then Call Add(TOKEN_KEYWORD_WORD): GoTo NextWord
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_X Then '................................................. `x` ...
        'If only `x`, then it's the repeat operator
        If WordLen = 0 Then Call Add(TOKEN_OPERATOR_REP): GoTo NextWord
        'Z80 instructions
        If CompStr("xor") Then Call Add(TOKEN_Z80_XOR): GoTo NextWord
        'There is no `x` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_Y Then '................................................. `y` ...
        'There's nothing that begins with Y
        GoTo WordError
    
    ElseIf Asc1 = ASC_Z Then '................................................. `z` ...
        'Z80 flag conditions
        If Asc2 = 0 Then Call Add(TOKEN_Z80_Z): GoTo NextWord
        'There is no 'z' register
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_BLOCK_OPEN Then '..................................... `{` ...
        Call Add(TOKEN_BLOCKOPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_BLOCK_CLOSE Then  '................................... `}` ...
        Call Add(TOKEN_BLOCKCLOSE): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_OPEN Then '..................................... `(` ...
        Call Add(TOKEN_PARENOPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_CLOSE Then '.................................... `)` ...
        Call Add(TOKEN_PARENCLOSE): GoTo NextWord
    
    'A decimal number?
    ElseIf IsNumber(Asc1) Then '............................................. `0-9` ...
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordBegin + WordLen To WordBegin Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If IsNumber(Text(ii)) = False Then
                'ERROR: Invalid number!
                Let Tokenise = RaiseError(OZ80_ERROR_INVALIDNUMBER_DEC)
                GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (Text(ii) - ASC_0) * Multiplier
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
            If Text(ii) = ASC_1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf Text(ii) <> ASC_0 Then
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
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < ASC_A Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > ASC_S Then GoTo WordError
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
    'TODO: Send the system error details, Err.Number / Err.Description
    Let Tokenise = RaiseError(OZ80_ERROR_FILEREAD)
    GoTo Finish

ErrorOverflow: '.......................................................................
    On Error GoTo 0
    Let Tokenise = RaiseError(OZ80_ERROR_OVERFLOW)
    GoTo Finish
    
WordError: '...........................................................................
    On Error GoTo 0
    Let Tokenise = RaiseError(OZ80_ERROR_INVALIDWORD)
    GoTo Finish
    
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
        RaiseEvent Message(2, OZ80_LOG_INFO, "Words  " & Format$(WordCount, "#,#"))
        RaiseEvent Message(2, OZ80_LOG_INFO, "Tokens " & Format$(TokenCount, "#,#"))
        RaiseEvent Message(2, OZ80_LOG_INFO, "Time   " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
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
             ByRef Token As OZ80_TOKEN, _
    Optional ByRef Value As Long = 0 _
)
    'The `Kind` param MUST be 0-255, but is given as long for speed reasons; _
     the rest of this class will ensure that this is never out of bounds
    With Tokens(TokenIndex)
        Let .Kind = Token And &HFF
        Let .Value = Value
        'Note that the line and column numbers are global to this class to avoid _
         having to use them in calls hundreds of times over
        Let .Line = LineBegin
        Let .Col = ColBegin
    End With
    
    'Generate a debug message. We don't show string-based tokens (such as labels) _
     since we can't access the string database in the Assembler from here. Those _
     tokens are displayed in `AddString` below which has the text to hand
    Dim Msg As String
    If My_IsRegister(Token) Then
        Let Msg = "[REGISTER] " & TokenStrings(Token)
    ElseIf Token = TOKEN_NUMBER Then
        Let Msg = "[  NUMBER] " & Format$(Value, "#,#")
    ElseIf My_IsKeyword(Token) Then
        Let Msg = "[ KEYWORD] " & UCase(TokenStrings(Token))
    ElseIf My_IsOperator(Token) Then
        Let Msg = "[OPERATOR] " & TokenStrings(Token)
    ElseIf Token = TOKEN_BLOCKOPEN Or Token = TOKEN_BLOCKCLOSE Then
        Let Msg = "[   BLOCK] " & TokenStrings(Token)
    ElseIf My_IsInstruction(Token) Then
        Let Msg = "[     Z80] " & TokenStrings(Token)
    ElseIf Token = TOKEN_PARENOPEN Or Token = TOKEN_PARENCLOSE Then
        Let Msg = "[  Z80MEM] " & TokenStrings(Token)
    ElseIf Token = TOKEN_PREFIX_K Or Token = TOKEN_PREFIX_KB _
        Or Token = TOKEN_PREFIX_KBIT Then
        Let Msg = "[  PREFIX] " & TokenStrings(Token)
    End If
    If Msg <> vbNullString Then RaiseEvent Message(2, OZ80_LOG_DEBUG, _
        "<tkn> " & Format$(TokenIndex + 1, "#,#") & _
            " " & FormatLineCol & _
            " " & FormatLong(Value) & " " & Msg _
    )
    
    'Remember the line number used for this token to check for properties occuring _
     on a new line -- this constitutes a property definition in a block
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
End Sub

'AddString : Strings are stored outside of the token stream and referenced by number _
 ======================================================================================
Private Sub AddString( _
    ByRef Token As OZ80_TOKEN, ByRef Data As String _
)
    'Note Line & Col numbers are global variables to this class to save having to _
     use them in calls hundreds of times over
    Dim Key As Long
    
    'Pass the string back to the assembler. It'll want to pool together all of the _
     strings used across all files for the purpose of label lookups
    'Note that the assembler will produce the key for us and pass it back
    RaiseEvent StringAdded(Key, Data)
    
    Dim TokenType As String * 8
    If Token = TOKEN_LABEL Then
        Let TokenType = "   LABEL"
    ElseIf Token = TOKEN_PROPERTY_NEW Or Token = TOKEN_PROPERTY_USE Then
        Let TokenType = "PROPERTY"
    ElseIf Token = TOKEN_QUOTE Then
        Let TokenType = "   QUOTE"
    ElseIf Token = TOKEN_RAM Then
        Let TokenType = "     RAM"
    ElseIf Token = TOKEN_SECTION Then
        Let TokenType = " SECTION"
    ElseIf Token = TOKEN_VARIABLE Then
        Let TokenType = "VARIABLE"
    Else
        Stop
    End If
    
    RaiseEvent Message(2, OZ80_LOG_DEBUG, _
        "<tkn> " & Format$(TokenIndex + 1, "#,#") & _
            " " & FormatLineCol & _
            " " & FormatLong(Key) & " " & _
            "[" & TokenType & "] " & Chr(34) & Data & Chr(34) _
    )
    
    'Add the index key to the token so that we can retrieve the string later
    Call Add(Token, Key)
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
    
    Dim TextArray() As Byte
    Let TextArray = oz80.CRC.CastStringToArray(CompText)
    
    Dim ii As Long
    For ii = 0 To WordLen
        If LowCase(Text(WordBegin + ii)) <> TextArray(ii * 2) Then Exit Function
    Next ii
    
    Let CompStr = True
End Function

'FileExists : See if a file exists or not _
 ======================================================================================
'<cuinl.tripod.com/Tips/fileexist.htm>
Private Function FileExists(ByRef Path As String) As Boolean
    Let FileExists = CBool(Dir$(Path) <> vbNullString)
End Function

'FormatLineCol : Formats the current line and column number being parsed for logging _
 ======================================================================================
Private Function FormatLineCol() As String
    Let FormatLineCol = "(" & _
        Right$("    " & LineBegin, 5) & "," & Left$(ColBegin & "  ", 3) & _
    ")"
End Function

'FormatLong : Format a 32-bit hexadecimal number for logging _
 ======================================================================================
Private Function FormatLong(ByRef Value As Long) As String
    Let FormatLong = "$" & Right$("0000000" & Hex(Value), 8)
    Let FormatLong = Left$(FormatLong, 5) & "-" & Right$(FormatLong, 4)
End Function

'IsValidName : Check validity for variables, labels & properties _
 ======================================================================================
Private Function IsValidName(ByRef First As Long, ByRef Last As Long) As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    
    'The first letter must be the type identifier; _
     :label, ?function, @macro, .property, #variable, $.RAMconstant
    If Text(First) <> SYNTAX_LABEL And _
       Text(First) <> SYNTAX_FUNCT And _
       Text(First) <> SYNTAX_MACRO And _
       Text(First) <> SYNTAX_PROPERTY And _
       Text(First) <> SYNTAX_VARIABLE And _
       Text(First) <> SYNTAX_NUMBER_HEX _
    Then
        Exit Function
    End If
    
    'Check for special names with two character prefixes, _
     e.g. "::section" and "$.RAMconstant"
    Dim Begin As Long
    If (Text(First) = SYNTAX_LABEL And Text(First + 1) = SYNTAX_LABEL) _
    Or (Text(First) = SYNTAX_NUMBER_HEX And Text(First + 1) = SYNTAX_PROPERTY) Then
        Let Begin = First + 2
    Else
        Let Begin = First + 1
    End If
    
    'The name cannot begin with a number or a dot
    If Text(Begin) = SYNTAX_PROPERTY Then Exit Function
    If IsNumber(Text(Begin)) Then Exit Function
    'Cannot end in a dot
    If Text(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim i As Long
    For i = Begin To Last
        'Not a dot?
        If Text(i) <> SYNTAX_PROPERTY Then
            'Is letter not alphanumeric?
            If IsAlphaNum(Text(i)) = False Then
                'Finally, not an underscore?
                If Text(i) <> ASC_UNDRSCR Then Exit Function
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If IsNumber(Text(i + 1)) Then Exit Function
            'Two dots in a row not allowed
            If Text(i + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next i
    
    Let IsValidName = True
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Function RaiseError(ByRef Number As OZ80_ERROR) As OZ80_ERROR
    'Send the error to the controller (if it sinks events)
    RaiseEvent Error(Number, LineBegin, ColBegin)
    
    'We return the same error number we were given only to make the tail recursion _
     simpler, e.g. `let TokeniseFile = RaiseError(OZ80_ERROR_...`
    Let RaiseError = Number
End Function
