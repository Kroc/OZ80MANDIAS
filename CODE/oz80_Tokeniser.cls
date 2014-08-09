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

'--------------------------------------------------------------------------------------

'These define the various punctiation marks (in ASCII codes) for the language syntax
Private Enum OZ80_SYNTAX
    SYNTAX_COMMENT = ASC_BTK            ' ` - Comment marker. "``" for multi-line
    SYNTAX_HINT1 = ASC_SCOL             ' ; - register hint, e.g. `a;index`
    SYNTAX_HINT2 = ASC_APOS             ' ' - shadow register hint, e.g. `ex af 'af`
    SYNTAX_QUOTE = ASC_QUOT             ' " - string identifier
    SYNTAX_LABEL = ASC_COL              ' : - label identifier
    SYNTAX_PROPERTY = ASC_DOT           ' . - property identifier
    SYNTAX_VARIABLE = ASC_HASH          ' # - variable identifier
    SYNTAX_MACRO = ASC_AT               ' @ - macro identifier
    SYNTAX_FUNCT = ASC_QM               ' ? - function identifier
    SYNTAX_NUMBER_HEX = ASC_DOL         ' $ - hexadecimal number, e.g. `$FFFF`
    SYNTAX_NUMBER_BIN = ASC_PERC        ' % - binary number, e.g. `%10101011`
    SYNTAX_NEXT = ASC_COM               ' , - item seperator, optional
    SYNTAX_PAREN_OPEN = ASC_LP          ' ( - memory reference open parenthesis
    SYNTAX_PAREN_CLOSE = ASC_RP         ' ) - memory reference close parenthesis
    SYNTAX_BLOCK_OPEN = ASC_LB          ' { - block open brace
    SYNTAX_BLOCK_CLOSE = ASC_RB         ' } - block close brace
    SYNTAX_OPERATOR_ADD = ASC_PLUS      ' + - Add
    SYNTAX_OPERATOR_SUB = ASC_HYP       ' - - Subtract
    SYNTAX_OPERATOR_MUL = ASC_STAR      ' * - Multiply
    SYNTAX_OPERATOR_DIV = ASC_FSL       ' / - Divide
    SYNTAX_OPERATOR_POW = ASC_CRT       ' ^ - Power
    SYNTAX_OPERATOR_MOD = ASC_BSL       ' \ - Modulus
    SYNTAX_OPERATOR_OR = ASC_VB         ' | - Bitwise OR
    SYNTAX_OPERATOR_AND = ASC_AMP       ' & - Bitwise AND
    SYNTAX_OPERATOR_NOT = ASC_EXC       ' ! - Bitwise NOT
    SYNTAX_OPERATOR_XOR = ASC_TIL       ' ~ - Bitwise XOR
End Enum

'Look-up tables to validate ASCII codes:
Private IsValid(0 To 255) As Boolean    'Allowable ASCII chars (i.e. no NULLs &c.)
Private UpCase(0 To 255) As Byte        'From lower case to upper case
Private IsHex(0 To 255) As Boolean      'Validates hexadecimal digits (0-9, A-F)
Private IsAlpha(0 To 255) As Boolean    'Validate alpha characters (both cases)
Private IsAlphaNum(0 To 255) As Boolean 'Validate alphanumeric chars (both cases)
Private IsNumber(0 To 255) As Boolean   'Validates decimal digits

'For logging, we will want to get a text representation of any of the Tokens
Private TokenNames(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String

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
 parameter on a Z80 instruction, and a label-property definition (within a block) _
 which occurs on a new line. We keep track of the previous line number used and _
 look for a property occuring on a new line
Private PrevLine As Long

'As the source text is parsed, valid letters are queued into the word buffer. _
 We use this to do comparisons as the source text may contain skipped letters
Private WordArr() As Byte
'The size of the word buffer. We avoid resizing this with every word (slow), _
 instead opting for a starting size and increasing permenantly if it goes over
Private WordArrLimit As Long
Private WordLen As Long

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
Private TokenKind As Byte
Private TokenValue As Double
Private TokenLine As Long
Private TokenCol As Long

'Look up tables to determine what category a Token belongs to; _
 these are exposed publicly through properties so other classes can do quick checks
Private My_IsInstruction(0 To 255) As Boolean
Private My_IsRegister(0 To 255) As Boolean
Private My_IsRegister8Bit(0 To 255) As Boolean
Private My_IsRegister16Bit(0 To 255) As Boolean
Private My_IsFlag(0 To 255) As Boolean
Private My_IsKeyword(0 To 255) As Boolean
Private My_IsOperator(0 To 255) As Boolean
Private My_IsExpression(0 To 255) As Boolean
Private My_IsParameter(0 To 255) As Boolean

'We could just unfold strings into a series of number tokens, but for quick access _
 to filenames with INCLUDE directives, we store strings in their own database
'NOTE: This is public so that the assembler can access the right string for a token _
 in order to normalise name items, e.g. `::section:label.property`
Public Strings As Scripting.Dictionary

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message( _
    ByRef Depth As Long, ByRef LogLevel As OZ80_LOG, _
    ByRef Text As String _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Set Stopwatch = New PerformanceCounter
    
    'Initialise the string database. _
     Since it's public, we want to ensure it's instantiated from the off
    Set Strings = New Scripting.Dictionary
    
    Let TokenNames(TOKEN_Z80_ADC) = "ADC"
    Let TokenNames(TOKEN_Z80_ADD) = "ADD"
    Let TokenNames(TOKEN_Z80_AND) = "AND"
    Let TokenNames(TOKEN_Z80_BIT) = "BIT"
    Let TokenNames(TOKEN_Z80_CALL) = "CALL"
    Let TokenNames(TOKEN_Z80_CCF) = "CCF"
    Let TokenNames(TOKEN_Z80_CP) = "CP"
    Let TokenNames(TOKEN_Z80_CPD) = "CPD"
    Let TokenNames(TOKEN_Z80_CPDR) = "CPDR"
    Let TokenNames(TOKEN_Z80_CPI) = "CPI"
    Let TokenNames(TOKEN_Z80_CPIR) = "CPIR"
    Let TokenNames(TOKEN_Z80_CPL) = "CPL"
    Let TokenNames(TOKEN_Z80_DAA) = "DAA"
    Let TokenNames(TOKEN_Z80_DEC) = "DEC"
    Let TokenNames(TOKEN_Z80_DI) = "DI"
    Let TokenNames(TOKEN_Z80_DJNZ) = "DJNZ"
    Let TokenNames(TOKEN_Z80_EI) = "EI"
    Let TokenNames(TOKEN_Z80_EX) = "EX"
    Let TokenNames(TOKEN_Z80_EXX) = "EXX"
    Let TokenNames(TOKEN_Z80_HALT) = "HALT"
    Let TokenNames(TOKEN_Z80_IM) = "IM"
    Let TokenNames(TOKEN_Z80_IN) = "IN"
    Let TokenNames(TOKEN_Z80_INC) = "INC"
    Let TokenNames(TOKEN_Z80_IND) = "IND"
    Let TokenNames(TOKEN_Z80_INDR) = "INDR"
    Let TokenNames(TOKEN_Z80_INI) = "INI"
    Let TokenNames(TOKEN_Z80_INIR) = "INIR"
    Let TokenNames(TOKEN_Z80_JP) = "JP"
    Let TokenNames(TOKEN_Z80_JR) = "JR"
    Let TokenNames(TOKEN_Z80_LD) = "LD"
    Let TokenNames(TOKEN_Z80_LDD) = "LDD"
    Let TokenNames(TOKEN_Z80_LDDR) = "LDDR"
    Let TokenNames(TOKEN_Z80_LDI) = "LDI"
    Let TokenNames(TOKEN_Z80_LDIR) = "LDIR"
    Let TokenNames(TOKEN_Z80_NEG) = "NEG"
    Let TokenNames(TOKEN_Z80_NOP) = "NOP"
    Let TokenNames(TOKEN_Z80_OR) = "OR"
    Let TokenNames(TOKEN_Z80_OUT) = "OUT"
    Let TokenNames(TOKEN_Z80_OUTD) = "OUTD"
    Let TokenNames(TOKEN_Z80_OTDR) = "OTDR"
    Let TokenNames(TOKEN_Z80_OUTI) = "OUTI"
    Let TokenNames(TOKEN_Z80_OTIR) = "OTIR"
    Let TokenNames(TOKEN_Z80_POP) = "POP"
    Let TokenNames(TOKEN_Z80_PUSH) = "PUSH"
    Let TokenNames(TOKEN_Z80_RES) = "RES"
    Let TokenNames(TOKEN_Z80_RET) = "RET"
    Let TokenNames(TOKEN_Z80_RETI) = "RETI"
    Let TokenNames(TOKEN_Z80_RETN) = "RETN"
    Let TokenNames(TOKEN_Z80_RLA) = "RLA"
    Let TokenNames(TOKEN_Z80_RL) = "RL"
    Let TokenNames(TOKEN_Z80_RLCA) = "RLCA"
    Let TokenNames(TOKEN_Z80_RLC) = "RLC"
    Let TokenNames(TOKEN_Z80_RLD) = "RLD"
    Let TokenNames(TOKEN_Z80_RRA) = "RRA"
    Let TokenNames(TOKEN_Z80_RR) = "RR"
    Let TokenNames(TOKEN_Z80_RRCA) = "RRCA"
    Let TokenNames(TOKEN_Z80_RRC) = "RRC"
    Let TokenNames(TOKEN_Z80_RRD) = "RRD"
    Let TokenNames(TOKEN_Z80_RST) = "RST"
    Let TokenNames(TOKEN_Z80_SBC) = "SBC"
    Let TokenNames(TOKEN_Z80_SCF) = "SCF"
    Let TokenNames(TOKEN_Z80_SET) = "SET"
    Let TokenNames(TOKEN_Z80_SLA) = "SLA"
    Let TokenNames(TOKEN_Z80_SRA) = "SRA"
    Let TokenNames(TOKEN_Z80_SLL) = "SLL"
    Let TokenNames(TOKEN_Z80_SRL) = "SRL"
    Let TokenNames(TOKEN_Z80_SUB) = "SUB"
    Let TokenNames(TOKEN_Z80_XOR) = "XOR"
    
    'Z80 Registers ....................................................................
    Let TokenNames(TOKEN_Z80_A) = "A"
    Let TokenNames(TOKEN_Z80_AF) = "AF"
    Let TokenNames(TOKEN_Z80_B) = "B"
    Let TokenNames(TOKEN_Z80_C) = "C"
    Let TokenNames(TOKEN_Z80_NC) = "NC"
    Let TokenNames(TOKEN_Z80_BC) = "BC"
    Let TokenNames(TOKEN_Z80_D) = "D"
    Let TokenNames(TOKEN_Z80_E) = "E"
    Let TokenNames(TOKEN_Z80_DE) = "DE"
    Let TokenNames(TOKEN_Z80_H) = "H"
    Let TokenNames(TOKEN_Z80_L) = "L"
    Let TokenNames(TOKEN_Z80_HL) = "HL"
    Let TokenNames(TOKEN_Z80_I) = "I"
    Let TokenNames(TOKEN_Z80_IX) = "IX"
    Let TokenNames(TOKEN_Z80_IXL) = "IXL"
    Let TokenNames(TOKEN_Z80_IXH) = "IXH"
    Let TokenNames(TOKEN_Z80_IY) = "IY"
    Let TokenNames(TOKEN_Z80_IYL) = "IYL"
    Let TokenNames(TOKEN_Z80_IYH) = "IYH"
    Let TokenNames(TOKEN_Z80_M) = "M"
    Let TokenNames(TOKEN_Z80_P) = "P"
    Let TokenNames(TOKEN_Z80_PC) = "PC"
    Let TokenNames(TOKEN_Z80_PE) = "PE"
    Let TokenNames(TOKEN_Z80_PO) = "PO"
    Let TokenNames(TOKEN_Z80_R) = "R"
    Let TokenNames(TOKEN_Z80_SP) = "SP"
    Let TokenNames(TOKEN_Z80_Z) = "Z"
    Let TokenNames(TOKEN_Z80_NZ) = "NZ"
    
    'Operators ........................................................................
    Let TokenNames(TOKEN_OPERATOR_ADD) = Chr$(SYNTAX_OPERATOR_ADD)
    Let TokenNames(TOKEN_OPERATOR_SUB) = Chr$(SYNTAX_OPERATOR_SUB)
    Let TokenNames(TOKEN_OPERATOR_MUL) = Chr$(SYNTAX_OPERATOR_MUL)
    Let TokenNames(TOKEN_OPERATOR_DIV) = Chr$(SYNTAX_OPERATOR_DIV)
    Let TokenNames(TOKEN_OPERATOR_POW) = Chr$(SYNTAX_OPERATOR_POW)
    Let TokenNames(TOKEN_OPERATOR_MOD) = Chr$(SYNTAX_OPERATOR_MOD)
    Let TokenNames(TOKEN_OPERATOR_REP) = "x"
    Let TokenNames(TOKEN_OPERATOR_OR) = Chr$(SYNTAX_OPERATOR_OR)
    Let TokenNames(TOKEN_OPERATOR_AND) = Chr$(SYNTAX_OPERATOR_AND)
    Let TokenNames(TOKEN_OPERATOR_NOT) = Chr$(SYNTAX_OPERATOR_NOT)
    Let TokenNames(TOKEN_OPERATOR_XOR) = Chr$(SYNTAX_OPERATOR_XOR)
    
    'Keywords .........................................................................
    Let TokenNames(TOKEN_KEYWORD_AT) = "AT"
    Let TokenNames(TOKEN_KEYWORD_AS) = "AS"
    Let TokenNames(TOKEN_KEYWORD_BANK) = "BANK"
    Let TokenNames(TOKEN_KEYWORD_BINARY) = "BINARY"
    Let TokenNames(TOKEN_KEYWORD_BYTE) = "BYTE"
    Let TokenNames(TOKEN_KEYWORD_DATA) = "DATA"
    Let TokenNames(TOKEN_KEYWORD_DEFAULT) = "DEFAULT"
    Let TokenNames(TOKEN_KEYWORD_ECHO) = "ECHO"
    Let TokenNames(TOKEN_KEYWORD_ELSE) = "ELSE"
    Let TokenNames(TOKEN_KEYWORD_EXISTS) = "EXISTS"
    Let TokenNames(TOKEN_KEYWORD_FAIL) = "FAIL"
    Let TokenNames(TOKEN_KEYWORD_FILL) = "FILL"
    Let TokenNames(TOKEN_KEYWORD_IF) = "IF"
    Let TokenNames(TOKEN_KEYWORD_INCLUDE) = "INCLUDE"
    Let TokenNames(TOKEN_KEYWORD_LENGTH) = "LENGTH"
    Let TokenNames(TOKEN_KEYWORD_OBJECT) = "OBJECT"
    Let TokenNames(TOKEN_KEYWORD_PARAMS) = "PARAMS"
    Let TokenNames(TOKEN_KEYWORD_PROC) = "PROC"
    Let TokenNames(TOKEN_KEYWORD_RAM) = "RAM"
    Let TokenNames(TOKEN_KEYWORD_RETURN) = "RETURN"
    Let TokenNames(TOKEN_KEYWORD_SECTION) = "SECTION"
    Let TokenNames(TOKEN_KEYWORD_SLOT) = "SLOT"
    Let TokenNames(TOKEN_KEYWORD_START) = "START"
    Let TokenNames(TOKEN_KEYWORD_STOP) = "STOP"
    Let TokenNames(TOKEN_KEYWORD_STRUCT) = "STRUCT"
    Let TokenNames(TOKEN_KEYWORD_TABLE) = "TABLE"
    Let TokenNames(TOKEN_KEYWORD_VAR) = "VAR"
    Let TokenNames(TOKEN_KEYWORD_WORD) = "WORD"
    
    Let TokenNames(TOKEN_PREFIX_K) = "K"
    Let TokenNames(TOKEN_PREFIX_KB) = "KB"
    Let TokenNames(TOKEN_PREFIX_KBIT) = "KBIT"
    
    Let TokenNames(TOKEN_PARENOPEN) = Chr$(SYNTAX_PAREN_OPEN)
    Let TokenNames(TOKEN_PARENCLOSE) = Chr$(SYNTAX_PAREN_CLOSE)
    Let TokenNames(TOKEN_BLOCKOPEN) = Chr$(SYNTAX_BLOCK_OPEN)
    Let TokenNames(TOKEN_BLOCKCLOSE) = Chr$(SYNTAX_BLOCK_CLOSE)
    
    Let TokenNames(TOKEN_QUOTE) = Chr$(SYNTAX_QUOTE)
    Let TokenNames(TOKEN_LABEL) = Chr$(SYNTAX_LABEL)
    Let TokenNames(TOKEN_PROPERTY_USE) = Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_PROPERTY_NEW) = Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_RAM) = Chr$(SYNTAX_NUMBER_HEX) & Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_SECTION) = String(2, Chr$(SYNTAX_LABEL))
    Let TokenNames(TOKEN_VARIABLE) = Chr$(SYNTAX_VARIABLE)
    
    '----------------------------------------------------------------------------------
    
    'Save time doing and/or comparisons by compiling look up tables of ASCII _
     codes that count as numeric / hex / alpha and alphanumeric
    Dim ii As OZ80_TOKEN
    For ii = 0 To 255
        'Is this a visible, allowed ASCII-code?
        Let IsValid(ii) = (ii = ASC_TAB) Or (ii >= ASC_SPC And ii < ASC_DEL)
        
        'Is this a number?
        If ii >= ASC_0 And ii <= ASC_9 Then
            Let IsNumber(ii) = True     'Is a valid number digit
            Let IsHex(ii) = True        'Is a valid hexadecimal digit
            Let IsAlphaNum(ii) = True   'Is alphanumeric
            Let UpCase(ii) = ii         'Is already upper case
        
'        'Allow the decimal point for numbers
'        ElseIf ii = ASC_DOT Then
'            Let IsNumber(ii) = True
        
        'Is this lowercase a-z?
        ElseIf ii >= ASC_a_ And ii <= ASC_z_ Then
            'Convert to uppercase
            Let UpCase(ii) = ii And Not 32
            Let IsAlpha(ii) = True      'Is alpha
            Let IsAlphaNum(ii) = True   'Is alphanumeric too
            'If A-F then also hexadecimal
            If ii <= ASC_f_ Then Let IsHex(ii) = True
        
        'Is uppercase A-Z?
        ElseIf ii >= ASC_A And ii <= ASC_Z Then
            Let UpCase(ii) = ii         'Already uppercase
            Let IsAlpha(ii) = True      'Is alpha
            Let IsAlphaNum(ii) = True   'Is alphanumeric
            'If A-F then also hexadecimal
            If ii <= ASC_F Then Let IsHex(ii) = True
        
        'No special case
        Else
            Let UpCase(ii) = ii         'Is already uppercase
        End If
        
        '..............................................................................
        
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
    Erase IsValid, UpCase, IsHex, IsAlpha, IsAlphaNum, IsNumber
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

'PROPERTY Col : Column number in the source of the current token _
 ======================================================================================
Public Property Get Col() As Long: Let Col = TokenCol: End Property

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
    Let EOF = CBool(TokenIndex >= TokenCount)
End Property

'PROPERTY Error : Expose the internal error number _
 ======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property

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
    Let IsExpression = My_IsExpression(TokenKind)
End Property

'PROPERTY IsFlag : Is this token number a condition flag? _
 ======================================================================================
Public Property Get IsFlag() As Boolean
    Let IsFlag = My_IsFlag(TokenKind)
End Property

'PROPERTY IsInstruction : Is this token number a Z80 instruction? _
 ======================================================================================
Public Property Get IsInstruction() As Boolean
    Let IsInstruction = My_IsInstruction(TokenKind)
End Property

'PROPERTY IsKeyword : Is this token number a OZ80 Keyword? _
 ======================================================================================
Public Property Get IsKeyword() As Boolean
    Let IsKeyword = My_IsKeyword(TokenKind)
End Property

'PROPERTY IsOperator : Is this token number an OZ80 operator? _
 ======================================================================================
Public Property Get IsOperator() As Boolean
    Let IsOperator = My_IsOperator(TokenKind)
End Property

'PROPERTY IsPrameter : Is this token number a Z80 instruction parameter? _
 ======================================================================================
Public Property Get IsParameter() As Boolean
    Let IsParameter = My_IsParameter(TokenKind)
End Property

'PROPERTY IsRegister : Is this token number a Register (8 or 16-Bit)? _
 ======================================================================================
Public Property Get IsRegister() As Boolean
    Let IsRegister = My_IsRegister(TokenKind)
End Property

'PROPERTY IsRegister8Bit : Is this token number an 8-bit Register? _
 ======================================================================================
Public Property Get IsRegister8Bit() As Boolean
    Let IsRegister8Bit = My_IsRegister8Bit(TokenKind)
End Property

'PROPERTY IsRegister16Bit : Is this token number a 16-bit Register? _
 ======================================================================================
Public Property Get IsRegister16Bit() As Boolean
    Let IsRegister16Bit = My_IsRegister16Bit(TokenKind)
End Property

'PROPERTY Kind : Token Kind of current token _
 ======================================================================================
Public Property Get Kind() As OZ80_TOKEN: Let Kind = TokenKind: End Property

'PROPERTY Line : Line number in the source of the current token _
 ======================================================================================
Public Property Get Line() As Long: Let Line = TokenLine: End Property

'PROPERTY Value : Token Value of current token _
 ======================================================================================
Public Property Get Value() As Double: Let Value = TokenValue: End Property

'Forward : Move onto the next token, returns false if the TokenStream hit the end _
 ======================================================================================
Public Function Forward() As Boolean
    Let TokenIndex = TokenIndex + 1
    Call CacheToken
    
    If TokenIndex >= TokenCount Then Exit Function
    Let Forward = True
    
    'Generate a debug message with the details of the now-current token
    RaiseEvent Message(2, OZ80_LOG_DEBUG, FormatToken())
End Function

'Rewind : Go back to the start of the tokenStream _
 ======================================================================================
Public Sub Rewind()
    Let TokenIndex = 0
    Call CacheToken
End Sub

'Tokenise : Break the source file into individual words and convert to machine tokens _
 ======================================================================================
Public Function Tokenise(ByRef FilePath As String) As OZ80_ERROR
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
    If Not FileExists(FilePath) Then Call RaiseError( _
        OZ80_ERROR_FILENOTFOUND _
    ): GoTo Finish
    
    'The source text file is held here as a byte-array whilst it gets parsed
    Dim Text() As Byte
    
    'Once the file is open, ensure it's closed upon error
    On Error GoTo ErrorOpenFile
    Dim FileNumber As Integer: Let FileNumber = FreeFile
    Open FilePath For Input Access Read Lock Write As #FileNumber
    
    'Read the file as a binary string with one-byte per character (ASCII). _
     We first normalise to Unicode to deal with locale specifics and then _
     downsample into ASCII to save memory and make stepping through easier
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
        "Tokenising " & Chr(ASC_QUOT) & FilePath & Chr(ASC_QUOT) _
    )
    RaiseEvent Message(2, OZ80_LOG_INFO, _
        "Length " & Format$(TextLen, "#,#") & " bytes" _
    )
    
    'Clear the string database, just to be extra safe
    Call Strings.RemoveAll
    
    'Begin parsing words: _
     ----------------------------------------------------------------------------------
    Dim IsQuote As Boolean
    Dim IsComment As Boolean
    Dim IsCommentMulti As Boolean
    
    'Reset the buffer used to hold a word as it's built
    Let WordArrLimit = 255
    ReDim WordArr(0 To WordArrLimit) As Byte
    
    'These are used whilst converting words into real numbers
    Dim Number As Double, Multiplier As Long
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
    
    'Read a single word: _
     ..................................................................................
ReadWord:
    'If the word has begun (first letter is read)
    Dim IsWord As Boolean
    Let IsWord = False
    'Reset the length of the word. We have to count as we go as not all letters are _
     valid and constitute part of the word (line-breaks for example), therefore it is _
     important to note that `WordLen` is not necessarily `WordEnd - WordBegin`
    Let WordLen = 0
SkipChar:
    'Move on to the next character in the source
    Let TextPos = TextPos + 1
ReadChar:
    'If the file ends?
    If TextPos >= TextLen Then
        'If the word is currently underway then we are all done
        If Not IsWord Then GoTo Finish
        'Otherwise process this final word
        GoTo EndWord
    End If
    
    'Read a character:
    Dim AscChar As ASCII
    Let AscChar = Text(TextPos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the assembling stage)
    
    'New line:
    If AscChar = ASC_LF Then
        'NOTE: Carriage-Returns are ignored to avoid doubling line number and to _
         maintain compatibility with UNIX line-endings which use Line-Feed only.
        'NOTE: This causes the side-effect that a word will end on the Line-Feed, _
         unintentially leaving the Carriage-Return as part of the word, counting _
         `WordLen` only with each valid letter is done to avoid errors
        Let Line = Line + 1: Let Col = 0
    
    ElseIf AscChar = ASC_TAB Then
        'Tab goes to the next 8 char boundary:
        Let Col = Col + 8 - (Col Mod 8)
    
    'Non-printable characters are skipped and do not add to the column number
    ElseIf Not IsValid(AscChar) Then
        GoTo SkipChar
    Else
        Let Col = Col + 1
    End If
    
    'Comments: ........................................................................
    If IsComment Then
        'For single-line comments:
        If Not IsCommentMulti Then
            'Is this a multi-line comment? (two backticks)
            If IsWord Then
                If AscChar = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
            'The end of a line ends a single-line comment:
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
        Let WordBegin = TextPos: Let LineBegin = Line: Let ColBegin = Col
        Let WordLen = 0: Let IsWord = True
    End If
    'Since this is a valid letter, capture it and increase the valid word length
    Let WordArr(WordLen) = AscChar
    Let WordLen = WordLen + 1
    'Increase our word buffer if the word gets too long (likely a string)
    If WordLen > WordArrLimit Then
        Let WordArrLimit = WordArrLimit + 256
        ReDim Preserve WordArr(WordArrLimit) As Byte
    End If
    GoTo SkipChar

    '----------------------------------------------------------------------------------
Punctuation:
    Let WordBegin = TextPos
    Let WordEnd = TextPos
    Let WordArr(0) = AscChar
    Let WordLen = 1
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let TextPos = TextPos - 1: Let Col = Col - 1
    Let WordEnd = TextPos
    GoTo TokenWord
    
EndWord:
    Let WordEnd = TextPos - 1

    'Tokenise the word: _
     ----------------------------------------------------------------------------------
TokenWord:
    'Extract the 1st, 2nd, 3rd & 4th letters of the word as these make up most _
     comparisons and can be done quickly without resorting to strings
    Dim Asc1 As Long, Asc2 As Long, Asc3 As Long, Asc4 As Long
    Let Asc1 = UpCase(WordArr(0))
    If WordLen > 3 Then Let Asc4 = UpCase(WordArr(3)) Else Let Asc4 = ASC_NUL
    If WordLen > 2 Then Let Asc3 = UpCase(WordArr(2)) Else Let Asc3 = ASC_NUL
    If WordLen > 1 Then Let Asc2 = UpCase(WordArr(1)) Else Let Asc2 = ASC_NUL
    
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
    ElseIf IsNumber(Asc1) Then
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For ii = WordLen - 1 To 0 Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If Not IsNumber(WordArr(ii)) Then
                
                GoSub GetWord
                Debug.Print Word
                Stop
                
                'ERROR: Invalid number!
                Call RaiseError(OZ80_ERROR_INVALID_NUMBER_DEC)
                GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (WordArr(ii) - ASC_0) * Multiplier
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
        If WordLen > 33 Then GoTo ErrorOverflow
        Let Number = 0: Let Multiplier = 0
        For ii = WordLen - 1 To 1 Step -1
            If WordArr(ii) = ASC_1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf WordArr(ii) <> ASC_0 Then
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
        'Validate the number:
        'Check the digits are 0-9, A-F
        For ii = 1 To WordLen - 1
            'ERROR: Invalid hexadecimal number!
            If IsHex(WordArr(ii)) = False Then Call RaiseError( _
                OZ80_ERROR_INVALID_NUMBER_HEX _
            ): GoTo Finish
        Next ii
        'Any more than 8-digits will be an overflow (32-bit number maximum)
        If WordLen > 8 Then GoTo ErrorOverflow
        
        'Now cast the string into a real number _
         (Note: The digits are valid, but it could be an overflow)
        On Error GoTo ErrorOverflow
        GoSub GetWord
        Let Number = Val("&H" & Mid$(Word, 2) & "&")
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
             i.e. it defines a new sub-label of a block
            If LineBegin <> PrevLine Then _
                Call AddString(TOKEN_PROPERTY_NEW) Else _
                Call AddString(TOKEN_PROPERTY_USE)
            GoTo NextWord
        End If
    
    '.......................................................................... `#` ...
    ElseIf Asc1 = SYNTAX_VARIABLE Then
        'A variable name cannot be just '#'
        'A variable cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        If (Asc2 = ASC_NUL) Or (Not IsValidName()) Then
            'TODO: Variable specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
        Else
            Call AddString(TOKEN_VARIABLE)
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
        If Asc2 = ASC_NUL Then
            'Just `K`
            Call Add(TOKEN_PREFIX_K): GoTo NextWord
        
        'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = ASC_B Then
            'If only two letters, then it's `KB`
            If Asc3 = ASC_NUL Then Call Add(TOKEN_PREFIX_KB): GoTo NextWord
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
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_M): GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_N Then '................................................. `n` ...
        'Z80 instructions
        If CompStr("neg") Then Call Add(TOKEN_Z80_NEG): GoTo NextWord
        If CompStr("nop") Then Call Add(TOKEN_Z80_NOP): GoTo NextWord
        'Z80 flag conditions
        If Asc3 = ASC_NUL Then
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
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_P): GoTo NextWord
        If Asc3 = ASC_NUL Then
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
        If Asc2 = ASC_NUL Then Call Add(TOKEN_Z80_Z): GoTo NextWord
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
    'Build a string of the word from the byte array
    Let Word = vbNullString
    For ii = 0 To WordLen - 1: Let Word = Word & Chr$(WordArr(ii)): Next ii
    Return

    '----------------------------------------------------------------------------------
ErrorOpenFile:
    'If an error occured during file reading, ensure the file handle is closed
    On Error GoTo 0
    Close #FileNumber
    'TODO: Send the system error details, Err.Number / Err.Description
    Call RaiseError(OZ80_ERROR_FILEREAD)
    GoTo Finish

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
        
        Erase Text
        RaiseEvent Message(2, OZ80_LOG_INFO, "Words  " & Format$(WordCount, "#,#"))
        RaiseEvent Message(2, OZ80_LOG_INFO, "Tokens " & Format$(TokenCount, "#,#"))
        RaiseEvent Message(2, OZ80_LOG_INFO, "Time   " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    Else
        'clean up!
        'TODO: This leaves the class in a broken state _
         (no tokens array anymore and no call to `cacheToken`)
        Erase Tokens
        Let TokenCount = 0
    End If
    
    'Return the error number to the assembler
    Let Tokenise = My_Error
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Add : Add a token (any) to the stream _
 ======================================================================================
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
    RaiseEvent Message( _
        2, OZ80_LOG_DEBUG, FormatToken() _
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
    
    Call CacheToken
End Sub

'AddString : Strings are stored outside of the token stream and referenced by number _
 ======================================================================================
Private Sub AddString( _
    ByRef Token As OZ80_TOKEN _
)
    'Note Line & Col numbers are global variables to this class to save having to _
     use them in calls hundreds of times over. Again, speed over properness
    
    'Convert the word buffer into a string
    Dim Data As String, ii As Long
    For ii = 0 To WordLen - 1: Let Data = Data & Chr$(WordArr(ii)): Next ii
    'Ignore the quote mark that is left on the front of OZ80 strings
    If Token = TOKEN_QUOTE Then Let Data = Mid$(Data, 2)
    
    'Hash the string to quickly identify if it has already been added
    Dim Key As Long
    Let Key = oz80.CRC.Calculate(Data)
    If Not Strings.Exists(Key) Then Call Strings.Add(Key, Data)
    
    'Add the index key to the token so that we can retrieve the string later
    Call Add(Token, CDbl(Key))
End Sub

'CacheToken : Cache the current token info _
 ======================================================================================
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

'CompStr : Compare a word from the byte array to a string _
 ======================================================================================
Private Function CompStr( _
    ByRef CompText As String _
) As Boolean
    'This function is designed as a quick'n'dirty replacement for `StrComp`, _
     but using the byte array as the source so we no longer have to concatenate _
     strings in the parsing routine
    
    'If the lengths don't match then they can't be the same
    If WordLen <> Len(CompText) Then Exit Function
    
    'Cast the string parameter into a byte array
    Dim TextArray() As Byte
    Let TextArray = oz80.CRC.CastStringToArray(CompText)
    
    Dim ii As Long
    For ii = 0 To (WordLen - 1)
        'NOTE: the test string is in Unicode, so 2-bytes per char
        If UpCase(WordArr(ii)) <> UpCase(TextArray(ii * 2)) Then Exit Function
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

'FormatToken : Format the details of the current token for logging _
 ======================================================================================
Private Function FormatToken() As String
    Dim Msg As String
    If TokenKind = TOKEN_BLOCKOPEN _
    Or TokenKind = TOKEN_BLOCKCLOSE Then
        Let Msg = "[   BLOCK] " & TokenNames(TokenKind)
    ElseIf My_IsKeyword(TokenKind) Then
        Let Msg = "[ KEYWORD] " & UCase(TokenNames(TokenKind))
    ElseIf TokenKind = TOKEN_LABEL Then
        Let Msg = "[   LABEL] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf TokenKind = TOKEN_NUMBER Then
        Let Msg = "[  NUMBER] " & Format$(Value, "#,#")
    ElseIf My_IsOperator(TokenKind) Then
        Let Msg = "[OPERATOR] " & TokenNames(TokenKind)
    ElseIf TokenKind = TOKEN_PREFIX_K _
        Or TokenKind = TOKEN_PREFIX_KB _
        Or TokenKind = TOKEN_PREFIX_KBIT Then
        Let Msg = "[  PREFIX] " & TokenNames(TokenKind)
    ElseIf TokenKind = TOKEN_PROPERTY_NEW _
        Or TokenKind = TOKEN_PROPERTY_USE Then
        Let Msg = "[PROPERTY] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf TokenKind = TOKEN_QUOTE Then
        Let Msg = "[   QUOTE] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf TokenKind = TOKEN_RAM Then
        Let Msg = "[     RAM] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf My_IsRegister(TokenKind) Then
        Let Msg = "[REGISTER] " & TokenNames(TokenKind)
    ElseIf TokenKind = TOKEN_SECTION Then
        Let Msg = "[ SECTION] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf TokenKind = TOKEN_VARIABLE Then
        Let Msg = "[VARIABLE] " & Chr(34) & Strings(TokenValue) & Chr(34)
    ElseIf My_IsInstruction(TokenKind) Then
        Let Msg = "[     Z80] " & TokenNames(TokenKind)
    ElseIf TokenKind = TOKEN_PARENOPEN _
        Or TokenKind = TOKEN_PARENCLOSE Then
        Let Msg = "[  Z80MEM] " & TokenNames(TokenKind)
    End If
    'TODO: need to check how the 64-bit number gets clipped to 32-bits here
    Let FormatToken = _
        "<tkn> " & Format$(TokenIndex + 1, "#,#") & _
            " " & FormatLineCol & _
            " " & FormatLong(CLng(TokenValue)) & " " & Msg
End Function

'GetValidNumber : Fully validate and convert a number from text into a real number _
 ======================================================================================
Private Function GetValidNumber( _
    ByRef ReturnValue As Double _
) As OZ80_ERROR
    'Check the first character to handle hexadecimal and binary numbers first
'    If Text(First) = SYNTAX_NUMBER_HEX Then
'        '
'    ElseIf Text(First) = SYNTAX_NUMBER_BIN Then
'        '
'    End If
    
    'TODO: Validate the range
End Function

'IsValidName : Check validity for named items _
 ======================================================================================
Private Function IsValidName() As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    
    'The first letter must be the type identifier; _
     `:label`, `?function`, `@macro`, `.property`, `#variable`, `$.RAMname`
    If WordArr(0) <> SYNTAX_LABEL And _
       WordArr(0) <> SYNTAX_FUNCT And _
       WordArr(0) <> SYNTAX_MACRO And _
       WordArr(0) <> SYNTAX_PROPERTY And _
       WordArr(0) <> SYNTAX_VARIABLE And _
       WordArr(0) <> SYNTAX_NUMBER_HEX _
    Then
        Exit Function
    End If
    
    'Check for special names with two character prefixes, _
     e.g. "::section" and "$.RAMconstant"
    Dim Begin As Long
    'TODO: Watch out for word-length! Could be a 1-letter label
    If (WordArr(0) = SYNTAX_LABEL And WordArr(1) = SYNTAX_LABEL) _
    Or (WordArr(0) = SYNTAX_NUMBER_HEX And WordArr(1) = SYNTAX_PROPERTY) Then
        Let Begin = 2
    Else
        Let Begin = 1
    End If
    
    'The name cannot begin with a number or a dot
    If WordArr(Begin) = SYNTAX_PROPERTY Then Exit Function
    If IsNumber(WordArr(Begin)) Then Exit Function
    
    'Cannot end in a dot
    Dim Last As Long
    Let Last = WordLen - 1
    If WordArr(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim i As Long
    For i = Begin To Last
        'Not a dot?
        If WordArr(i) <> SYNTAX_PROPERTY Then
            'Is letter not alphanumeric?
            If IsAlphaNum(WordArr(i)) = False Then
                'Finally, not an underscore?
                If WordArr(i) <> ASC_USC Then Exit Function
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If IsNumber(WordArr(i + 1)) Then Exit Function
            'Two dots in a row not allowed
            If WordArr(i + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next i
    
    Let IsValidName = True
End Function

'RiaseError : Provide error details through an event _
 ======================================================================================
Private Sub RaiseError(ByRef Number As OZ80_ERROR)
    'Set the error number on this class
    Let My_Error = Number
    
    'Send the error to the controller (if it sinks events)
    RaiseEvent Error(Number, LineBegin, ColBegin)
End Sub
