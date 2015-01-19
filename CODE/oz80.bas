Attribute VB_Name = "oz80"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: oz80

'Public, shared stuff

'/// DEBUG ////////////////////////////////////////////////////////////////////////////

Public Profiler As New bluProfiler

Public Enum PROFILER_EVENTS
    EVENT_TOKENISE                      'File.Tokenise
    EVENT_TOKENISE_READWORD             '- parse out a single word
    EVENT_TOKENISE_TOKENWORD            '- Tokenise a single word
    EVENT_FORMATTOKEN                   '- Format token data for logging
    EVENT_PROCESSZ80                    'Assembler.ProcessZ80
    EVENT_PROCESSZ80_LOG                '- Output the disassembly
End Enum

'/// PUBLIC ENUMS /////////////////////////////////////////////////////////////////////

Public Enum OZ80_LOG
    OZ80_LOG_ACTION                     'The key important happenings
    OZ80_LOG_INFO                       'Optional information, not actions happening
    OZ80_LOG_STATUS                     'Display variable values &c. when assigned
    OZ80_LOG_DEBUG                      'Internal information for debugging purposes
End Enum

Public Enum OZ80_WARNING
    OZ80_WARNING_NONE                   'Skip "0"
    OZ80_WARNING_ROUND                  'A decimal number had to be round-down
End Enum

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_DUPLICATE_CONSTANT       '- Constant already defined
    OZ80_ERROR_DUPLICATE_LABEL          '- Duplicate Label
    OZ80_ERROR_DUPLICATE_PROC_HELP      '- Duplicate `HELP` parameter
    OZ80_ERROR_DUPLICATE_PROC_INTERRUPT '- Duplicate `INTERRUPT` parameter
    OZ80_ERROR_DUPLICATE_PROC_PARAMS    '- Duplicate `PARAMS` parameter
    OZ80_ERROR_DUPLICATE_PROC_RETURN    '- Duplicate `RETURN` parameter
    OZ80_ERROR_DUPLICATE_PROC_SECTION   '- Duplicate `SECTION` parameter
    OZ80_ERROR_DUPLICATE_SECTION        '- Can't define a Section twice
    OZ80_ERROR_DUPLICATE_START          '- Duplicate `START` Procedure
    OZ80_ERROR_DUPLICATE_TABLE_SECTION  '- Duplicate `SECTION` parameter
    OZ80_ERROR_EXPECTED                 'Incorrect content at the current scope
    OZ80_ERROR_EXPECTED_BRACKET         '- Close bracket ("}","]",")") without open
    OZ80_ERROR_EXPECTED_EXPRESSION      '- Expression required here
    OZ80_ERROR_EXPECTED_PROC_NAME       '- A label name must follow `PROC`
    OZ80_ERROR_EXPECTED_PROC_PARAMS     '- Invalid stuff in the `PARAMS` list
    OZ80_ERROR_EXPECTED_PROC_RETURN     '- Invalid stuff in the `RETURN` list
    OZ80_ERROR_EXPECTED_ROOT            '- Only certain keywords allowed at root
    OZ80_ERROR_EXPECTED_SECTION_NAME    '- A section name must follow `SECTION`
    OZ80_ERROR_EXPECTED_SYSTEM_NAME     '- A system name must follow `SYSTEM`
    OZ80_ERROR_EXPECTED_TABLE_NAME      '- A label name must follow `TABLE`
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_EXPRESSION_Z80           '- Not a valid Z80 instruction parameter
    OZ80_ERROR_FILE_END                 'Unexpected end of file
    OZ80_ERROR_FILE_NOTFOUND            'Requested file does not exist
    OZ80_ERROR_FILE_READ                'Some kind of problem while file handle open
    OZ80_ERROR_INDEFINITE               'Indefinite value cannot be used here
    OZ80_ERROR_INVALID_INTERRUPT        'Invalid Interrupt address
    OZ80_ERROR_INVALID_NAME             'Invalid label/property/variable name
    OZ80_ERROR_INVALID_NAME_RAM         '- Invalid RAM name, i.e. `$.name`
    OZ80_ERROR_INVALID_NAME_HASH        '- Invalid hash name, i.e. `#hash`
    OZ80_ERROR_INVALID_NUMBER           'Not a valid number
    OZ80_ERROR_INVALID_NUMBER_DEC       '- Invalid decimal number
    OZ80_ERROR_INVALID_NUMBER_HEX       '- Invalid hexadecimal number
    OZ80_ERROR_INVALID_NUMBER_BIN       '- Invalid binary number
    OZ80_ERROR_INVALID_PROC_INTERRUPT   'SECTION & INTERRUPT params cannot co-exist
    OZ80_ERROR_INVALID_SECTION          'Section used, but not defined
    OZ80_ERROR_INVALID_SLOT             'Incorrect use of the Slot parameter
    OZ80_ERROR_INVALID_WORD             'Couldn't parse a word
    OZ80_ERROR_INVALID_Z80PARAMS        'Not the right parameters for a Z80 instruction
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_OVERFLOW_HILO            '- HI & LO functions limited to 16-bit inputs
    OZ80_ERROR_OVERFLOW_LINE            '- Line too long
    OZ80_ERROR_OVERFLOW_FILE            '- File too long / large
    OZ80_ERROR_OVERFLOW_Z80             '- 16-bit number used with an 8-bit instruction
    OZ80_ERROR_TEXT_CHAR                'Character code out of range
    OZ80_ERROR_UNDEFINED                'Named item is used, but undefined
    OZ80_ERROR_UNDEFINED_CONST          '- A Constant has been used before definition
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_TOKEN
    TOKEN_NONE                          'Skip "0"
    
    [_TOKEN_FIRST]
    [_TOKEN_Z80_BEGIN] = [_TOKEN_FIRST]
    
    'These are just the mnemonic tokens -- the assembler checks the
     'parameters and determines which opcode should be used
    TOKEN_Z80_ADC = [_TOKEN_FIRST]      'Add with Carry
    TOKEN_Z80_ADD                       'Add
    TOKEN_Z80_AND                       'Bitwise AND
    TOKEN_Z80_BIT                       'Bit test
    TOKEN_Z80_CALL                      'Call routine
    TOKEN_Z80_CCF                       'Clear Carry Flag
    TOKEN_Z80_CP                        'Compare
    TOKEN_Z80_CPD                       'Compare and Decrement
    TOKEN_Z80_CPDR                      'Compare, Decrement and Repeat
    TOKEN_Z80_CPI                       'Compare and Increment
    TOKEN_Z80_CPIR                      'Compare, Increment and Repeat
    TOKEN_Z80_CPL                       'Complement (bitwise NOT)
    TOKEN_Z80_DAA                       'Decimal Adjust Accumulator
    TOKEN_Z80_DEC                       'Decrement
    TOKEN_Z80_DI                        'Disable Interrupts
    TOKEN_Z80_DJNZ                      'Decrement and Jump if Not Zero
    TOKEN_Z80_EI                        'Enable Inettupts
    TOKEN_Z80_EX                        'Exchange
    TOKEN_Z80_EXX                       'Exchange shadow registers
    TOKEN_Z80_HALT                      'Stop CPU (wait for interrupt)
    TOKEN_Z80_IM                        'Interrupt Mode
    TOKEN_Z80_IN                        'Input from port
    TOKEN_Z80_INC                       'Increment
    TOKEN_Z80_IND                       'Input and Decrement
    TOKEN_Z80_INDR                      'Input, Decrement and Repeat
    TOKEN_Z80_INI                       'Input and Increment
    TOKEN_Z80_INIR                      'Input, Increment and Repeat
    TOKEN_Z80_JP                        'Jump
    TOKEN_Z80_JR                        'Jump Relative
    TOKEN_Z80_LD                        'Load
    TOKEN_Z80_LDD                       'Load and Decrement
    TOKEN_Z80_LDDR                      'Load, Decrement and Repeat
    TOKEN_Z80_LDI                       'Load and Increment
    TOKEN_Z80_LDIR                      'Load, Increment and Repeat
    TOKEN_Z80_NEG                       'Negate (flip the sign)
    TOKEN_Z80_NOP                       'No Operation (do nothing)
    TOKEN_Z80_OR                        'Bitwise OR
    TOKEN_Z80_OUT                       'Output to port
    TOKEN_Z80_OUTD                      'Output and Decrement
    TOKEN_Z80_OTDR                      'Output, Decrement and Repeat
    TOKEN_Z80_OUTI                      'Output and Increment
    TOKEN_Z80_OTIR                      'Output, Increment and Repeat
    TOKEN_Z80_POP                       'Pull from stack
    TOKEN_Z80_PUSH                      'Push onto stack
    TOKEN_Z80_RES                       'Reset bit
    TOKEN_Z80_RET                       'Return from routine
    TOKEN_Z80_RETI                      'Return from Interrupt
    TOKEN_Z80_RETN                      'Return from NMI
    TOKEN_Z80_RLA                       'Rotate Left (Accumulator)
    TOKEN_Z80_RL                        'Rotate Left
    TOKEN_Z80_RLCA                      'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLC                       'Rotate Left Circular
    TOKEN_Z80_RLD                       'Rotate Left 4-bits
    TOKEN_Z80_RRA                       'Rotate Right (Accumulator)
    TOKEN_Z80_RR                        'Rotate Right
    TOKEN_Z80_RRCA                      'Rotate Right Circular (Accumulator)
    TOKEN_Z80_RRC                       'Rotate Right Circular
    TOKEN_Z80_RRD                       'Rotate Right 4-bits
    TOKEN_Z80_RST                       '"Restart" -- Call a page 0 routine
    TOKEN_Z80_SBC                       'Subtract with Carry
    TOKEN_Z80_SCF                       'Set Carry Flag
    TOKEN_Z80_SET                       'Set bit
    TOKEN_Z80_SLA                       'Shift Left Arithmetic
    TOKEN_Z80_SRA                       'Shift Right Arithmetic
    TOKEN_Z80_SLL                       'Shift Left Logical
    TOKEN_Z80_SRL                       'Shift Right Logical
    TOKEN_Z80_SUB                       'Subtract
    TOKEN_Z80_XOR                       'Bitwise XOR
    [_TOKEN_Z80_END] = TOKEN_Z80_XOR
    
    'Z80 Registers & Flags ............................................................
    [_TOKEN_REGS_BEGIN]
    TOKEN_Z80_A = [_TOKEN_REGS_BEGIN]   'Accumulator
    TOKEN_Z80_AF                        'Accumulator and Flags
    TOKEN_Z80_B                         'Register B
    TOKEN_Z80_C                         'Register C or Carry flag
    TOKEN_Z80_NC                        'Carry unset flag
    TOKEN_Z80_BC                        'Register pair B & C
    TOKEN_Z80_D                         'Register D
    TOKEN_Z80_E                         'Register E
    TOKEN_Z80_DE                        'Register pair D & E
    TOKEN_Z80_H                         'Register H
    TOKEN_Z80_L                         'Register L
    TOKEN_Z80_HL                        'Register pair H & L
    TOKEN_Z80_I                         'Interrupt - not to be confused with IX & IY
    TOKEN_Z80_IX                        'Register IX
    TOKEN_Z80_IXL                       'Undocumented low-byte of register IX
    TOKEN_Z80_IXH                       'Undocumented high-byte of register IX
    TOKEN_Z80_IY                        'Register IY
    TOKEN_Z80_IYL                       'Undocumented low-byte of register IY
    TOKEN_Z80_IYH                       'Undocumented high-byte of register IY
    TOKEN_Z80_M                         'Sign is set flag
    TOKEN_Z80_P                         'Sign is not set flag
    TOKEN_Z80_PC                        'Program Counter
    TOKEN_Z80_PE                        'Parity/Overflow is set flag
    TOKEN_Z80_PO                        'Parity/Overflow is not set flag
    TOKEN_Z80_R                         'Refresh register (pseudo-random)
    TOKEN_Z80_SP                        'Stack Pointer
    TOKEN_Z80_Z                         'Zero set flag
    TOKEN_Z80_NZ                        'Zero not set flag
    [_TOKEN_REGS_END] = TOKEN_Z80_NZ
    
    'Operators ........................................................................
    [_TOKEN_OPERATORS_BEGIN]
    TOKEN_OPERATOR_ADD                  'Add "+"
    TOKEN_OPERATOR_SUB                  'Subtract "-"
    TOKEN_OPERATOR_MUL                  'Multiply "*"
    TOKEN_OPERATOR_DIV                  'Divide "/"
    TOKEN_OPERATOR_POW                  'Power "^"
    TOKEN_OPERATOR_MOD                  'Modulus "\"
    TOKEN_OPERATOR_REP                  'Repeat "x"
    TOKEN_OPERATOR_OR                   'Bitwise OR "|"
    TOKEN_OPERATOR_AND                  'Bitwise AND "&"
    TOKEN_OPERATOR_NOT                  'Bitwise NOT "!"
    TOKEN_OPERATOR_XOR                  'Bitwise XOR "~"
    [_TOKEN_OPERATORS_END]
    
    'Keywords .........................................................................
    [_TOKEN_KEYWORDS_BEGIN]
    TOKEN_KEYWORD_BOOL                  'Boolean data type (1-bit)
    TOKEN_KEYWORD_BYTE                  'Byte data type
    TOKEN_KEYWORD_DEF                   'Define constant
    TOKEN_KEYWORD_HASH                  'Define a hash-array
    TOKEN_KEYWORD_HELP                  'Documentation marker
    TOKEN_KEYWORD_HI                    '`HI` function -- high byte of 16-bit Value
    TOKEN_KEYWORD_INCLUDE               'Include another file
    TOKEN_KEYWORD_INTERRUPT             'Interrupt `PROC :<label> INTERRUPT <expr>`
    TOKEN_KEYWORD_LO                    '`LO` function -- low byte of a 16-bit Value
    TOKEN_KEYWORD_LONG                  'Long data type (4-bytes)
    TOKEN_KEYWORD_NYBL                  'Nybble data type (4-bits)
    TOKEN_KEYWORD_PARAMS                'Parameter list `PROC :<label> PARAMS <list>`
    TOKEN_KEYWORD_PROC                  'Procedure Chunk `PROC :<label> { ... }`
    TOKEN_KEYWORD_RAM                   'RAM definition
    TOKEN_KEYWORD_RETURN                'Returns list `PROC :<label> RETURN <list>`
    TOKEN_KEYWORD_SECTION               'Section definition `SECTION ::<section>
    TOKEN_KEYWORD_SLOT                  'Section Slot pattern `SLOT 0, 1, 2`
    TOKEN_KEYWORD_START                 'The starting vector for the System
    TOKEN_KEYWORD_SYSTEM                'System identifier `SYSTEM "SMS"`
    TOKEN_KEYWORD_TABLE                 'Data table
    TOKEN_KEYWORD_TRIP                  'Triple data type (3-bytes)
    TOKEN_KEYWORD_WORD                  'Word data type (2-bytes)
    [_TOKEN_KEYWORDS_END]
    
    TOKEN_NUMBER
    'Number prefixes ("K" & "KB")
    TOKEN_PREFIX_K                      'x1000
    TOKEN_PREFIX_KB                     'x1024
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_BRACES_OPEN                   '"{" Code/data Chunk, `PROC :<label> { ... }`
    TOKEN_BRACES_CLOSE                  '"}"
    TOKEN_SQUARE_OPEN                   '"[" Hash array and
    TOKEN_SQUARE_CLOSE                  '"]" Memory reference `ld a, [hl]`
    TOKEN_PARENS_OPEN                   '"(" Expression nesting,
    TOKEN_PARENS_CLOSE                  '")" e.g. `HI ($8000 + $80)`
    
    TOKEN_CONST                         'e.g. `!CONST`
    TOKEN_HASH                          'e.g. `#hash`
    TOKEN_LABEL                         'e.g. `:label`
    TOKEN_PROPERTY_USE
    TOKEN_PROPERTY_NEW
    TOKEN_RAM                           'e.g. `$.ram`
    TOKEN_SECTION                       'e.g. `::section`
    TOKEN_TEXT                          'e.g. `"..."`
    
    [_TOKEN_LAST]                       'Do not go above 255!
End Enum

'--------------------------------------------------------------------------------------

'A list of system targets. Only the SEGA Master System is supported at the moment, _
 but I will consider supporting other Z80 systems in the future
Public Enum OZ80_SYSTEM
    SYSTEM_NONE                         'System not yet defined
    SYSTEM_SMS                          'SEGA Master System
End Enum

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'For logging, we will want to get a text representation of any of the Tokens _
 (oh how I wish VB6 supported static constant arrays)
Private My_TokenName(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String
'We'll have to populate the above with code, so flag it for preparation
Private My_TokenNameInit As Boolean

'Lookup table of hexadecimal prettyprint, _
 saves repetitive text manipulation when logging
Private My_HexStr8(-1 To &HFF) As String * 2
Private My_HexStr8Init As Boolean
Private My_HexStr16(0 To &HFFFF&) As String * 4
Private My_HexStr16Init As Boolean

'HexStr8 : Get a text-representation of an 8-bit (0-255) number in hexadecimal
'======================================================================================
Public Property Get HexStr8( _
    ByRef Index As Long _
) As String
    'If the lookup array is not ready yet, populate it
    If My_HexStr8Init = 0 Then
        Dim i As Long
        For i = 0 To &HF&:          Let My_HexStr8(i) = "0" & Hex$(i):  Next i
        For i = &H10& To &HFF&:     Let My_HexStr8(i) = Hex$(i):        Next i
        Let My_HexStr8(-1) = "··"
        Let My_HexStr8Init = True
    End If
    
    Let HexStr8 = My_HexStr8(Index)
End Property

'HexStr16 : Get a text-representation of a 16-bit (0-65535) number in hexadecimal
'======================================================================================
Public Property Get HexStr16( _
    ByRef Index As Long _
) As String
    'If the lookup array is not ready yet, populate it
    If My_HexStr16Init = 0 Then
        Dim i As Long
        For i = 0 To &HF&:          Let My_HexStr16(i) = "000" & Hex$(i):   Next i
        For i = &H10& To &HFF&:     Let My_HexStr16(i) = "00" & Hex$(i):    Next i
        For i = &H100& To &HFFF&:   Let My_HexStr16(i) = "0" & Hex$(i):     Next i
        For i = &H1000& To &HFFFF&: Let My_HexStr16(i) = Hex$(i):           Next i
        Let My_HexStr16Init = True
    End If
    
    Let HexStr16 = My_HexStr16(Index And &HFFFF&)
End Property

'TokenName : Get the string representation of a token number
'======================================================================================
Public Property Get TokenName( _
    ByRef Token As OZ80_TOKEN _
) As String
    'If the lookup array is not ready yet, populate it
    If My_TokenNameInit = 0 Then
        Let My_TokenNameInit = True
        'Z80 Instructions .............................................................
        Let My_TokenName(TOKEN_Z80_ADC) = "ADC"
        Let My_TokenName(TOKEN_Z80_ADD) = "ADD"
        Let My_TokenName(TOKEN_Z80_AND) = "AND"
        Let My_TokenName(TOKEN_Z80_BIT) = "BIT"
        Let My_TokenName(TOKEN_Z80_CALL) = "CALL"
        Let My_TokenName(TOKEN_Z80_CCF) = "CCF"
        Let My_TokenName(TOKEN_Z80_CP) = "CP"
        Let My_TokenName(TOKEN_Z80_CPD) = "CPD"
        Let My_TokenName(TOKEN_Z80_CPDR) = "CPDR"
        Let My_TokenName(TOKEN_Z80_CPI) = "CPI"
        Let My_TokenName(TOKEN_Z80_CPIR) = "CPIR"
        Let My_TokenName(TOKEN_Z80_CPL) = "CPL"
        Let My_TokenName(TOKEN_Z80_DAA) = "DAA"
        Let My_TokenName(TOKEN_Z80_DEC) = "DEC"
        Let My_TokenName(TOKEN_Z80_DI) = "DI"
        Let My_TokenName(TOKEN_Z80_DJNZ) = "DJNZ"
        Let My_TokenName(TOKEN_Z80_EI) = "EI"
        Let My_TokenName(TOKEN_Z80_EX) = "EX"
        Let My_TokenName(TOKEN_Z80_EXX) = "EXX"
        Let My_TokenName(TOKEN_Z80_HALT) = "HALT"
        Let My_TokenName(TOKEN_Z80_IM) = "IM"
        Let My_TokenName(TOKEN_Z80_IN) = "IN"
        Let My_TokenName(TOKEN_Z80_INC) = "INC"
        Let My_TokenName(TOKEN_Z80_IND) = "IND"
        Let My_TokenName(TOKEN_Z80_INDR) = "INDR"
        Let My_TokenName(TOKEN_Z80_INI) = "INI"
        Let My_TokenName(TOKEN_Z80_INIR) = "INIR"
        Let My_TokenName(TOKEN_Z80_JP) = "JP"
        Let My_TokenName(TOKEN_Z80_JR) = "JR"
        Let My_TokenName(TOKEN_Z80_LD) = "LD"
        Let My_TokenName(TOKEN_Z80_LDD) = "LDD"
        Let My_TokenName(TOKEN_Z80_LDDR) = "LDDR"
        Let My_TokenName(TOKEN_Z80_LDI) = "LDI"
        Let My_TokenName(TOKEN_Z80_LDIR) = "LDIR"
        Let My_TokenName(TOKEN_Z80_NEG) = "NEG"
        Let My_TokenName(TOKEN_Z80_NOP) = "NOP"
        Let My_TokenName(TOKEN_Z80_OR) = "OR"
        Let My_TokenName(TOKEN_Z80_OUT) = "OUT"
        Let My_TokenName(TOKEN_Z80_OUTD) = "OUTD"
        Let My_TokenName(TOKEN_Z80_OTDR) = "OTDR"
        Let My_TokenName(TOKEN_Z80_OUTI) = "OUTI"
        Let My_TokenName(TOKEN_Z80_OTIR) = "OTIR"
        Let My_TokenName(TOKEN_Z80_POP) = "POP"
        Let My_TokenName(TOKEN_Z80_PUSH) = "PUSH"
        Let My_TokenName(TOKEN_Z80_RES) = "RES"
        Let My_TokenName(TOKEN_Z80_RET) = "RET"
        Let My_TokenName(TOKEN_Z80_RETI) = "RETI"
        Let My_TokenName(TOKEN_Z80_RETN) = "RETN"
        Let My_TokenName(TOKEN_Z80_RLA) = "RLA"
        Let My_TokenName(TOKEN_Z80_RL) = "RL"
        Let My_TokenName(TOKEN_Z80_RLCA) = "RLCA"
        Let My_TokenName(TOKEN_Z80_RLC) = "RLC"
        Let My_TokenName(TOKEN_Z80_RLD) = "RLD"
        Let My_TokenName(TOKEN_Z80_RRA) = "RRA"
        Let My_TokenName(TOKEN_Z80_RR) = "RR"
        Let My_TokenName(TOKEN_Z80_RRCA) = "RRCA"
        Let My_TokenName(TOKEN_Z80_RRC) = "RRC"
        Let My_TokenName(TOKEN_Z80_RRD) = "RRD"
        Let My_TokenName(TOKEN_Z80_RST) = "RST"
        Let My_TokenName(TOKEN_Z80_SBC) = "SBC"
        Let My_TokenName(TOKEN_Z80_SCF) = "SCF"
        Let My_TokenName(TOKEN_Z80_SET) = "SET"
        Let My_TokenName(TOKEN_Z80_SLA) = "SLA"
        Let My_TokenName(TOKEN_Z80_SRA) = "SRA"
        Let My_TokenName(TOKEN_Z80_SLL) = "SLL"
        Let My_TokenName(TOKEN_Z80_SRL) = "SRL"
        Let My_TokenName(TOKEN_Z80_SUB) = "SUB"
        Let My_TokenName(TOKEN_Z80_XOR) = "XOR"
        
        'Z80 Registers / Flags ........................................................
        Let My_TokenName(TOKEN_Z80_A) = "A"
        Let My_TokenName(TOKEN_Z80_AF) = "AF"
        Let My_TokenName(TOKEN_Z80_B) = "B"
        Let My_TokenName(TOKEN_Z80_C) = "C"
        Let My_TokenName(TOKEN_Z80_NC) = "NC"
        Let My_TokenName(TOKEN_Z80_BC) = "BC"
        Let My_TokenName(TOKEN_Z80_D) = "D"
        Let My_TokenName(TOKEN_Z80_E) = "E"
        Let My_TokenName(TOKEN_Z80_DE) = "DE"
        Let My_TokenName(TOKEN_Z80_H) = "H"
        Let My_TokenName(TOKEN_Z80_L) = "L"
        Let My_TokenName(TOKEN_Z80_HL) = "HL"
        Let My_TokenName(TOKEN_Z80_I) = "I"
        Let My_TokenName(TOKEN_Z80_IX) = "IX"
        Let My_TokenName(TOKEN_Z80_IXL) = "IXL"
        Let My_TokenName(TOKEN_Z80_IXH) = "IXH"
        Let My_TokenName(TOKEN_Z80_IY) = "IY"
        Let My_TokenName(TOKEN_Z80_IYL) = "IYL"
        Let My_TokenName(TOKEN_Z80_IYH) = "IYH"
        Let My_TokenName(TOKEN_Z80_M) = "M"
        Let My_TokenName(TOKEN_Z80_P) = "P"
        Let My_TokenName(TOKEN_Z80_PC) = "PC"
        Let My_TokenName(TOKEN_Z80_PE) = "PE"
        Let My_TokenName(TOKEN_Z80_PO) = "PO"
        Let My_TokenName(TOKEN_Z80_R) = "R"
        Let My_TokenName(TOKEN_Z80_SP) = "SP"
        Let My_TokenName(TOKEN_Z80_Z) = "Z"
        Let My_TokenName(TOKEN_Z80_NZ) = "NZ"
        
        'Operators ....................................................................
        Let My_TokenName(TOKEN_OPERATOR_ADD) = Chr$(SYNTAX_OPERATOR_ADD)
        Let My_TokenName(TOKEN_OPERATOR_SUB) = Chr$(SYNTAX_OPERATOR_SUB)
        Let My_TokenName(TOKEN_OPERATOR_MUL) = Chr$(SYNTAX_OPERATOR_MUL)
        Let My_TokenName(TOKEN_OPERATOR_DIV) = Chr$(SYNTAX_OPERATOR_DIV)
        Let My_TokenName(TOKEN_OPERATOR_POW) = Chr$(SYNTAX_OPERATOR_POW)
        Let My_TokenName(TOKEN_OPERATOR_MOD) = Chr$(SYNTAX_OPERATOR_MOD)
        Let My_TokenName(TOKEN_OPERATOR_REP) = "x"
        Let My_TokenName(TOKEN_OPERATOR_OR) = Chr$(SYNTAX_OPERATOR_OR)
        Let My_TokenName(TOKEN_OPERATOR_AND) = Chr$(SYNTAX_OPERATOR_AND)
        Let My_TokenName(TOKEN_OPERATOR_XOR) = Chr$(SYNTAX_OPERATOR_XOR)
        
        'Keywords .....................................................................
        Let My_TokenName(TOKEN_KEYWORD_BOOL) = "BOOL"
        Let My_TokenName(TOKEN_KEYWORD_BYTE) = "BYTE"
        Let My_TokenName(TOKEN_KEYWORD_DEF) = "DEF"
        Let My_TokenName(TOKEN_KEYWORD_HASH) = "HASH"
        Let My_TokenName(TOKEN_KEYWORD_HELP) = "HELP"
        Let My_TokenName(TOKEN_KEYWORD_HI) = "HI"
        Let My_TokenName(TOKEN_KEYWORD_INCLUDE) = "INCLUDE"
        Let My_TokenName(TOKEN_KEYWORD_INTERRUPT) = "INTERRUPT"
        Let My_TokenName(TOKEN_KEYWORD_LO) = "LO"
        Let My_TokenName(TOKEN_KEYWORD_LONG) = "LONG"
        Let My_TokenName(TOKEN_KEYWORD_NYBL) = "NYBL"
        Let My_TokenName(TOKEN_KEYWORD_PARAMS) = "PARAMS"
        Let My_TokenName(TOKEN_KEYWORD_PROC) = "PROC"
        Let My_TokenName(TOKEN_KEYWORD_RAM) = "RAM"
        Let My_TokenName(TOKEN_KEYWORD_RETURN) = "RETURN"
        Let My_TokenName(TOKEN_KEYWORD_SECTION) = "SECTION"
        Let My_TokenName(TOKEN_KEYWORD_SLOT) = "SLOT"
        Let My_TokenName(TOKEN_KEYWORD_START) = "START"
        Let My_TokenName(TOKEN_KEYWORD_SYSTEM) = "SYSTEM"
        Let My_TokenName(TOKEN_KEYWORD_TABLE) = "TABLE"
        Let My_TokenName(TOKEN_KEYWORD_TRIP) = "TRIP"
        Let My_TokenName(TOKEN_KEYWORD_WORD) = "WORD"
        
        Let My_TokenName(TOKEN_PREFIX_K) = "K"
        Let My_TokenName(TOKEN_PREFIX_KB) = "KB"
        
        Let My_TokenName(TOKEN_BRACES_OPEN) = Chr$(SYNTAX_BRACES_OPEN)
        Let My_TokenName(TOKEN_BRACES_CLOSE) = Chr$(SYNTAX_BRACES_CLOSE)
        Let My_TokenName(TOKEN_PARENS_OPEN) = Chr$(SYNTAX_PARENS_OPEN)
        Let My_TokenName(TOKEN_PARENS_CLOSE) = Chr$(SYNTAX_PARENS_CLOSE)
        Let My_TokenName(TOKEN_SQUARE_OPEN) = Chr$(SYNTAX_SQUARE_OPEN)
        Let My_TokenName(TOKEN_SQUARE_CLOSE) = Chr$(SYNTAX_SQUARE_CLOSE)
        
        Let My_TokenName(TOKEN_TEXT) = Chr$(SYNTAX_TEXT)
        Let My_TokenName(TOKEN_LABEL) = Chr$(SYNTAX_LABEL)
        Let My_TokenName(TOKEN_PROPERTY_USE) = Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_PROPERTY_NEW) = Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_RAM) = Chr$(SYNTAX_NUMBER_HEX) & Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_SECTION) = String(2, Chr$(SYNTAX_LABEL))
    End If
    
    Let TokenName = My_TokenName(Token)
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'GetOZ80Error : Return an error description for a given error number
'======================================================================================
Public Sub GetOZ80Error( _
    ByRef ErrorNumber As OZ80_ERROR, _
    ByRef ReturnTitle As String, _
    ByRef ReturnDescription As String _
)
    Select Case ErrorNumber
    
    Case OZ80_ERROR_DUPLICATE
        '..............................................................................
        Let ReturnTitle = "Duplicate Definition"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_DUPLICATE_SECTION
        '..............................................................................
        Let ReturnTitle = "Duplicate Definition"
        Let ReturnDescription = _
            "You cannot define a section name twice. There should be only one " & _
            "`SECTION` statement for each section in use."
        
    Case OZ80_ERROR_FILE_END
        '..............................................................................
        Let ReturnTitle = "Unexpected End of File"
        'TODO
        Let ReturnDescription = "The file ended "
        
    Case OZ80_ERROR_EXPECTED
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        'TODO
        Let ReturnDescription = _
            ""
        
    Case OZ80_ERROR_EXPECTED_PROC_NAME
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        Let ReturnDescription = _
            "A label name must follow the `PROC` statement. " & _
            "E.g. `PROC :myProcedure`"
    
    Case OZ80_ERROR_EXPECTED_ROOT
        '..............................................................................
        Let ReturnTitle = "Keyword expected"
        Let ReturnDescription = _
            "Expected `DEF`, `IF`, `INCLUDE`, `PROC, `SECTION`, `SYSTEM` or `TABLE` " & _
            "keywords at this scope. Have you correctly closed any brackets that were " & _
            "open?"
    
    Case OZ80_ERROR_EXPECTED_SECTION_NAME
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        Let ReturnDescription = _
            "A section name must follow the `SECTION` statement. " & _
            "E.g. `SECTION ::graphics`"
    
    Case OZ80_ERROR_EXPECTED_SYSTEM_NAME
        '..............................................................................
        Let ReturnTitle = "Expected Valid System Name"
        Let ReturnDescription = _
            "A valid System name must follow the `SYSTEM` statement. " & _
            "E.g. `SYSTEM ""SMS""`."
        
    Case OZ80_ERROR_EXPRESSION
        '..............................................................................
        Let ReturnTitle = "Invalid Expression"
        Let ReturnDescription = _
            "An expression can be any Number, Label, Property, RAM Name or " & _
            "calculation (via operators) of these."
        
    Case OZ80_ERROR_EXPRESSION_Z80
        '..............................................................................
        Let ReturnTitle = "Invalid Z80 Instruction Parameter"
        Let ReturnDescription = _
            "Parameters following a Z80 instruction must be either a Z80 Register " & _
            "(`a`, `b`, `c` etc.), a Z80 memory expression `[ix+$FF]` or a valid " & _
            "numerical expression, i.e. a calculation, a label name or RAM name."
            
    Case OZ80_ERROR_FILE_NOTFOUND
        '..............................................................................
        Let ReturnTitle = "File Not Found"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_FILE_READ
        '..............................................................................
        Let ReturnTitle = "Cannot Read File"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INDEFINITE
        '..............................................................................
        Let ReturnTitle = "Cannot Use Indefinite Value"
        Let ReturnDescription = _
            "A variable cannot be defined with an indefinite value, that is, " & _
            "an expression containing a yet-unknown value, such as a label. " & _
            "label addresses are not set until after assembly."
    
    Case OZ80_ERROR_INVALID_NAME
        '..............................................................................
        Let ReturnTitle = "Invalid Name"
        Let ReturnDescription = _
            "Variable, label and property names can contain A-Z, 0-9 underscore " & _
            "and dot with the following exceptions: " & _
            "1. the first letter cannot be a number or a dot, " & _
            "2. two dots cannot occur in a row " & _
            "3. a number cannot follow a dot, and " & _
            "4. the name cannot end in a dot" _
    
    Case OZ80_ERROR_INVALID_NAME_RAM
        '..............................................................................
        Let ReturnTitle = "Invalid Name"
        Let ReturnDescription = _
            "RAM names must begin with '$.' and follow standard naming rules " & _
            "beyond that, i.e." & _
            "1. the first letter cannot be a number or a dot " & _
              "(this does not include the dot that follows the dollar sign)" & _
            "2. two dots cannot occur in a row " & _
            "3. a number cannot follow a dot, and " & _
            "4. the name cannot end in a dot" _
        
    Case OZ80_ERROR_INVALID_NUMBER
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_INVALID_NUMBER_DEC
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_INVALID_NUMBER_HEX
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        Let ReturnDescription = _
            "Hexadecimal numbers must begin with '$' and must contain 0-9 & A-F " & _
            "letters only. E.g. `$1234ABCD`"
    
    Case OZ80_ERROR_INVALID_NUMBER_BIN
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INVALID_WORD
        '..............................................................................
        Let ReturnTitle = "Invalid Word"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INVALID_Z80PARAMS
        '..............................................................................
        Let ReturnTitle = "Invalid Parameters For Z80 Instruction"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_OVERFLOW
        '..............................................................................
        Let ReturnTitle = "Overflow"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_UNDEFINED_CONST
        '..............................................................................
        Let ReturnTitle = "Undefined Constant Used"
        Let ReturnDescription = _
            "You've used a Constant name which has not been defined yet. " & _
            "Ensure that early on in your source code you define the Constant " & _
            "Value: " & vbCrLf & vbCrLf & "DEF !CONSTANT 123"
    Case Else
        Stop
    End Select
End Sub

