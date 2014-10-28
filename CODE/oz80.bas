Attribute VB_Name = "oz80"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: oz80

'Public, shared stuff

'For speed, we'll be hashing strings into numerical IDs, _
 which both the Assembler and TokenStream classes need to do
Public CRC As New CRC32

'Some expressions cannot be calculated until the Z80 code has been assembled. _
 For example, Label addresses are chosen after all code has been parsed and the sizes _
 of the chunks are known. A special Value is used that lies outside of the allowable _
 range of numbers in OZ80 (32-bit) to mark an Expression with a yet-unknown Value

'VB does not allow implicit Double (64-bit) values greater than 32-bits, _
 a trick is used here to build the largest possible 64-bit number: _
 <stackoverflow.com/questions/929069/how-do-i-declare-max-double-in-vb6/933490#933490>
Public Const OZ80_INDEFINITE As Double = 1.79769313486231E+308 + 5.88768018655736E+293

'/// PUBLIC ENUMS /////////////////////////////////////////////////////////////////////

'This makes life a whole lot easier when processing text as ASCII codes
Public Enum ASCII
    'Non-visible control codes:
:   ASC_NUL:    ASC_SOH:    ASC_STX:    ASC_ETX:    ASC_EOT:    ASC_ENQ:    ASC_ACK
:   ASC_BEL:    ASC_BS:     ASC_TAB:    ASC_LF:     ASC_VT:     ASC_FF:     ASC_CR
:   ASC_SO:     ASC_SI:     ASC_DLE:    ASC_DC1:    ASC_DC2:    ASC_DC3:    ASC_DC4
:   ASC_NAK:    ASC_SYN:    ASC_ETB:    ASC_CAN:    ASC_EM:     ASC_SUB:    ASC_ESC
:   ASC_FS:     ASC_GS:     ASC_RS:     ASC_US

    ASC_SPC                             '` ` Space
    ASC_EXC                             '`!` Exclamation Mark
    ASC_QUOT                            '`"` Quote
    ASC_HASH                            '`#` Hash / Pound / Octothorpe
    ASC_DOL                             '`$` Dollar
    ASC_PERC                            '`%` Per-Cent
    ASC_AMP                             '`&` Ampersand
    ASC_APOS                            '`'` Single-Quote / Apostrophe
    ASC_LP                              '`(` Left Parenthesis
    ASC_RP                              '`)` Right Parenthesis
    ASC_STAR                            '`*` Asterisk
    ASC_PLUS                            '`+` Plus
    ASC_COM                             '`,` Comma
    ASC_HYP                             '`-` Hyphen
    ASC_DOT                             '`.` Dot
    ASC_FSL                             '`/` Forward-Slash

:   ASC_0:      ASC_1:      ASC_2:      ASC_3:      ASC_4:      ASC_5:      ASC_6
:   ASC_7:      ASC_8:      ASC_9
    
    ASC_COL                             '`:` Colon
    ASC_SCOL                            '`;` Semi-Colon
    ASC_LT                              '`<` Less-Than
    ASC_EQ                              '`=` Equals
    ASC_GT                              '`>` Greater-Than
    ASC_QM                              '`?` Question Mark
    ASC_AT                              '`@` At-Mark
    
:   ASC_A:      ASC_B:      ASC_C:      ASC_D:      ASC_E:      ASC_F:      ASC_G
:   ASC_H:      ASC_I:      ASC_J:      ASC_K:      ASC_L:      ASC_M:      ASC_N
:   ASC_O:      ASC_P:      ASC_Q:      ASC_R:      ASC_S:      ASC_T:      ASC_U
:   ASC_V:      ASC_W:      ASC_X:      ASC_Y:      ASC_Z
    
    ASC_LSB                             '`[` Left Square-Bracket
    ASC_BSL                             '`\` Back-Slash
    ASC_RSB                             '`]` Right Square-Bracket
    ASC_CRT                             '`^` Caret / Circumflex
    ASC_USC                             '`_` Underscore
    ASC_BTK                             '``` Backtick / Grave Accent

    'Lower-case letters
:   ASC_a_:     ASC_b_:     ASC_c_:     ASC_d_:     ASC_e_:     ASC_f_:     ASC_g_
:   ASC_h_:     ASC_i_:     ASC_j_:     ASC_k_:     ASC_l_:     ASC_m_:     ASC_n_
:   ASC_o_:     ASC_p_:     ASC_q_:     ASC_r_:     ASC_s_:     ASC_t_:     ASC_u_
:   ASC_v_:     ASC_w_:     ASC_x_:     ASC_y_:     ASC_z_:

    ASC_LB                              '`{` Left Brace
    ASC_VB                              '`|` Vertical Bar / Pipe
    ASC_RB                              '`}` Right Brace
    ASC_TIL                             '`~` Tilde
    
    ASC_DEL                             '"Delete" -- non-visible
End Enum

'--------------------------------------------------------------------------------------

'These define the various punctiation marks (in ASCII codes) for the language syntax
Public Enum OZ80_SYNTAX
    SYNTAX_COMMENT = ASC_BTK            ' ` - Comment marker. "``" for multi-line
    SYNTAX_HINT1 = ASC_SCOL             ' ; - register hint, e.g. `a;index`
    SYNTAX_HINT2 = ASC_APOS             ' ' - shadow register hint, e.g. `ex af 'af`
    SYNTAX_QUOTE = ASC_QUOT             ' " - string identifier
    SYNTAX_LABEL = ASC_COL              ' : - label identifier
    SYNTAX_PROPERTY = ASC_DOT           ' . - property identifier
    SYNTAX_OBJECT = ASC_HASH            ' # - object identifier
    SYNTAX_RAM = ASC_DOL                ' $ - RAM constant identifier -- "$.abc"
    SYNTAX_MACRO = ASC_AT               ' @ - macro identifier
    SYNTAX_FUNCT = ASC_QM               ' ? - function identifier
    SYNTAX_NUMBER_HEX = ASC_DOL         ' $ - hexadecimal number, e.g. `$FFFF`
    SYNTAX_NUMBER_BIN = ASC_PERC        ' % - binary number, e.g. `%10101011`
    SYNTAX_NEXT = ASC_COM               ' , - item seperator, optional
    SYNTAX_PAREN_OPEN = ASC_LP          ' ( - memory reference open parenthesis
    SYNTAX_PAREN_CLOSE = ASC_RP         ' ) - memory reference close parenthesis
    SYNTAX_CHUNK_OPEN = ASC_LB          ' { - open brace
    SYNTAX_CHUNK_CLOSE = ASC_RB         ' } - close brace
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

'--------------------------------------------------------------------------------------

Public Enum OZ80_LOG
    OZ80_LOG_ACTION                     'The key important happenings
    OZ80_LOG_INFO                       'Optional information, not actions happening
    OZ80_LOG_STATUS                     'Display variable values &c. when assigned
    OZ80_LOG_DEBUG                      'Internal information for debugging purposes
End Enum

Public Enum OZ80_WARNING
    OZ80_WARNING_NONE                   'Skip "0"
    
End Enum

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_DUPLICATE_PROC_INTERRUPT '- Duplicate `INTERRUPT` parameter
    OZ80_ERROR_DUPLICATE_PROC_PARAMS    '- Duplicate `PARAMS` parameter
    OZ80_ERROR_DUPLICATE_PROC_RETURN    '- Duplicate `RETURN` parameter
    OZ80_ERROR_DUPLICATE_PROC_SECTION   '- Duplicate `SECTION` parameter
    OZ80_ERROR_DUPLICATE_SECTION        '- Can't define a section twice
    OZ80_ERROR_EXPECTED                 'Incorrect content at the current scope
    OZ80_ERROR_EXPECTED_PROC_NAME       '- A label name must follow `PROC`
    OZ80_ERROR_EXPECTED_PROC_PARAMS     '- Invalid stuff in the `PARAMS` list
    OZ80_ERROR_EXPECTED_PROC_RETURN     '- Invalid stuff in the `RETURN` list
    OZ80_ERROR_EXPECTED_ROOT            '- Only certain keywords allowed at root
    OZ80_ERROR_EXPECTED_SECTION_NAME    '- A section name must follow `SECTION`
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_EXPRESSION_Z80           '- Not a valid Z80 instruction parameter
    OZ80_ERROR_FILE_END                 'Unexpected end of file
    OZ80_ERROR_FILE_NOTFOUND            'Requested file does not exist
    OZ80_ERROR_FILE_READ                'Some kind of problem while file handle open
    OZ80_ERROR_INDEFINITE               'Indefinite value cannot be used here
    OZ80_ERROR_INVALID_INTERRUPT        'Invalid Interrupt address
    OZ80_ERROR_INVALID_NAME             'Invalid label/property/variable name
    OZ80_ERROR_INVALID_NAME_RAM         '- Invalid RAM name, i.e. `$.name`
    OZ80_ERROR_INVALID_NUMBER           'Not a valid binary/hex/decimal number
    OZ80_ERROR_INVALID_NUMBER_DEC       '- Invalid decimal number
    OZ80_ERROR_INVALID_NUMBER_HEX       '- Invalid hexadecimal number
    OZ80_ERROR_INVALID_NUMBER_BIN       '- Invalid binary number
    OZ80_ERROR_INVALID_SECTION          'Section used, but not defined
    OZ80_ERROR_INVALID_SLOT             'Incorrect use of the Slot parameter
    OZ80_ERROR_INVALID_WORD             'Couldn't parse a word
    OZ80_ERROR_INVALID_Z80PARAMS        'Not the right parameters for a Z80 instruction
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_OVERFLOW_Z80             '16-bit number used with an 8-bit instruction
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_TOKEN
    TOKEN_NONE                          'Skip "0"
    
    'These are just the mnemonic tokens -- the assembler checks the
     'parameters and determines which opcode should be used
    [_TOKEN_INSTRUCTIONS_BEGIN]
    TOKEN_Z80_ADC                       'Add with Carry
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
    [_TOKEN_INSTRUCTIONS_END]
    
    'Z80 Registers & Flags ............................................................
    [_TOKEN_REGISTERS_BEGIN]
    TOKEN_Z80_A                         'Accumulator
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
    [_TOKEN_REGISTERS_END]
    
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
    TOKEN_KEYWORD_INTERRUPT             'Interrupt `PROC :<label> INTERRUPT <expr>`
    TOKEN_KEYWORD_PARAMS                'Parameter list `PROC :<label> PARAMS <list>`
    TOKEN_KEYWORD_PROC                  'Procedure Chunk `PROC :<label> { ... }`
    TOKEN_KEYWORD_RETURN                'Returns list `PROC :<label> RETURN <list>`
    TOKEN_KEYWORD_SECTION               'Section definition `SECTION ::<section>
    TOKEN_KEYWORD_SLOT                  'Section Slot pattern `SLOT 0, 1, 2`
    [_TOKEN_KEYWORDS_END]
    
    TOKEN_NUMBER
    'Number prefixes ("K", "KB" & "Kbit")
    TOKEN_PREFIX_K                      'x1000
    TOKEN_PREFIX_KB                     'x1024
    TOKEN_PREFIX_KBIT                   'x128 (1024 bits)
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN
    TOKEN_PARENCLOSE
    TOKEN_CHUNKOPEN
    TOKEN_CHUNKCLOSE
    
    TOKEN_QUOTE                         'e.g. `"..."`
    TOKEN_LABEL                         'e.g. `:label`
    TOKEN_SECTION                       'e.g. `::section`
    TOKEN_PROPERTY_USE
    TOKEN_PROPERTY_NEW
    TOKEN_RAM                           'e.g. `$.ram`
    
    [_TOKEN_LAST]                       'Do not go above 255!
End Enum

'--------------------------------------------------------------------------------------

'In order to compare the hundreds of permutations of parameters for Z80 instructions, _
 we assign each parameter type a single bit. We can thus check very quickly if a given _
 parameter falls within an allowed list of accepted types

Public Enum OZ80_MASK
    
    [_MASK_REGS_BEGIN] = 1
    MASK_REG_B = 2 ^ 0
    MASK_REG_C = 2 ^ 1
    MASK_REG_D = 2 ^ 2
    MASK_REG_E = 2 ^ 3
    MASK_REG_H = 2 ^ 4
    MASK_REG_L = 2 ^ 5
    MASK_MEM_HL = 2 ^ 6
    MASK_REG_A = 2 ^ 7
    
    'The presence of an IX/IY prefix on the opcode changes H/L to IXH/IYH or IXL/IYL
     'respectively, but only on instructions that use single byte opcodes.
     'This is officially undocumented, but obviously fair game for old systems
    MASK_REG_IXH = 2 ^ 8
    MASK_REG_IXL = 2 ^ 9
    MASK_REG_IYH = 2 ^ 10
    MASK_REG_IYL = 2 ^ 11
    
    'The presence of an IX/IY prefix on the opcode changes a memory reference "(HL)"
     'to IX/IY, with an offset value e.g. "(IX+$8)"
    MASK_MEM_IX = 2 ^ 12
    MASK_MEM_IY = 2 ^ 13
    [_MASK_REGS_END] = MASK_MEM_IY
    
    'A couple of undocumented instructions allow for IX/IY memory references,
     'but not the standard "(HL)" reference
    MASK_MEM_IXY = MASK_MEM_IX Or MASK_MEM_IY
    'And this is the common "(HL|IX+$8|IY+$8)" form that is used often throughout
    MASK_MEM_HLIXY = MASK_MEM_HL Or MASK_MEM_IXY
    
    'The main 8-bit registers are a common instruction parameter
    MASK_REGS_ABCDEHL = MASK_REG_A Or MASK_REG_B Or MASK_REG_C Or MASK_REG_D Or MASK_REG_E Or MASK_REG_E Or MASK_REG_H Or MASK_REG_L
    'The Z80 clumps HL/IX & IY memory references together with 8-bit registers when
     'building opcodes, i.e. "A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)"
    MASK_REGS_ABCDEHL_MEM_HLIXY = MASK_REGS_ABCDEHL Or MASK_MEM_HLIXY
    'The use of the IX/IY prefix turns H/L into IXH/IXL/IYH/IYL in many instances
    MASK_REGS_IXHL = MASK_REG_IXH Or MASK_REG_IXL
    MASK_REGS_IYHL = MASK_REG_IYH Or MASK_REG_IYL
    MASK_REGS_IXYHL = MASK_REGS_IXHL Or MASK_REGS_IYHL
    MASK_REGS_ABCDEIXYHL_MEM_HLIXY = MASK_REGS_ABCDEHL_MEM_HLIXY Or MASK_REGS_IXYHL
    
    'Very uncommon 8-bit registers
    MASK_REG_I = 2 ^ 14                 'Interrupt register
    MASK_REG_R = 2 ^ 15                 'Refresh register, pseudo-random
    
    'The 16-bit register pairs
    MASK_REG_AF = 2 ^ 16                'The Accumulator and the processor Flags
    MASK_REG_BC = 2 ^ 17                'Registers B & C
    MASK_REG_DE = 2 ^ 18                'Registers D & E
    MASK_REG_HL = 2 ^ 19                'Registers H & L
    MASK_REG_SP = 2 ^ 20                'Stack Pointer
    
    MASK_REG_IX = 2 ^ 21
    MASK_REG_IY = 2 ^ 22
    
    MASK_REGS_BC_DE_SP = MASK_REG_BC Or MASK_REG_DE Or MASK_REG_SP
    'Some instructions accept BC/DE/HL/SP, but not IX & IY due to existing prefixes
    MASK_REGS_BC_DE_HL_SP = MASK_REGS_BC_DE_SP Or MASK_REG_HL
    
    'HL, IX & IY are synonymous as they use an opcode prefix to determine which
    MASK_REGS_HL_IXY = MASK_REG_HL Or MASK_REG_IX Or MASK_REG_IY
    'PUSH / POP allow AF but not SP
    MASK_REGS_AF_BC_DE_HL_IXY = MASK_REG_AF Or MASK_REG_BC Or MASK_REG_DE Or MASK_REGS_HL_IXY
    'The LD instruction can take most 16-bit registers
    MASK_REGS_BC_DE_HL_SP_IXY = MASK_REGS_BC_DE_HL_SP Or MASK_REG_IX Or MASK_REG_IY
    
    MASK_VAL = 2 ^ 23
    
    '..................................................................................
    
    'Register C & Flag C cannot be distinguished by the tokeniser (it isn't aware of
     'context) so they are treated as the same thing. Another bit covers NC/Z/NZ so
     'that these are not accidentally taken as Register C elsewhere
    MASK_FLAGS_CZ = MASK_REG_C Or (2 ^ 24)
    MASK_FLAGS_MP = (2 ^ 25)
    
    MASK_FLAGS = MASK_FLAGS_CZ Or MASK_FLAGS_MP
    
    '..................................................................................
    
    'The IN and OUT instructions can use port "C" (which is, in reality, BC)
    MASK_MEM_BC = 2 ^ 26
    MASK_MEM_DE = 2 ^ 27
    MASK_MEM_SP = 2 ^ 28
    
    MASK_MEM_VAL = 2 ^ 29
    
    '..................................................................................
    
    'This is a shorthand to check for any instance of IX/IY so that we can add the
     'relevant opcode prefix with the simplest of tests
    MASK_ANY_IX = MASK_REG_IX Or MASK_REG_IXH Or MASK_REG_IXL Or MASK_MEM_IX
    MASK_ANY_IY = MASK_REG_IY Or MASK_REG_IYH Or MASK_REG_IYL Or MASK_MEM_IY
    MASK_ANY_IXY = MASK_ANY_IX Or MASK_ANY_IY
End Enum

Public Type oz80Param
    Mask As OZ80_MASK
    Token As OZ80_TOKEN
    Value As Long
End Type

'--------------------------------------------------------------------------------------

'A list of system targets. Only the SEGA Master System is supported at the moment, _
 but I will consider supporting other Z80 systems in the future.
Public Enum OZ80_SYSTEM
    SYSTEM_NONE                         'System not yet defined
    SYSTEM_SMS                          'SEGA Master System
End Enum

'--------------------------------------------------------------------------------------

'Whilst in the syntax `SLOT` uses a list (i.e. `SLOT 0, 1, 2`), we convert that into _
 a bit pattern to make it quick and easy to work with instead of iterating an array
Public Enum OZ80_SLOT
    SLOT0 = 2 ^ 0
    SLOT1 = 2 ^ 1
    SLOT2 = 2 ^ 2
End Enum

'/// PUBLIC PROPERTIES ////////////////////////////////////////////////////////////////

'For logging, we will want to get a text representation of any of the Tokens _
 (oh how I wish VB6 supported static constant arrays)
Private My_TokenName(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String
'We'll have to populate the above with code, so flag it for preparation
Private My_TokenNameInit As Boolean

'Lookup table of hexadecimal prettyprint, _
 saves repetitive text manipulation when logging
Private My_HexStr8(0 To &HFF) As String * 2
Private My_HexStr8Init As Boolean
Private My_HexStr16(0 To &HFFFF&) As String * 4
Private My_HexStr16Init As Boolean

'GET HexStr8 : Get a text-representation of an 8-bit (0-255) number in hexadecimal _
 ======================================================================================
Public Property Get HexStr8(ByRef Index As Long) As String
    'If the lookup array is not ready yet, populate it
    If My_HexStr8Init = 0 Then
        Dim i As Long
        For i = 0 To &HF&:          Let My_HexStr8(i) = "0" & Hex$(i):  Next i
        For i = &H10& To &HFF&:     Let My_HexStr8(i) = Hex$(i):        Next i
        Let My_HexStr8Init = True
    End If
    
    Let HexStr8 = My_HexStr8(Index And &HFF&)
End Property

'GET HexStr16 : Get a text-representation of a 16-bit (0-65535) number in hexadecimal _
 ======================================================================================
Public Property Get HexStr16(ByRef Index As Long) As String
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

'GET TokenName : Get the string representation of a token number _
 ======================================================================================
Public Property Get TokenName(ByRef Token As OZ80_TOKEN) As String
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
        Let My_TokenName(TOKEN_OPERATOR_NOT) = Chr$(SYNTAX_OPERATOR_NOT)
        Let My_TokenName(TOKEN_OPERATOR_XOR) = Chr$(SYNTAX_OPERATOR_XOR)
        
        'Keywords .....................................................................
        Let My_TokenName(TOKEN_KEYWORD_INTERRUPT) = "INTERRUPT"
        Let My_TokenName(TOKEN_KEYWORD_PARAMS) = "PARAMS"
        Let My_TokenName(TOKEN_KEYWORD_PROC) = "PROC"
        Let My_TokenName(TOKEN_KEYWORD_RETURN) = "RETURN"
        Let My_TokenName(TOKEN_KEYWORD_SECTION) = "SECTION"
        Let My_TokenName(TOKEN_KEYWORD_SLOT) = "SLOT"
        
        Let My_TokenName(TOKEN_PREFIX_K) = "K"
        Let My_TokenName(TOKEN_PREFIX_KB) = "KB"
        Let My_TokenName(TOKEN_PREFIX_KBIT) = "KBIT"
        
        Let My_TokenName(TOKEN_PARENOPEN) = Chr$(SYNTAX_PAREN_OPEN)
        Let My_TokenName(TOKEN_PARENCLOSE) = Chr$(SYNTAX_PAREN_CLOSE)
        Let My_TokenName(TOKEN_CHUNKOPEN) = Chr$(SYNTAX_CHUNK_OPEN)
        Let My_TokenName(TOKEN_CHUNKCLOSE) = Chr$(SYNTAX_CHUNK_CLOSE)
        
        Let My_TokenName(TOKEN_QUOTE) = Chr$(SYNTAX_QUOTE)
        Let My_TokenName(TOKEN_LABEL) = Chr$(SYNTAX_LABEL)
        Let My_TokenName(TOKEN_PROPERTY_USE) = Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_PROPERTY_NEW) = Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_RAM) = Chr$(SYNTAX_NUMBER_HEX) & Chr$(SYNTAX_PROPERTY)
        Let My_TokenName(TOKEN_SECTION) = String(2, Chr$(SYNTAX_LABEL))
    End If
    
    Let TokenName = My_TokenName(Token)
End Property

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'GetOZ80Error : Return an error description for a given error number _
 ======================================================================================
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
            "Only the keywords `INCLUDE`, `OBJECT`, `PROC`, `SECTION`, `STRUCT` " & _
            "`TABLE` & `VAR` are allowed at this scope."
    
    Case OZ80_ERROR_EXPECTED_SECTION_NAME
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        Let ReturnDescription = _
            "A section name must follow the `SECTION` statement. " & _
            "E.g. `SECTION ::graphics`"
            
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
            "(`a`, `b`, `c` etc.), a Z80 memory expression `(ix+$FF)` or a valid " & _
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
        
    Case Else
        Stop
    End Select
End Sub

'ParamToString : Get a textual representation of a z80 instruction parameter _
 ======================================================================================
Public Function ParamToString( _
    ByRef Param As oz80Param _
) As String
    If Param.Mask = MASK_REG_A Then
        Let ParamToString = "A"
    ElseIf Param.Mask = MASK_REG_B Then Let ParamToString = "B"
    ElseIf Param.Mask = MASK_REG_C Then Let ParamToString = "C"
    ElseIf Param.Mask = MASK_REG_D Then Let ParamToString = "D"
    ElseIf Param.Mask = MASK_REG_E Then Let ParamToString = "E"
    ElseIf Param.Mask = MASK_REG_H Then Let ParamToString = "H"
    ElseIf Param.Mask = MASK_REG_L Then Let ParamToString = "L"
    ElseIf Param.Mask = MASK_REG_I Then Let ParamToString = "I"
    ElseIf Param.Mask = MASK_REG_R Then Let ParamToString = "R"
    ElseIf Param.Mask = MASK_REG_AF Then Let ParamToString = "AF"
    ElseIf Param.Mask = MASK_REG_BC Then Let ParamToString = "BC"
    ElseIf Param.Mask = MASK_REG_DE Then Let ParamToString = "DE"
    ElseIf Param.Mask = MASK_REG_HL Then Let ParamToString = "HL"
    ElseIf Param.Mask = MASK_REG_SP Then Let ParamToString = "SP"
    ElseIf Param.Mask = MASK_REG_IX Then Let ParamToString = "IX"
    ElseIf Param.Mask = MASK_REG_IXL Then Let ParamToString = "IXL"
    ElseIf Param.Mask = MASK_REG_IXH Then Let ParamToString = "IXH"
    ElseIf Param.Mask = MASK_REG_IY Then Let ParamToString = "IY"
    ElseIf Param.Mask = MASK_REG_IYL Then Let ParamToString = "IYL"
    ElseIf Param.Mask = MASK_REG_IYH Then Let ParamToString = "IYH"
    
    ElseIf Param.Mask = MASK_VAL Then
        If Param.Value > 255 Then
            Let ParamToString = "$" & oz80.HexStr16(Param.Value)
        Else
            Let ParamToString = "$" & oz80.HexStr8(Param.Value)
        End If
        
    'The mask bits do not specify every flag, _
     we refer to the token kind for that
    ElseIf (Param.Mask And MASK_FLAGS) <> 0 Then
        If Param.Token = TOKEN_Z80_C Then
            Let ParamToString = "C"
        ElseIf Param.Token = TOKEN_Z80_NC Then Let ParamToString = "NC"
        ElseIf Param.Token = TOKEN_Z80_Z Then Let ParamToString = "Z"
        ElseIf Param.Token = TOKEN_Z80_NZ Then Let ParamToString = "NZ"
        ElseIf Param.Token = TOKEN_Z80_P Then Let ParamToString = "P"
        ElseIf Param.Token = TOKEN_Z80_PE Then Let ParamToString = "PE"
        ElseIf Param.Token = TOKEN_Z80_PO Then Let ParamToString = "PO"
        ElseIf Param.Token = TOKEN_Z80_M Then Let ParamToString = "M"
        End If
    
    'Memory references
    ElseIf Param.Mask = MASK_MEM_HL Then
        Let ParamToString = "(HL)"
    ElseIf Param.Mask = MASK_MEM_IX Then
        Let ParamToString = "(IX+$" & oz80.HexStr8(Param.Value) & ")"
    ElseIf Param.Mask = MASK_MEM_IY Then
        Let ParamToString = "(IY+$" & oz80.HexStr8(Param.Value) & ")"
    ElseIf Param.Mask = MASK_MEM_BC Then Let ParamToString = "(BC)"
    ElseIf Param.Mask = MASK_MEM_DE Then Let ParamToString = "(DE)"
    ElseIf Param.Mask = MASK_MEM_SP Then Let ParamToString = "(SP)"
    ElseIf Param.Mask = MASK_MEM_VAL Then
        If Param.Value > 255 Then
            Let ParamToString = "($" & oz80.HexStr16(Param.Value) & ")"
        Else
            Let ParamToString = "($" & oz80.HexStr8(Param.Value) & ")"
        End If
    End If
End Function
