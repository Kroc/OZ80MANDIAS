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

'/// ENUMS ////////////////////////////////////////////////////////////////////////////

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
    OZ80_ERROR_DUPLICATE_PROC_PARAMS    '- Duplicate `PARAMS` parameter
    OZ80_ERROR_DUPLICATE_PROC_RETURN    '- Duplicate `RETURN` parameter
    OZ80_ERROR_DUPLICATE_PROC_SECTION   '- Duplicate `SECTION` parameter
    OZ80_ERROR_DUPLICATE_SECTION        '- Can't define a section twice
    OZ80_ERROR_DUPLICATE_SECTION_BANK   '- Duplicate `BANK` parameter
    OZ80_ERROR_DUPLICATE_SECTION_SLOT   '- Duplicate `SLOT` parameter
    OZ80_ERROR_ENDOFFILE                'Unexpected end of file
    OZ80_ERROR_EXPECTED                 'Incorrect content at the current scope
    OZ80_ERROR_EXPECTED_PROC_NAME       '- A label name must follow `PROC`
    OZ80_ERROR_EXPECTED_PROC_PARAMS     '- Invalid stuff in the `PARAMS` list
    OZ80_ERROR_EXPECTED_PROC_RETURN     '- Invalid stuff in the `RETURN` list
    OZ80_ERROR_EXPECTED_ROOT            '- Only certain keywords allowed at root
    OZ80_ERROR_EXPECTED_SECTION_NAME    '- A section name must follow `SECTION`
    OZ80_ERROR_EXPECTED_VAR_NAME        '- A variable name must follow `VAR`
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_EXPRESSION_Z80           '- Not a valid Z80 instruction parameter
    OZ80_ERROR_FILENOTFOUND             'Requested file does not exist
    OZ80_ERROR_FILEREAD                 'Some kind of problem with file handle open
    OZ80_ERROR_INDEFINITE               'Indefinite value cannot be used here
    OZ80_ERROR_INVALID_NAME             'Invalid label/property/variable name
    OZ80_ERROR_INVALID_NAME_RAM         '- Invalid RAM name, i.e. `$.name`
    OZ80_ERROR_INVALID_NUMBER           'Not a valid binary/hex/decimal number
    OZ80_ERROR_INVALID_NUMBER_DEC       '- Invalid decimal number
    OZ80_ERROR_INVALID_NUMBER_HEX       '- Invalid hexadecimal number
    OZ80_ERROR_INVALID_NUMBER_BIN       '- Invalid binary number
    OZ80_ERROR_INVALID_SECTION          'Section used, but not defined
    OZ80_ERROR_INVALID_WORD             'Couldn't parse a word
    OZ80_ERROR_INVALID_Z80PARAMS        'Not the right parameters for a Z80 instruction
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
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
    TOKEN_KEYWORD_AT
    TOKEN_KEYWORD_AS
    TOKEN_KEYWORD_BANK
    TOKEN_KEYWORD_BINARY
    TOKEN_KEYWORD_BYTE
    TOKEN_KEYWORD_DATA
    TOKEN_KEYWORD_DEFAULT
    TOKEN_KEYWORD_ECHO
    TOKEN_KEYWORD_ELSE
    TOKEN_KEYWORD_EXISTS
    TOKEN_KEYWORD_FAIL
    TOKEN_KEYWORD_FILL
    TOKEN_KEYWORD_IF
    TOKEN_KEYWORD_INCLUDE
    TOKEN_KEYWORD_LENGTH
    TOKEN_KEYWORD_OBJECT
    TOKEN_KEYWORD_PARAMS
    TOKEN_KEYWORD_PROC
    TOKEN_KEYWORD_RAM
    TOKEN_KEYWORD_RETURN
    TOKEN_KEYWORD_SECTION
    TOKEN_KEYWORD_SLOT
    TOKEN_KEYWORD_START
    TOKEN_KEYWORD_STOP
    TOKEN_KEYWORD_STRUCT
    TOKEN_KEYWORD_TABLE
    TOKEN_KEYWORD_VAR
    TOKEN_KEYWORD_WORD
    [_TOKEN_KEYWORDS_END]
    
    TOKEN_NUMBER
    'Number prefixes ("K", "KB" & "Kbit")
    TOKEN_PREFIX_K                      'x1000
    TOKEN_PREFIX_KB                     'x1024
    TOKEN_PREFIX_KBIT                   'x128 (1024 bits)
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN
    TOKEN_PARENCLOSE
    TOKEN_BLOCKOPEN
    TOKEN_BLOCKCLOSE
    
    TOKEN_QUOTE
    TOKEN_LABEL                         'e.g. `:myProc`
    TOKEN_SECTION                       'e.g. `::section`
    TOKEN_PROPERTY_USE
    TOKEN_PROPERTY_NEW
    TOKEN_VARIABLE                      'e.g. `#myVar`
    TOKEN_RAM                           'e.g. `$.thing`
    
    [_TOKEN_LAST]                       'Do not go above 255!
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_MASK
    'The 8-bit registers
     '(excluding the undocumented IX/IY halves)
    MASK_REG_A = 2 ^ 0
    MASK_REG_B = 2 ^ 1
    MASK_REG_C = 2 ^ 2
    MASK_REG_D = 2 ^ 3
    MASK_REG_E = 2 ^ 4
    MASK_REG_H = 2 ^ 5
    MASK_REG_L = 2 ^ 6
    
    'The main 8-bit registers are a common instruction parameter
    MASK_REGS_ABCDEHL = MASK_REG_A Or MASK_REG_B Or MASK_REG_C Or MASK_REG_D Or MASK_REG_E Or MASK_REG_E Or MASK_REG_H Or MASK_REG_L
    
    MASK_REG_I = 2 ^ 7                  'Interrupt register
    MASK_REG_R = 2 ^ 8                  'Refresh register, pseudo-random
    
    'The 16-bit register pairs
    MASK_REG_AF = 2 ^ 9                 'The Accumulator and the processor Flags
    MASK_REG_BC = 2 ^ 10                'Registers B & C
    MASK_REG_DE = 2 ^ 11                'Registers D & E
    MASK_REG_HL = 2 ^ 12                'Registers H & L
    MASK_REG_SP = 2 ^ 13                'Stack Pointer
    
    MASK_REG_IX = 2 ^ 14
    MASK_REG_IXL = 2 ^ 15
    MASK_REG_IXH = 2 ^ 16
    MASK_REG_IY = 2 ^ 17
    MASK_REG_IYL = 2 ^ 18
    MASK_REG_IYH = 2 ^ 19
    
    'HL, IX & IY are synonymous as they use an opcode prefix to determine which
    MASK_REGS_HL_IXY = MASK_REG_HL Or MASK_REG_IX Or MASK_REG_IY
    'Some instructions accept BC/DE/HL/SP, but not IX & IY due to existing prefixes
    MASK_REGS_BC_DE_HL_SP = MASK_REG_BC Or MASK_REG_DE Or MASK_REG_HL Or MASK_REG_SP
    'PUSH / POP allow AF but not SP
    MASK_REGS_AF_BC_DE_HL = MASK_REG_AF Or MASK_REG_BC Or MASK_REG_DE Or MASK_REG_HL
    'The LD instruction can take most 16-bit registers
    MASK_REGS_BC_DE_HL_SP_IXY = MASK_REGS_BC_DE_HL_SP Or MASK_REG_IX Or MASK_REG_IY
    
    '..................................................................................
    
    [_FLAG_BIT1] = 2 ^ 20
    [_FLAG_BIT2] = 2 ^ 21
    
    'Register C & Flag C cannot be distinguished by the tokeniser (it isn't aware of
     'context), so they are treated as the same thing, which saves a Bit here
    MASK_FLAG_C = MASK_REG_C
    MASK_FLAG_NC = [_FLAG_BIT1]
    MASK_FLAG_Z = [_FLAG_BIT2]
    MASK_FLAG_NZ = [_FLAG_BIT2] Or [_FLAG_BIT1]
    
    MASK_FLAGS_CZ = MASK_FLAG_C Or [_FLAG_BIT2] Or [_FLAG_BIT1]
    
    [_FLAG_BIT3] = 2 ^ 22
    [_FLAG_BIT4] = 2 ^ 23
    [_FLAG_BIT5] = 2 ^ 24
    
    MASK_FLAG_P = [_FLAG_BIT3]
    MASK_FLAG_PE = [_FLAG_BIT4]
    MASK_FLAG_PO = [_FLAG_BIT4] Or [_FLAG_BIT3]
    MASK_FLAG_M = [_FLAG_BIT5]
    
    MASK_FLAGS_MP = [_FLAG_BIT5] Or [_FLAG_BIT4] Or [_FLAG_BIT3]
    
    MASK_FLAGS = MASK_FLAGS_CZ Or MASK_FLAGS_MP
    
    MASK_VAL = 2 ^ 25
    
    '..................................................................................
    
    [_MEM_BIT1] = 2 ^ 26
    [_MEM_BIT2] = 2 ^ 27
    
    MASK_MEM_HL = [_MEM_BIT1]
    MASK_MEM_IX = [_MEM_BIT2]
    MASK_MEM_IY = [_MEM_BIT2] Or [_MEM_BIT1]
    
    MASK_MEMS_HL_IXY = [_MEM_BIT2] Or [_MEM_BIT1]
    
    MASK_REGS_ABCDEHL_MEMS_HL_IXY = MASK_REGS_ABCDEHL Or MASK_MEMS_HL_IXY
    
    'We don't have enough bits left to give all the memory references their own bit,
     'so we combine the existing bits with this extra bit. This is checked specially
     'when comparing parameters to ensure it matches the test
    [_MASK_MEM] = -1                    '= (2 ^ 31)
    
    'The IN and OUT instructions can use port "C" (which is, in reality, BC)
    MASK_MEM_C = [_MASK_MEM] Or MASK_REG_C
    
    MASK_MEM_BC = [_MASK_MEM] Or MASK_REG_BC
    MASK_MEM_DE = [_MASK_MEM] Or MASK_REG_DE
    MASK_MEM_SP = [_MASK_MEM] Or MASK_REG_SP
    
    MASK_MEM_VAL = [_MASK_MEM] Or MASK_VAL
End Enum

'--------------------------------------------------------------------------------------

Public Type oz80Param
    Mask As OZ80_MASK
    Token As OZ80_TOKEN
    Value As Double
End Type

'Some expressions cannot be calculated until the Z80 code has been assembled, _
 for example label addresses are chosen after all code has been parsed and the sizes _
 of the blocks are known. A special value is used that lies outside of the allowable _
 range of numbers in OZ80 (32-bit) to mark an expression with a yet-unknown value

'VB does not allow implicit Double (64-bit) values greater than 32-bits, _
 a trick is used here to build the largest possible 64-bit number: _
 <stackoverflow.com/questions/929069/how-do-i-declare-max-double-in-vb6/933490#933490>
Public Const OZ80_INDEFINITE As Double = 1.79769313486231E+308 + 5.88768018655736E+293

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
    
    Case OZ80_ERROR_DUPLICATE_SECTION_BANK
        '..............................................................................
        Let ReturnTitle = "Duplicate Parameter"
        Let ReturnDescription = _
            "You cannot specify the `BANK` parameter twice in one `SECTION`!"
        
    Case OZ80_ERROR_DUPLICATE_SECTION
        '..............................................................................
        Let ReturnTitle = "Duplicate Definition"
        Let ReturnDescription = _
            "You cannot define a section name twice. There should be only one " & _
            "`SECTION` statement for each section in use."
        
    Case OZ80_ERROR_DUPLICATE_SECTION_SLOT
        '..............................................................................
        Let ReturnTitle = "Duplicate Parameter"
        Let ReturnDescription = _
            "You cannot specify the `SLOT` parameter twice in one `SECTION`!"
        
    Case OZ80_ERROR_ENDOFFILE
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
            
    Case OZ80_ERROR_FILENOTFOUND
        '..............................................................................
        Let ReturnTitle = "File Not Found"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_FILEREAD
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
    
    Case OZ80_ERROR_OVERFLOW
        '..............................................................................
        Let ReturnTitle = "Overflow"
        'TODO
        Let ReturnDescription = ""
        
    Case Else
        Stop
    End Select
End Sub
