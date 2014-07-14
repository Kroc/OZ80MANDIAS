Attribute VB_Name = "oz80"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: oz80

'Public, shared stuff

'For speed, we'll be hashing strings into numerical IDs, which both the Assembler _
 and TokenStream classes need to do
Public CRC As New CRC32

'/// ENUMS ////////////////////////////////////////////////////////////////////////////

Public Enum OZ80_LOG
    OZ80_LOG_ACTION                     'The key important happenings
    OZ80_LOG_INFO                       'Optional information, not actions happening
    OZ80_LOG_DEBUG                      'Internal information for debugging purposes
End Enum

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_FILENOTFOUND             'Requested file does not exist
    OZ80_ERROR_FILEREAD                 'Some kind of problem with file handle open
    OZ80_ERROR_INVALIDWORD              'Couldn't parse a word
    OZ80_ERROR_INVALIDNAME              'Invalid label/property/variable name
    OZ80_ERROR_INVALIDNAME_RAM          '- Invalid RAM name, i.e. `$.name`
    OZ80_ERROR_INVALIDNUMBER            'Not a valid binary/hex/decimal number
    OZ80_ERROR_INVALIDNUMBER_DEC        '- Invalid decimal number
    OZ80_ERROR_INVALIDNUMBER_HEX        '- Invalid hexadecimal number
    OZ80_ERROR_INVALIDNUMBER_BIN        '- Invalid binary number
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_EXPRESSION_Z80           '- Not a valid Z80 instruction parameter
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_UNEXPECTED               'Incorrect content at the current scope
    OZ80_ERROR_UNEXPECTED_PROC_NAME     '- A label name must follow `PROC`
    OZ80_ERROR_UNEXPECTED_SECTION_NAME  '- A section name must follow `SECTION`
    OZ80_ERROR_UNEXPECTED_VAR_NAME      '- A variable name must follow `VAR`
    OZ80_ERROR_ENDOFFILE                'Unexpected end of file
    OZ80_ERROR_INDEFINITEVALUE          'Indefinite value cannot be used here
End Enum

'--------------------------------------------------------------------------------------

Public Enum OZ80_TOKEN
    TOKEN_NONE
    
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
    
    'Z80 Registers ....................................................................
    TOKEN_Z80_A                         'Accumulator
    TOKEN_Z80_AF
    TOKEN_Z80_B                         'Register B
    TOKEN_Z80_C                         'Register C and Carry flag
    TOKEN_Z80_NC                        'Carry unset flag
    TOKEN_Z80_BC
    TOKEN_Z80_D
    TOKEN_Z80_E
    TOKEN_Z80_DE
    TOKEN_Z80_H
    TOKEN_Z80_L
    TOKEN_Z80_HL
    TOKEN_Z80_I                         'Interrupt - not to be confused with IX & IY
    TOKEN_Z80_IX
    TOKEN_Z80_IXL
    TOKEN_Z80_IXH
    TOKEN_Z80_IY
    TOKEN_Z80_IYL
    TOKEN_Z80_IYH
    TOKEN_Z80_M                         'Sign is set
    TOKEN_Z80_P                         'Sign is not set
    TOKEN_Z80_PC                        'Program Counter
    TOKEN_Z80_PE                        'Parity/Overflow is set
    TOKEN_Z80_PO                        'Parity/Overflow is not set
    TOKEN_Z80_R                         'Refresh register (pseudo-random)
    TOKEN_Z80_SP                        'Stack Pointer
    TOKEN_Z80_Z                         'Zero set
    TOKEN_Z80_NZ                        'Zero not set
    
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
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
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
    
    [_TOKEN_LAST]                       'Do not go above 256!
End Enum
