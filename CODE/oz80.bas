Attribute VB_Name = "oz80"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'Module :: oz80

'This is the public interface of OZ80 you should use, it will spawn the classes it _
 needs

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE = 0                 'Assembly completed successfully
    OZ80_ERROR_FILENOTFOUND = 2         'Requested file does not exist
    OZ80_ERROR_FILEREAD = 3             'Some kind of problem with file handle open
    OZ80_ERROR_INVALIDNAME = 4          'Invalid label/property/variable name
    OZ80_ERROR_BADWORD = 5              'Couldn't parse a word
    OZ80_ERROR_BADNUMBER_DEC = 6        'Not a valid decimal number
    OZ80_ERROR_OVERFLOW = 7             'A number overflowed the maximum
    OZ80_ERROR_Z80_PARAMETER = 8        'An unexpected parameter for a z80 instruction
    OZ80_ERROR_OPERAND = 9              'Not a valid operand for an expression
    OZ80_ERROR_EXPRESSION = 10          'Not a valid expression
End Enum

Public Enum OZ80_TOKEN
    TOKEN_Z80 = &H1                     'Z80 instruction
    TOKEN_REGISTER = &H2                'Z80 register
    TOKEN_FLAG = &H2                    'Z80 flag condition (used on JP, CALL &  RET)
                                         '(part of registers, due to shared "C")
    TOKEN_OPERATOR = &H3                'Operator (e.g. "+ - * /")
    TOKEN_KEYWORD = &H4                 'Keyword (IF/DATA/ECHO &c.)
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H70
    'Number prefixes ("K", "KB" & "Kbit")
    TOKEN_PREFIX_K = &H7A               'x1000
    TOKEN_PREFIX_KB = &H7B              'x1024
    TOKEN_PREFIX_KBIT = &H7C            'x128 (1024 bits)
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN = &HD0
    TOKEN_PARENCLOSE = &HD1
    TOKEN_BLOCKOPEN = &HD2
    TOKEN_BLOCKCLOSE = &HD3
    
    TOKEN_QUOTE = &HE0
    
    TOKEN_LABELDEF = &HA0
    TOKEN_LABEL = &HA1
    TOKEN_PROPERTYDEF = &HA2
    TOKEN_PROPERTY = &HA3
    TOKEN_VARIABLE = &HA4
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
    TOKEN_KEYWORD_DEF = &H7
    TOKEN_KEYWORD_DEFAULT = &H8
    TOKEN_KEYWORD_ECHO = &H9
    TOKEN_KEYWORD_ELSE = &HA
    TOKEN_KEYWORD_EXISTS = &HB
    TOKEN_KEYWORD_FAIL = &HC
    TOKEN_KEYWORD_FILL = &HD
    TOKEN_KEYWORD_IF = &HE
    TOKEN_KEYWORD_INCLUDE = &HF
    TOKEN_KEYWORD_LENGTH = &H10
    TOKEN_KEYWORD_OBJECT = &H11
    TOKEN_KEYWORD_PARAMS = &H12
    TOKEN_KEYWORD_PROC = &H13
    TOKEN_KEYWORD_RETURN = &H14
    TOKEN_KEYWORD_SLOT = &H15
    TOKEN_KEYWORD_START = &H16
    TOKEN_KEYWORD_STOP = &H17
    TOKEN_KEYWORD_STRUCT = &H18
    TOKEN_KEYWORD_TABLE = &H19
    TOKEN_KEYWORD_WORD = &H1A
End Enum

Public Type oz80Token
    Kind As Byte                        '=OZ80_TOKEN, but use 1-byte instead of 4
    Data As OZ80_TOKEN_DATA             'Associated value
    Line As Long                        'Line number in the original source text
    Col As Long                         'Column number in the original source text
End Type

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Assemble : Take a source file and produce a binary _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    Dim Tokeniser As oz80Tokeniser
    
    Debug.Print
    Debug.Print "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision
    
    Dim StartTime As Single
    Let StartTime = Timer
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
    'Create a tokeniser object to hold the machine representation of the text files; _
     the assembler doesn't work with the original text directly
    Set Tokeniser = New oz80Tokeniser
    'Explode the source code file into tokens
    Let Assemble = Tokeniser.Tokenise(FilePath)
    If Assemble <> OZ80_ERROR_NONE Then GoTo Finish
    
    'Stage 2: Assemble _
     ----------------------------------------------------------------------------------
    Dim Assembler As oz80Assembler
    Set Assembler = New oz80Assembler
    Let Assemble = Assembler.Process(Tokeniser)
    If Assemble <> OZ80_ERROR_NONE Then GoTo Finish
    
Finish:
    'For any error that occured, it will be assumed that the relevant function has _
     already printed the error since it will have the right information to hand
    Set Assembler = Nothing
    Set Tokeniser = Nothing
End Function
