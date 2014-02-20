VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80TokenStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS:: oz80TokenStream

'A Token Stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. The assembler uses _
 the token stream to 1. complete the calculations and 2. convert tokens to assembly

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'/// DEFS /////////////////////////////////////////////////////////////////////////////

Private Data() As Byte                  'The token stream will live here
Private DataLimit As Long               'The length of the above array
Private DataLength As Long              'The current amount of space used in the array
Private DataCursor As Long              'Where the current insertion point is

Public Enum OZ80_TOKEN
    'This token specifies that the data field is a Z80 mnemonic (`OZ80_MNEMONICS`).
     'Note that this is only a mnemonic token, not the Z80 opcode (handled by the
     'assembler) or parameters (handled by the next tokens in the stream)
    TOKEN_Z80 = &H1
    
    'Specifies a Z80 register
    TOKEN_REGISTER = &H2
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H3
    TOKEN_PREFIX = &H4                  'Number prefixes ("K", "KB" & "Kbit")
    
    'Specifies an operator
    TOKEN_OPERATOR = &H5
    
    'A list is a series of 1 or more expressions separated by commas,
     'i.e.parameter lists
    TOKEN_LIST = &H6
    TOKEN_NEXT = &H7                    'Sandwiched between list items, i.e. ","
    
    'Debug tokens:
    'These allow us to keep track of where the token originated in the source file(s)
     'so that errors in parsing the tokens can print out friendly error messages
    TOKEN_FILE = &HF0                   'specify a change in file name
    TOKEN_LINE = &HF1                   'specify a change in line number
    TOKEN_COL = &HF2                    'specify a change in column number
End Enum

'Z80 Assembly Mnemonics: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_Z80
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
End Enum

'Z80 Registers: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_REGISTER
    TOKEN_REGISTER_A = 1                'Accumulator
    TOKEN_REGISTER_B = 2
    TOKEN_REGISTER_C = 3
    TOKEN_REGISTER_D = 4
    TOKEN_REGISTER_E = 5
    TOKEN_REGISTER_F = 6                'Flags register
    TOKEN_REGISTER_H = 7
    TOKEN_REGISTER_I = 8                'Interrupt - not to be confused with IX & IY
    TOKEN_REGISTER_L = 9
    TOKEN_REGISTER_R = 10               'Refresh register (pseudo-random)

    TOKEN_REGISTER_AF = 11
    TOKEN_REGISTER_BC = 12
    TOKEN_REGISTER_DE = 13
    TOKEN_REGISTER_HL = 14
    TOKEN_REGISTER_IX = 17
    TOKEN_REGISTER_IY = 20
    
    'Undocumented Z80 instructions can access the 8-bit halves of IX & IY
    TOKEN_REGISTER_IXL = 15
    TOKEN_REGISTER_IXH = 16
    TOKEN_REGISTER_IYL = 18
    TOKEN_REGISTER_IYH = 19

    TOKEN_REGISTER_SP = 21              'Stack pointer
    TOKEN_REGISTER_PC = 22              'Program counter
End Enum

'Number Prefixes: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_PREFIX
    TOKEN_PREFIX_K = 1                  'x1000
    TOKEN_PREFIX_KB = 2                 'x1024
    TOKEN_PREFIX_KBIT = 3               'x128 (1024 bits)
End Enum

'Operators: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_OPERATOR
    TOKEN_OPERATOR_ADD = 1              'Add "+"
    TOKEN_OPERATOR_SUB = 2              'Subtract "-"
    TOKEN_OPERATOR_MUL = 3              'Multiply "*"
    TOKEN_OPERATOR_DIV = 4              'Divide "/"
    TOKEN_OPERATOR_POW = 5              'Power "^"
    TOKEN_OPERATOR_MOD = 6              'Modulus "\"
    TOKEN_OPERATOR_REP = 7              'Repeat "x"
End Enum


'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Prepare a large buffer (512K) to accept the token stream _
     (using `ReDim` for every token added is *incredibly* slow)
    'If the data overflows this buffer, we'll increase it by _
     another chunk there and then
    Let DataLimit = 524287
    ReDim Data(0 To DataLimit) As Byte
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Data
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddToken_Z80 : Add a Z80 instruction to the stream _
 ======================================================================================
Public Sub AddToken_Z80(ByRef Token As OZ80_TOKEN_Z80)
    Call Add(OZ80_TOKEN.TOKEN_Z80)      'Add a byte to say this is a Z80 instructions
    Call Add(Token)                     'And add the Z80 instruction token
End Sub

'AddToken_Register : Add a Z80 register to the stream _
 ======================================================================================
Public Sub AddToken_Register(ByRef Token As OZ80_TOKEN_REGISTER)
    Call Add(OZ80_TOKEN.TOKEN_REGISTER) 'Add a byte to say this is a Z80 register
    Call Add(Token)                     'And add the register token
End Sub

'AddToken_Operator : Add an operator to the stream _
 ======================================================================================
Public Sub AddToken_Operator(ByRef Token As OZ80_TOKEN_OPERATOR)
    Call Add(OZ80_TOKEN.TOKEN_OPERATOR) 'Add a byte to say this is an OZ80 operator
    Call Add(Token)                     'And add the operator token
End Sub

'AddToken_Prefix : Add a number prefix to the stream _
 ======================================================================================
Public Sub AddToken_Prefix(ByRef Token As OZ80_TOKEN_PREFIX)
    Call Add(OZ80_TOKEN.TOKEN_PREFIX)   'Add a byte to say this is a number prefix
    Call Add(Token)                     'And add the prefix token
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'AddToken_Operator : Add an operator to the stream _
 ======================================================================================
Private Sub Add(ByRef Value As Long)
    'The param MUST be 0-255, but is given as long for speed reasons
    Let Data(DataCursor) = Value
    Let DataCursor = DataCursor + 1
End Sub
