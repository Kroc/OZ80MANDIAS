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

Public Enum OZ80_TOKEN
    'This token specifies that the data field is a Z80 mnemonic (`OZ80_MNEMONICS`).
     'Note that this is only a mnemonic token, not the Z80 opcode (handled by the
     'assembler) or parameters (handled by the next tokens in the stream)
    TOKEN_Z80 = &H0
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H1
    
    'Specifies a Z80 register
    TOKEN_REGISTER = &H2
    
    'Specifies an operator
    TOKEN_OPERATOR = &H3
    
    'A list is a series of 1 or more expressions separated by commas,
     'i.e.parameter lists
    TOKEN_LIST = &H4                    'The data field with give the list length
    TOKEN_LIST_NEXT = &H5               'Sandwiched between list items, i.e. ","
    
    'Debug tokens:
    'These allow us to keep track of where the token originated in the source file(s)
     'so that errors in parsing the tokens can print out friendly error messages
    TOKEN_FILE = &HF0                   'specify a change in file name
    TOKEN_LINE = &HF1                   'specify a change in line number
    TOKEN_COL = &HF2                    'specify a change in column number
End Enum

'The data field of a token can be any 32-bit number and not just limited to the defs _
 below. These are provided for tokens that have well defined data values
Public Enum OZ80_TOKEN_DATA
    'Z80 Assembly Mnemonics:
    '----------------------------------------------------------------------------------
    'These are just the mnemonic tokens -- the assembly routine itself checks the
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
    
    'Z80 Registers:
    '----------------------------------------------------------------------------------
    'When the token is OZ80_TOKEN_REGISTER then the following specifies which register
    
    TOKEN_REGISTER_A = &H100000         'Accumulator
    TOKEN_REGISTER_B = &H100001
    TOKEN_REGISTER_C = &H100002
    TOKEN_REGISTER_D = &H100004
    TOKEN_REGISTER_E = &H100008
    TOKEN_REGISTER_F = &H100010         'Flags register
    TOKEN_REGISTER_H = &H100020
    TOKEN_REGISTER_I = &H100040         'Interrupt - not to be confused with IX & IY
    TOKEN_REGISTER_L = &H100080
    TOKEN_REGISTER_R = &H100100         'Refresh register (pseudo-random)
    
    TOKEN_REGISTER_AF = TOKEN_REGISTER_A Or TOKEN_REGISTER_F
    TOKEN_REGISTER_BC = TOKEN_REGISTER_B Or TOKEN_REGISTER_C
    TOKEN_REGISTER_DE = TOKEN_REGISTER_D Or TOKEN_REGISTER_E
    TOKEN_REGISTER_HL = TOKEN_REGISTER_H Or TOKEN_REGISTER_L
    
    'Undocumented Z80 instructions can access the 8-bit halves of IX & IY
    TOKEN_REGISTER_IXL = &H100200
    TOKEN_REGISTER_IXH = &H100201
    TOKEN_REGISTER_IX = TOKEN_REGISTER_IXL Or TOKEN_REGISTER_IXH
    TOKEN_REGISTER_IYL = &H100202
    TOKEN_REGISTER_IYH = &H100204
    TOKEN_REGISTER_IY = TOKEN_REGISTER_IYL Or TOKEN_REGISTER_IYH
    
    TOKEN_REGISTER_SP = &H100300        'Stack pointer
    TOKEN_REGISTER_PC = &H100301        'Program counter
    
    'Operators:
    '----------------------------------------------------------------------------------
    'When then token is OZ80_OPERATOR then the following specifies which opeartor
    
    TOKEN_OPERATOR_ADD = &HF00000       'Add "+"
    TOKEN_OPERATOR_SUB = &HF00001       'Subtract "-"
    TOKEN_OPERATOR_MUL = &HF00002       'Multiply "*"
    TOKEN_OPERATOR_DIV = &HF00002       'Divide "/"
    TOKEN_OPERATOR_POW = &HF00003       'Power "^"
    TOKEN_OPERATOR_MOD = &HF00004       'Modulus "\"
    TOKEN_OPERATOR_REP = &HF00010       'Repeat "x"
End Enum

Private Type Token
    'An `OZ80_TOKEN` value, though not specified as such otherwise it'll use 4 bytes
    Kind As Byte
    Data As OZ80_TOKEN_DATA
End Type

Private Tokens() As Token

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Tokens
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddToken : Add a token to the assembler's internal tokenised code representation _
 ======================================================================================
Public Sub AddToken( _
    ByVal Kind As OZ80_TOKEN, _
    Optional ByVal Data As OZ80_TOKEN_DATA = 0 _
)
    'Add an element to the Tokens array
    Static Dimmed As Boolean
    If Dimmed = False Then
        ReDim Tokens(0) As Token
        Let Dimmed = True
    Else
        ReDim Preserve Tokens(UBound(Tokens) + 1) As Token
    End If
    
    With Tokens(UBound(Tokens))
        Let .Kind = Kind
        Let .Data = Data
    End With
End Sub

'ArrayDimmed : Is an array dimmed? _
 ======================================================================================
'Taken from: https://groups.google.com/forum/?_escaped_fragment_=msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J#!msg/microsoft.public.vb.general.discussion/3CBPw3nMX2s/zCcaO-hiCI0J
Private Function ArrayDimmed(varArray As Variant) As Boolean
    Dim pSA As Long
    'Make sure an array was passed in:
    If IsArray(varArray) Then
        'Get the pointer out of the Variant:
        Call kernel32_RtlMoveMemory( _
            ptrDestination:=pSA, ptrSource:=ByVal VarPtr(varArray) + 8, Length:=4 _
        )
        If pSA Then
            'Try to get the descriptor:
            Call kernel32_RtlMoveMemory( _
                ptrDestination:=pSA, ptrSource:=ByVal pSA, Length:=4 _
            )
            'Array is initialized only if we got the SAFEARRAY descriptor:
            Let ArrayDimmed = (pSA <> 0)
        End If
    End If
End Function
