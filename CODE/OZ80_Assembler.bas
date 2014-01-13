Attribute VB_Name = "OZ80_Assembler"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: OZ80_Assembler

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Public Declare Sub kernel32_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

Public Enum OZ80_TOKEN
    'Z80 Assembly Mnemonics -----------------------------------------------------------
    'These are just the mnemonic tokens -- the assembly routine itself checks the
     'parameters and determines which opcode should be used
    TOKEN_Z80_ADC = &H1         'Add with Carry
    TOKEN_Z80_ADD = &H2         'Add
    TOKEN_Z80_AND = &H3         'Bitwise AND
    TOKEN_Z80_BIT = &H4         'Bit test
    TOKEN_Z80_CALL = &H5        'Call routine
    TOKEN_Z80_CCF = &H6         'Clear Carry Flag
    TOKEN_Z80_CP = &H7          'Compare
    TOKEN_Z80_CPD = &H8         'Compare and Decrement
    TOKEN_Z80_CPDR = &H9        'Compare, Decrement and Repeat
    TOKEN_Z80_CPI = &HA         'Compare and Increment
    TOKEN_Z80_CPIR = &HB        'Compare, Increment and Repeat
    TOKEN_Z80_CPL = &HC         'Complement (bitwise NOT)
    TOKEN_Z80_DAA = &HD         'Decimal Adjust Accumulator
    TOKEN_Z80_DEC = &HE         'Decrement
    TOKEN_Z80_DI = &HF          'Disable Interrupts
    TOKEN_Z80_DJNZ = &H10       'Decrement and Jump if Not Zero
    TOKEN_Z80_EI = &H11         'Enable Inettupts
    TOKEN_Z80_EX = &H12         'Exchange
    TOKEN_Z80_EXX = &H13        'Exchange shadow registers
    TOKEN_Z80_HALT = &H14       'Stop CPU (wait for interrupt)
    TOKEN_Z80_IM = &H15         'Interrupt Mode
    TOKEN_Z80_IN = &H16         'Input from port
    TOKEN_Z80_INC = &H17        'Increment
    TOKEN_Z80_IND = &H18        'Input and Decrement
    TOKEN_Z80_INDR = &H19       'Input, Decrement and Repeat
    TOKEN_Z80_INI = &H1A        'Input and Increment
    TOKEN_Z80_INIR = &H1B       'Input, Increment and Repeat
    TOKEN_Z80_JP = &H1C         'Jump
    TOKEN_Z80_JR = &H1D         'Jump Relative
    TOKEN_Z80_LD = &H1E         'Load
    TOKEN_Z80_LDD = &H1F        'Load and Decrement
    TOKEN_Z80_LDDR = &H20       'Load, Decrement and Repeat
    TOKEN_Z80_LDI = &H21        'Load and Increment
    TOKEN_Z80_LDIR = &H22       'Load, Increment and Repeat
    TOKEN_Z80_NEG = &H23        'Negate (flip the sign)
    TOKEN_Z80_NOP = &H24        'No Operation (do nothing)
    TOKEN_Z80_OR = &H25         'Bitwise OR
    TOKEN_Z80_OUT = &H26        'Output to port
    TOKEN_Z80_OUTD = &H27       'Output and Decrement
    TOKEN_Z80_OUTDR = &H28      'Output, Decrement and Repeat
    TOKEN_Z80_OUTI = &H29       'Output and Increment
    TOKEN_Z80_OUTIR = &H2A      'Output, Increment and Repeat
    TOKEN_Z80_POP = &H2B        'Pull from stack
    TOKEN_Z80_PUSH = &H2C       'Push onto stack
    TOKEN_Z80_RES = &H2D        'Reset bit
    TOKEN_Z80_RET = &H2E        'Return from routine
    TOKEN_Z80_RETI = &H2F       'Return from Interrupt
    TOKEN_Z80_RETN = &H30       'Return from NMI
    TOKEN_Z80_RLA = &H31        'Rotate Left (Accumulator)
    TOKEN_Z80_RL = &H32         'Rotate Left
    TOKEN_Z80_RLCA = &H33       'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLC = &H34        'Rotate Left Circular
    TOKEN_Z80_RLD = &H35        'Rotate Left 4-bits
    TOKEN_Z80_RRA = &H36        'Rotate Right (Accumulator)
    TOKEN_Z80_RR = &H37         'Rotate Right
    TOKEN_Z80_RRCA = &H38       'Rotate Right Circular (Accumulator)
    TOKEN_Z80_RRC = &H39        'Rotate Right Circular
    TOKEN_Z80_RRD = &H3A        'Rotate Right 4-bits
    TOKEN_Z80_RST = &H3B        '"Restart" -- Call a page 0 routine
    TOKEN_Z80_SBC = &H3C        'Subtract with Carry
    TOKEN_Z80_SCF = &H3D        'Set Carry Flag
    TOKEN_Z80_SET = &H3E        'Set bit
    TOKEN_Z80_SLA = &H3F        'Shift Left Arithmetic
    TOKEN_Z80_SRA = &H40        'Shift Right Arithmetic
    TOKEN_Z80_SLL = &H41        'Shift Left Logical
    TOKEN_Z80_SRL = &H42        'Shift Right Logical
    TOKEN_Z80_SUB = &H43        'Subtract
    TOKEN_Z80_XOR = &H44        'Bitwise XOR
    
    TOKEN_NUMBER = &H45
    TOKEN_LABEL = &H46
End Enum

Private Type Token
    File As Byte
    Line As Long
    Col As Integer
    Kind As Byte
End Type

Private Tokens() As Token

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddToken : Add a token to the assembler's internal tokenised code representation _
 ======================================================================================
Public Sub AddToken( _
    ByVal Kind As Byte, _
    Optional ByVal File As Byte = 0, _
    Optional ByVal Line As Long = -1, _
    Optional ByVal Col As Integer = -1 _
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
