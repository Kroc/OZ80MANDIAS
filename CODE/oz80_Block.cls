VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_Block"
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
'CLASS :: oz80_Block

'Code and/or data is assembled in separate blocks (this class) which are arranged _
 according to Section layout before being written out to the ROM. _
 i.e. each `PROC` / `TABLE` will be its own block

'This class helps keep the complexities of the opcode mapping private from the _
 assembler

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'Which section the block belongs to. This will be the CRC32 of the section name, _
 as stored in oz80_Assembler. Value will be 0 for no particular section
Public Section As Long

'The absolute position of the block in the ROM. _
 This will be assigned after the first assembly phase when all block sizes are known _
 and they can be placed in their given sections
Public Address As Long

Public Banks As Scripting.Dictionary
Public Slots As Scripting.Dictionary

'The assembled code/data goes here
Private Block() As Byte
Private At As Long

Private My_Size As Long                 'Length of the block in bytes
Private My_CyclesMin As Long            'Minimum number of cycles in this block
Private My_CyclesMax As Long            'Maximum number of cycles in this block

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

'Each block starts at this size (many blocks will be small Z80 routines). _
 If this buffer size is filled, it will be doubled. Once the block is "finalised", _
 the buffer will be cut down to the size of the actual contents
Private Const BLOCK_BUFFER As Long = 1024


Private Enum OZ80_PARAM
    OZ80_PARAM_NOMEM = 2 ^ 0
    OZ80_PARAM_MEM = 2 ^ 1
    
    OZ80_PARAM_A = 2 ^ 2
    OZ80_PARAM_B = 2 ^ 3
    OZ80_PARAM_C = 2 ^ 4
    OZ80_PARAM_D = 2 ^ 5
    OZ80_PARAM_E = 2 ^ 6
    OZ80_PARAM_H = 2 ^ 7
    OZ80_PARAM_L = 2 ^ 8
    OZ80_PARAM_I = 2 ^ 9
    OZ80_PARAM_R = 2 ^ 10
    
    OZ80_PARAM_AF = 2 ^ 11
    OZ80_PARAM_BC = 2 ^ 12
    OZ80_PARAM_DE = 2 ^ 13
    OZ80_PARAM_HL = 2 ^ 14
    OZ80_PARAM_SP = 2 ^ 15
    
    OZ80_PARAM_IX = 2 ^ 16
    OZ80_PARAM_IY = 2 ^ 17
    
    [_BIT9] = 2 ^ 18
    [_BIT10] = 2 ^ 19
    [_BIT11] = 2 ^ 20
    
    OZ80_PARAM_FLAG_C = [_BIT9]
    OZ80_PARAM_FLAG_NC = [_BIT10]
    OZ80_PARAM_FLAG_Z = [_BIT10] Or [_BIT9]
    OZ80_PARAM_FLAG_NZ = [_BIT11]
    
    OZ80_PARAM_FLAGS_CZ = [_BIT11] Or [_BIT10] Or [_BIT9]
    
    [_BIT12] = 2 ^ 21
    [_BIT13] = 2 ^ 22
    [_BIT14] = 2 ^ 23
    
    OZ80_PARAM_FLAG_P = [_BIT12]
    OZ80_PARAM_FLAG_PE = [_BIT13]
    OZ80_PARAM_FLAG_PO = [_BIT13] Or [_BIT12]
    OZ80_PARAM_FLAG_M = [_BIT14]
    
    OZ80_PARAM_FLAGS_MP = [_BIT14] Or [_BIT13] Or [_BIT12]
    
    OZ80_PARAM_FLAGS = OZ80_PARAM_FLAGS_CZ Or OZ80_PARAM_FLAGS_MP
    
    OZ80_PARAM_VAL_8 = 2 ^ 24
    OZ80_PARAM_VAL_16 = 2 ^ 25
End Enum

'--------------------------------------------------------------------------------------

'http://z80.info/decoding.htm
    
' .---------------.
' |7|6|5|4|3|2|1|0|
' '---------------'
' | x |  y  |  z  |
' '---------------'
'     | p |q|
'     '-----'
    
Private Const x1 As Byte = 2 ^ 6
Private Const x2 As Byte = 2 ^ 7
Private Const x3 As Byte = x1 Or x2
    
Private Const y1 As Byte = 2 ^ 3
Private Const y2 As Byte = 2 ^ 4
Private Const y3 As Byte = y1 Or y2
Private Const y4 As Byte = 2 ^ 5
Private Const y5 As Byte = y4 Or y1
Private Const y6 As Byte = y4 Or y2
Private Const y7 As Byte = y4 Or y4
    
Private Const z1 As Byte = 2 ^ 0
Private Const z2 As Byte = 2 ^ 1
Private Const z3 As Byte = y1 Or y2
Private Const z4 As Byte = 2 ^ 2
Private Const z5 As Byte = z4 Or z1
Private Const z6 As Byte = z4 Or z2
Private Const z7 As Byte = z4 Or z4
    
Private Const p1 As Byte = 2 ^ 4
Private Const p2 As Byte = 2 ^ 5
Private Const p3 As Byte = p1 Or p2

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Block
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddByte : Add a byte to the assembled output and move along _
 ======================================================================================
Public Sub AddByte(ByRef Value As Byte)
    'Add the byte to the ROM and move to the next byte
    Let Block(At) = Value
    Let At = At + 1
    
    If At > My_Size Then
        Let My_Size = My_Size * 2
        ReDim Preserve Block(0 To My_Size) As Byte
    End If
End Sub

'AddInstruction : Assemble a Z80 mneomic _
 ======================================================================================
Friend Sub AddInstruction( _
    ByRef Instruction As OZ80_TOKEN, _
    ByRef Param1 As oz80Param, _
    ByRef Param2 As oz80Param _
)
    'OPCODE MAP: _
     Here's a complete list of the Z80 instruction set and the opcodes used for each _
     addressing type.
     
     
    'http://www.smspower.org/Development/InstructionSet
    'http://z80.info/decoding.htm
    
    'NOTE:
    '   +rr     BC = +00
    '           DE = +10
    '           HL = +20 (Also IX, IY)
    '           SP = +30 (Also AF)
    '   rr      IX = DD
    '           IY = FD
    '           HL = rr and ?? is omitted
    '   +f      nz = +00
    '           z  = +08
    '           nc = +10
    '           c  = +18
    '           po = +20
    '           pe = +28
    '           p  = +30
    '           m  = +38
    '   +b      bit * 8
    '   +im     0  = +00
    '           1  = +10
    '           2  = +18
        
    'Z80    Param 1                 Param 2             Opcode              Cycles _
     ----------------------------------------------------------------------------------
    'ADC    a                       a|b|c|d|e|h|l       88+r                4
    '                               $8                  CE ??               7
    '                               (hl|ix+$8|iy+$8)    rr 8E ??            7
    '       hl                      bc|de|hl|sp         ED 4A+rr            15
                                
    'ADD    a                       a|b|c|d|e|h|l       80+r                4
    '                               $8                  C6 ??               7
    '                               (hl|ix+$8|iy+$8)    rr 86 ??            7
    '       hl                      bc|de|hl|sp         09+rr               11
    '       ix                      bc|de|ix|sp         rr 09+rr            11
    '       iy                      bc|de|iy|sp         rr 09+rr            11
    
    'AND    a|b|c|d|e|h|l                               A0+r
    '       $8                                          E6 ??
    '       (hl|ix+$8|iy+$8)                            rr A6 ??
    
    'BIT    0-7                     a|b|c|d|e|h|l       CB 40+r+b
    '                               (hl|ix+$8|iy+$8)    rr CB ?? 40+r+b
    
    'CALL   $16                                         CD ?? ??
    '       c|nc|m|p|z|nz|pe|po     $16                 C4+f
    
    'CCF                                                3F
    
    'CP     a|b|c|d|e|h|l                               B8+r
    '       (hl|ix+$8|iy+$8)                            rr BE ??
    '       $8                                          FE ??
    
    'CPD                                                ED A9
    'CPDR                                               ED B9
    'CPI                                                ED A1
    'CPIR                                               ED B1
    'CPL                                                2F
    'DAA                                                27
    
    'DEC    a|b|c|d|e|h|l                               05+r*8
    '       (hl|ix+$8|iy+$8)                            rr 35 ??
    '       bc|de|hl|sp                                 0B+rr
    '       ix|iy                                       rr 2B
    
    'DI                                                 F3
    
    'DJNZ   $8                                          10 ??
    
    'EI                                                 FB
    
    'EX     (sp)                    hl|ix|iy            rr E3
    '       af                      af                  08
    '       de                      hl                  EB
    
    'EXX                                                D9
    'HALT                                               76
    
    'IM     0-2                                         46+im
    
    'IN     a                       ($8)                DB ??
    '       (c)                                         ED 70
    '       a|b|c|d|e|h|l           (c)                 ED 40+r
    
    'INC    a|b|c|d|e|h|l                               04+r*8
    '       bc|de|hl|sp|ix|iy                           rr 03+rr
    '       (hl|ix+$8|iy+$8)                            rr 34 ??
    
    'IND                                                ED AA
    'INDR                                               ED BA
    'INI                                                ED A2
    'INIR                                               ED B2
    
    'JP     $16                                         C3 ?? ??
    '       (hl|ix|iy)                                  rr E9
    '       c|nc|m|p|z|nz|pe|po     $16                 C2+f ?? ??
    
    'JR     $8                                          18 ??
    '       c|nc|z|nz               $8                  20+f
    
    'LD     i                       a                   ED 47
    '       r                       a                   ED 4F
    '       a                       i                   ED 57
    '                               r                   ED 5F
    '                               a|b|c|d|e|h|l       78+r
    '                               (hl|ix+$8|iy+$8)    rr 7E ??
    '                               $8                  3E ??
    '                               (bc|de)             0A+rr
    '                               ($16)               3A ?? ??
    '       b                       a|b|c|d|e|h|l       40+r
    '                               (hl|ix+$8|iy+$8)    rr 46 ??
    '                               $8                  06 ??
    '       c                       a|b|c|d|e|h|l       48+r
    '                               (hl|ix+$8|iy+$8)    rr 4E ??
    '                               $8                  0E ??
    '       d                       a|b|c|d|e|h|l       50+r
    '                               (hl|ix+$8|iy+$8)    rr 56 ??
    '                               $8                  16 ??
    '       e                       a|b|c|d|e|h|l       58+r
    '                               (hl|ix+$8|iy+$8)    rr 5E ??
    '                               $8                  1E ??
    '       h                       a|b|c|d|e|h|l       60+r
    '                               (hl|ix+$8|iy+$8)    rr 66 ??
    '                               $8                  26 ??
    '       l                       a|b|c|d|e|h|l       68+r
    '                               (hl|ix+$8|iy+$8)    rr 6E ??
    '                               $8                  2E ??
    '       bc                      ($16)               ED 4B ?? ??
    '                               $16                 01+rr ?? ??
    '       de                      ($16)               ED 5B ?? ??
    '                               $16                 01+rr ?? ??
    '       hl                      ($16)               2A ?? ??
    '                               $16                 01+rr ?? ??
    '       sp                      ($16)               ED 7B ?? ??
    '                               hl|ix|iy            rr F9
    '                               $16                 01+rr ?? ??
    '       ix|iy                   ($16)               rr 2A ?? ??
    '                               $16                 rr 21 ?? ??
    '       (hl)                    a|b|c|d|e|h|l       70+r
    '                               $8                  36 ??
    '       (bc)                    a                   02+rr
    '       (de)                    a                   02+rr
    '       ($16)                   a                   32 ?? ??
    '                               bc                  ED 43 ?? ??
    '                               de                  ED 53 ?? ??
    '                               hl|ix|iy            rr 22 ?? ??
    '                               sp                  ED 73 ?? ??
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l       rr 70+r ??
    '                               $8                  rr 36 ?? ??
    
    'LDD                                                ED A8
    'LDDR                                               ED B8
    'LDI                                                ED A0
    'LDIR                                               ED B0
    'NEG                                                ED 44
    'NOP                                                00
    
    'OR     a|b|c|d|e|h|l                               B0+r
    '       (hl|ix+$8|iy+$8)                            rr B6 ??
    '       $n                                          F6 ??
    
    'OUT    ($8)                    a                   D3 ??
    '       (C)                     0                   ED 71
    '                               a|b|c|d|e|h|l       ED 41+r*8
    
    'OUTD                                               ED AB
    'OTDR                                               ED BB
    'OUTI                                               ED A3
    'OTIR                                               ED B3
    
    'POP    af|bc|de|hl|ix|iy                           rr C1+rr
    
    'PUSH   af|bc|de|hl|ix|iy                           rr C5+rr
    
    'RES    0-7                     a|b|c|d|e|h|l       CB 80+r+b
    '                               (hl|ix+$8|iy+$8)    rr CB ?? 86+b
    
    'RET                                                C9
    '       c|nc|m|p|z|nz|pe|po                         C0+f
    
    'RETI                                               ED 4D
    'RETN                                               ED 45
    'RLA                                                17
    
    'RL     a|b|c|d|e|h|l                               CB 10+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 16
    
    'RLCA                                               07
    
    'RLC    a|b|c|d|e|h|l                               CB 00+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 06
    
    'RLD                                                ED 6F
    'RRA                                                1F
    
    'RR     a|b|c|d|e|h|l                               CB 18+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 1E
    
    'RRCA                                               0F
    
    'RRC    a|b|c|d|e|h|l                               CB 08+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 0E
    
    'RRD                                                ED 67
    
    'RST    0|08|10|18|20|28|30|38                      C7+num
    
    'SBC    a|b|c|d|e|h|l                               98+r
    '       (hl|ix+$8|iy+S8)                            rr 9E ??
    '       a                       $8                  DE ??
    '       hl                      bc|de|hl|sp         ED+rr
    
    'SCF                                                37
    
    'SET    0-7                     a|b|c|d|e|h|l       CB C0+b+r
    '                               (hl|ix+$8|iy+$8)    rr CB ?? C6+b+r
    
    'SLA    a|b|c|d|e|h|l                               CB 20+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 26
    
    'SRA    a|b|c|d|e|h|l                               CB 28+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 2E
    
    'SLL    a|b|c|d|e|h|l                               CB 30+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 36
    
    'SRL    a|b|c|d|e|h|l                               CB 38+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 3E
    
    'SUB    a|b|c|d|e|h|l                               90+r
    '       (hl|ix+$8|iy+$8)                            rr 96 ??
    '       $8                                          D6 ??
    
    'XOR    a|b|c|d|e|h|l                               A8+r
    '       (hl|ix+$8|iy+$8)                            rr AE ??
    '       $8                                          EE ??
    
    '----------------------------------------------------------------------------------
    
    'See, BASIC does have a good use!
    On Instruction - ([_TOKEN_INSTRUCTIONS_BEGIN] + 1) GoTo _
        zADC, zADD, zAND, zBIT, zCALL, zCCF, zCP, zCPD, zCPDR, zCPI, zCPIR, zCPL, _
        zDAA, zDEC, zDI, zDJNZ, zEI, zEX, zEXX, zHALT, zIM, zIN, zINC, zIND, zINDR, _
        zINI, zINIR, zJP, zJR, zLD, zLDD, zLDDR, zLDI, zLDIR, zNEG, zNOP, zOR, zOUT, _
        zOUTD, zOTDR, zOUTI, zOTIR, zPOP, zPUSH, zRES, zRET, zRETI, zRETN, zRLA, zRL, _
        zRLCA, zRLC, zRLD, zRRA, zRR, zRRCA, zRRC, zRRD, zRST, zSBC, zSCF, zSET, _
        zSLA, zSRA, zSLL, zSRL, zSUB, zXOR
    Stop
    
    Dim Opcode As Integer
    
zADC: '................................................................................
    'Add with Carry
    
    
    'ADC    a                       a|b|c|d|e|h|l       88+r                4
    '                               (hl|ix+$8|iy+$8)    rr 8E ??            7
    '                               $8                  CE ??               7
    '       hl                      bc|de|hl|sp         ED 4A+rr            15
    
    '`ADC a, ...`
    If Param1.Register = TOKEN_Z80_A Then
        '`ADC a, a|b|c|d|e|h|l|(hl|ix+$8|iy+$8)`
        If IsRegister8Bit_ABCDEHL_Mem_HLIXY8(Param2.Register) Then
            Let Opcode = x2 Or Opcode_Alu(Instruction) Or _
                         Opcode_RegisterZ(Param2.Register)
        
        '`ADC a, $8`
        ElseIf Param1.Register = TOKEN_NUMBER Then
            Let Opcode = x3 Or Opcode_Alu(Instruction) Or z6
        
        Else
            Stop
        End If
    
    '`ADC hl, ...`
    ElseIf Param1.Register = TOKEN_Z80_HL Then
        'TODO: ED prefix
        Let Opcode = x1 Or Opcode_RegisterPair Or q1 Or z2
        
    Else
        Stop
    End If
    GoTo Finish
    
zADD: '................................................................................
    If Param1.Register = TOKEN_Z80_A Then
        If IsRegister8Bit_ABCDEHL_Mem_HLIXY8(Param2.Register) Then
            Let Opcode = x2 Or Opcode_Alu(Instruction) Or _
                         Opcode_RegisterZ(Param2.Register)
        ElseIf Param1.Register = TOKEN_NUMBER Then
            Let Opcode = x3 Or Opcode_Alu(Instruction) Or z6
        Else
            Stop
        End If
    
    'TODO: IX/IY prefix
    ElseIf Param1.Register = TOKEN_Z80_HL Then
        Let Opcode = Opcode_RegisterPair(Register) Or q1 Or z1
    Else
        Stop
    End If
    GoTo Finish
    
zAND: '................................................................................
zBIT: '................................................................................
zCALL:
zCCF:
zCP:
zCPD:
zCPDR:
zCPI:
zCPIR:
zCPL:
zDAA:
zDEC:
zDI:
zDJNZ:
zEI:
zEX:
zEXX:
zHALT:
zIM:
zIN:
zINC:
zIND:
zINDR:
zINI:
zINIR:
zJP:
zJR:
zLD:
zLDD:
zLDDR:
zLDI:
zLDIR:
zNEG:
zNOP:
zOR:
zOUT:
zOUTD:
zOTDR:
zOUTI:
zOTIR:
zPOP:
zPUSH:
zRES:
zRET:
zRETI:
zRETN:
zRLA:
zRL:
zRLCA:
zRLC:
zRLD:
zRRA:
zRR:
zRRCA:
zRRC:
zRRD:
zRST:
zSBC:
zSCF:
zSET:
zSLA:
zSRA:
zSLL:
zSRL:
zSUB:
zXOR:


Finish:
End Sub

'IsRegister8Bit_ABCDEHL : Test for the common parameter of a main 8-bit register _
 ======================================================================================
Private Function IsRegister8Bit_ABCDEHL( _
    ByRef Register As OZ80_TOKEN _
) As Boolean
    '"Or" comparisions are not used for speed reasons, _
     we want to evaluate as little as possible in one go
    If Register = TOKEN_Z80_A Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_B Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_C Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_D Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_E Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_H Then
        Let IsRegister8Bit_ABCDEHL = True
    ElseIf Register = TOKEN_Z80_L Then
        Let IsRegister8Bit_ABCDEHL = True
    End If
End Function

'IsRegister8Bit_ABCDEHL_Mem_HL_IXY8 : Test for common register / memory-reference _
 ======================================================================================
Private Function IsRegister8Bit_ABCDEHL_Mem_HLIXY8( _
    ByRef Register As OZ80_TOKEN _
) As Boolean
    If Register = TOKEN_Z80_A Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_B Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_C Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_D Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_E Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_H Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_L Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_MEM_HL Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_MEM_IX Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    ElseIf Register = TOKEN_Z80_MEM_IY Then
        Let IsRegister8Bit_ABCDEHL_Mem_HLIXY8 = True
    End If
End Function

'IsValue : Test to see if a parameter would class as having an associated value _
 ======================================================================================
Private Function IsValue( _
    ByRef Register As OZ80_TOKEN _
) As Boolean
    If Register = TOKEN_NUMBER Then
        Let IsValue = True
    ElseIf Register = TOKEN_NUMBER Or TOKEN_Z80_MEM Then
        Let IsValue = True
    ElseIf Register = TOKEN_Z80_MEM_IX Then
        Let IsValue = True
    ElseIf Register = TOKEN_Z80_MEM_IY Then
        Let IsValue = True
    End If
End Function

'Opcode_Alu : Converts arithmetic instructions to bits 3, 4 & 5 of opcode _
 ======================================================================================
Private Function Opcode_Alu( _
    ByRef Instruction As OZ80_TOKEN _
) As Integer
    If Instruction = TOKEN_Z80_ADD Then
        Exit Function
    ElseIf Instruction = TOKEN_Z80_ADC Then
        Let Opcode_Alu = y1
    ElseIf Instruction = TOKEN_Z80_SUB Then
        Let Opcode_Alu = y2
    ElseIf Instruction = TOKEN_Z80_SBC Then
        Let Opcode_Alu = y3
    ElseIf Instruction = TOKEN_Z80_AND Then
        Let Opcode_Alu = y4
    ElseIf Instruction = TOKEN_Z80_XOR Then
        Let Opcode_Alu = y5
    ElseIf Instruction = TOKEN_Z80_OR Then
        Let Opcode_Alu = y6
    ElseIf Instruction = TOKEN_Z80_CP Then
        Let Opcode_Alu = y7
    Else
        Stop
    End If
End Function

'Opcode_Condition : Convert a condition flag to bits 3, 4 & 5 of opcode _
 ======================================================================================
Private Function Opcode_Condition( _
    ByRef Condition As OZ80_TOKEN _
) As Integer
    If Condition = TOKEN_Z80_NZ Then
        Exit Function
    ElseIf Condition = TOKEN_Z80_Z Then
        Let Opcode_Condition = y1
    ElseIf Condition = TOKEN_Z80_NC Then
        Let Opcode_Condition = y2
    ElseIf Condition = TOKEN_Z80_C Then
        Let Opcode_Condition = y3
    ElseIf Condition = TOKEN_Z80_PO Then
        Let Opcode_Condition = y4
    ElseIf Condition = TOKEN_Z80_PE Then
        Let Opcode_Condition = y5
    ElseIf Condition = TOKEN_Z80_P Then
        Let Opcode_Condition = y6
    ElseIf Condition = TOKEN_Z80_M Then
        Let Opcode_Condition = y7
    Else
        Stop
    End If
End Function

'Opcode_RegisterPair : Converts 16-bit register tokens to bits 4 & 5 of the opcode _
 ======================================================================================
Private Function Opcode_RegisterPair( _
    ByRef Register As OZ80_TOKEN _
) As Integer
    If Register = TOKEN_Z80_BC Then
        Exit Function
    ElseIf Register = TOKEN_Z80_DE Then
        Let Opcode_RegisterPair = p1
    ElseIf Register = TOKEN_Z80_HL Then
        Let Opcode_RegisterPair = p2
    ElseIf Register = TOKEN_Z80_SP Or Register = TOKEN_Z80_AF Then
        Let Opcode_RegisterPair = p3
    Else
        Stop
    End If
End Function

'Opcode_RegisterY : Sets the Y-bits (3, 4 & 5) of the opcode based on the register _
 ======================================================================================
Private Function Opcode_RegisterY( _
    ByRef Register As OZ80_TOKEN _
) As Integer
    If Register = TOKEN_Z80_B Then
        Exit Function
    ElseIf Register = TOKEN_Z80_C Then
        Let Opcode_RegisterY = y1
    ElseIf Register = TOKEN_Z80_D Then
        Let Opcode_RegisterY = y2
    ElseIf Register = TOKEN_Z80_E Then
        Let Opcode_RegisterY = y3
    ElseIf Register = TOKEN_Z80_H Then
        Let Opcode_RegisterY = y4
    ElseIf Register = TOKEN_Z80_L Then
        Let Opcode_RegisterY = y5
    ElseIf Register = TOKEN_Z80_MEM_HL Then
        Let Opcode_RegisterY = y6
    ElseIf Register = TOKEN_Z80_MEM_IX Then
        'TODO: opcode prefix
        Let Opcode_RegisterY = y6
    ElseIf Register = TOKEN_Z80_MEM_IY Then
        'TODO: opcode prefix
        Let Opcode_RegisterY = y6
    ElseIf Register = TOKEN_Z80_A Then
        Let Opcode_RegisterY = y7
    Else
        Stop
    End If
End Function

'Opcode_RegisterZ : Sets the Z-bits (0, 1 & 2) of the opcode based on the register _
 ======================================================================================
Private Function Opcode_RegisterZ( _
    ByRef Register As OZ80_TOKEN _
) As Integer
    If Register = TOKEN_Z80_B Then
        Exit Function
    ElseIf Register = TOKEN_Z80_C Then
        Let Opcode_RegisterZ = z1
    ElseIf Register = TOKEN_Z80_D Then
        Let Opcode_RegisterZ = z2
    ElseIf Register = TOKEN_Z80_E Then
        Let Opcode_RegisterZ = z3
    ElseIf Register = TOKEN_Z80_H Then
        Let Opcode_RegisterZ = z4
    ElseIf Register = TOKEN_Z80_L Then
        Let Opcode_RegisterZ = z5
    ElseIf Register = TOKEN_Z80_MEM_HL Then
        Let Opcode_RegisterZ = y6
    ElseIf Register = TOKEN_Z80_MEM_IX Then
        'TODO: opcode prefix
        Let Opcode_RegisterZ = y6
    ElseIf Register = TOKEN_Z80_MEM_IY Then
        'TODO: opcode prefix
        Let Opcode_RegisterZ = y6
    ElseIf Register = TOKEN_Z80_A Then
        Let Opcode_RegisterZ = z7
    Else
        Stop
    End If
End Function

'Opcode_Rot : Set opcode for shift / rotate instructions _
 ======================================================================================
Private Function Opcode_Rot( _
    ByRef Instruction As OZ80_TOKEN _
) As Integer
    If Instruction = TOKEN_Z80_RLC Then
        Exit Function
    ElseIf Instruction = TOKEN_Z80_RRC Then
        Let Opcode_Rot = y1
    ElseIf Instruction = TOKEN_Z80_RL Then
        Let Opcode_Rot = y2
    ElseIf Instruction = TOKEN_Z80_RR Then
        Let Opcode_Rot = y3
    ElseIf Instruction = TOKEN_Z80_SLA Then
        Let Opcode_Rot = y4
    ElseIf Instruction = TOKEN_Z80_SRA Then
        Let Opcode_Rot = y5
    ElseIf Instruction = TOKEN_Z80_SLL Then
        Let Opcode_Rot = y6
    ElseIf Instruction = TOKEN_Z80_SRL Then
        Let Opcode_Rot = y7
    Else
        Stop
    End If
End Function
