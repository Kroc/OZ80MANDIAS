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

'The assembled code/data goes here
Private Block() As Byte
Private BlockSize As Long
Private At As Long

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

Public Enum OZ80_PARAM
    PARAM_A = 2 ^ 0                     'Accumulator
    PARAM_B = 2 ^ 1                     'B register
    PARAM_C = 2 ^ 2                     'C register and Carry-set flag
    PARAM_D = 2 ^ 3                     'D register
    PARAM_E = 2 ^ 4                     'E register
  [_PARAM_F] = 2 ^ 5                    'F register, not directly accessible
    PARAM_H = 2 ^ 6                     'H register
    PARAM_I = 2 ^ 7                     'Interrupt, not to be confused with IX/IY
    PARAM_L = 2 ^ 8                     'L register
    PARAM_M = 2 ^ 9                     'Sign-set flag
    PARAM_NC = 2 ^ 10                   'No-Carry flag
    PARAM_NZ = 2 ^ 11                   'Not-Zero flag
    PARAM_P = 2 ^ 12                    'Sign-unset flag
    PARAM_PE = 2 ^ 13                   'Parity/Overflow set flag
    PARAM_PO = 2 ^ 14                   'Parity/Overflow unset flag
    PARAM_R = 2 ^ 15                    'R register
    PARAM_Z = 2 ^ 16                    'Zero flag
    
    PARAM_AF = PARAM_A Or [_PARAM_F]
    PARAM_BC = PARAM_B Or PARAM_C
    PARAM_DE = PARAM_D Or PARAM_E
    PARAM_HL = PARAM_H Or PARAM_L
    
    PARAM_IX = 2 ^ 17
    PARAM_IY = 2 ^ 18
    PARAM_SP = 2 ^ 19
    
    PARAM_MEM = 2 ^ 20                  'Is a memory reference, e.g. `(...)`
    PARAM_VAL = 2 ^ 21                  'Has an associated value
End Enum

Private TestParam1Type As OZ80_PARAM
Private TestParam2Type As OZ80_PARAM
Private TestInstruction As OZ80_TOKEN

'--------------------------------------------------------------------------------------

'The actual opcodes that go into the binary are complicated and are altered by a _
 number of factors, including which register is involved, the use of IX & IY, _
 and the presence of bits, condition flags and 8 or 16-bit values
 
'Public Enum OZ80_OPCODE
'    'Specific values used with opcodes for various instruction features:
'     'E.g. the use of IX/IY often requires extra bytes to denote these
'    '..................................................................................
'    'Opcodes that accept registers usually add these offsets to the opcode
'    OPCODE_REG_A = 7
'    OPCODE_REG_B = 0
'    OPCODE_REG_C = 1
'    OPCODE_REG_D = 2
'    OPCODE_REG_E = 3
'    OPCODE_REG_H = 4
'    OPCODE_REG_L = 5
'
'    'Register-pairs work in the same manner
'    OPCODE_REG_BC = &H0
'    OPCODE_REG_DE = &H10
'    OPCODE_REG_HL = &H20: OPCODE_REG_IX = OPCODE_REG_HL: OPCODE_REG_IY = OPCODE_REG_IX
'    OPCODE_REG_SP = &H30: OPCODE_REG_AF = OPCODE_REG_SP
'
'    'IX and IY have their own opcodes that precede the main opcode
'    OPCODE_IX = &HDD
'    OPCODE_IY = &HFD
'
'    'The flags add these offsets to the base opcode
'    OPCODE_FLAG_NZ = 0                  '..---...
'    OPCODE_FLAG_Z = &H8                 '..--x...
'    OPCODE_FLAG_NC = &H10               '..-x-...
'    OPCODE_FLAG_C = &H18                '..-xx...
'    OPCODE_FLAG_PO = &H20               '..x--...
'    OPCODE_FLAG_PE = &H28               '..x-x...
'    OPCODE_FLAG_P = &H30                '..xx-...
'    OPCODE_FLAG_M = &H38                '..xxx...
'End Enum

'Each block starts at this size (many blocks will be small Z80 routines). _
 If this buffer size is filled, it will be doubled. Once the block is "finalised", _
 the buffer will be cut down to the size of the actual contents
Private Const BLOCK_BUFFER As Long = 1024

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Block: Let BlockSize = 0
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddByte : Add a byte to the assembled output and move along _
 ======================================================================================
Public Sub AddByte(ByRef Value As Byte)
    'Add the byte to the ROM and move to the next byte
    Let Block(At) = Value
    Let At = At + 1
    
    If At > BlockSize Then
        Let BlockSize = BlockSize * 2
        ReDim Preserve Block(0 To BlockSize) As Byte
    End If
End Sub

'AddInstruction : Assemble a Z80 mneomic _
 ======================================================================================
Public Sub AddInstruction( _
             ByRef Instruction As OZ80_TOKEN, _
    Optional ByRef Param1Type As OZ80_PARAM, _
    Optional ByRef Param1Value As Long, _
    Optional ByRef Param2Type As OZ80_PARAM, _
    Optional ByRef Param2Value As Long _
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
    '   +f      nz = +0
    '           z  = +8
    '           nc = +10
    '           c  = +18
    '           po = +20
    '           pe = +28
    '           p  = +30
    '           m  = +38
    '   +b      bit * 8
    '   +im     0  = +0
    '           1  = +10
    '           2  = +18
        
    'Z80    Param 1                 Param 2             Opcode              Offset _
     ----------------------------------------------------------------------------------
    'ADC    a                       a|b|c|d|e|h|l       88+r                 --
    '                               $8                  CE ??               +46
    '                               (hl|ix+$8|iy+$8)    rr 8E ??            +06
    '       hl                      bc|de|hl|sp         ED 4A+rr
                                
    'ADD    a                       a|b|c|d|e|h|l       80+r                 --
    '                               $8                  C6 ??               +46
    '                               (hl|ix+$8|iy+$8)    rr 86 ??            +06
    '       hl                      bc|de|hl|sp         09+rr
    '       ix                      bc|de|ix|sp         rr 09+rr
    '       iy                      bc|de|iy|sp         rr 09+rr
    
    'AND    a|b|c|d|e|h|l                               A0+r                 --
    '       $8                                          E6 ??               +46
    '       (hl|ix+$8|iy+$8)                            rr A6 ??            +06
    
    'BIT    0-7                     a|b|c|d|e|h|l       CB 40+r+b
    '                               (hl|ix+$8|iy+$8)    rr CB ?? 40+r+b
    
    'CALL   $16                                         CD ?? ??
    '       c|nc|m|p|z|nz|pe|po     $16                 C4+f
    
    'CCF                                                3F
    
    'CP     a|b|c|d|e|h|l                               B8+r                 --
    '       (hl|ix+$8|iy+$8)                            rr BE ??            +06
    '       $8                                          FE ??               +46
    
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
    
    Let TestInstruction = Instruction
    Let TestParam1Type = Param1Type
    Let TestParam2Type = Param2Type
    
    Call TestOpcode(TOKEN_Z80_ADC, 2, 1, 0, PARAM_A)
End Sub

Private Function TestParams( _
             ByRef Param1Type As OZ80_PARAM, ByRef Param1Test As OZ80_PARAM, _
    Optional ByRef Param2Type As OZ80_PARAM, Optional ByRef Param2Test As OZ80_PARAM _
) As Boolean
    If (Param1Type And Param1Test) <> Param1Test Then Exit Function
    If Param2Test <> 0 Then
        If (Param2Type And Param2Test) <> Param2Test Then Exit Function
    End If
    Let TestParams = True
End Function

Private Function TestOpcode( _
             ByRef Instruction As OZ80_TOKEN, _
    Optional ByRef x As Byte, Optional ByRef y As Byte, Optional ByRef z As Byte, _
    Optional ByRef Param1Test As OZ80_PARAM, _
    Optional ByRef Param2Test As OZ80_PARAM _
) As Boolean
    If Instruction <> TestInstruction Then Exit Function
    
    'http://z80.info/decoding.htm
    
    ' .---------------.
    ' |7|6|5|4|3|2|1|0|
    ' '---------------'
    ' | x |  y  |  z  |
    ' '---------------'
    '     | p |q|
    '     '-----'
    
    Const x1 As Byte = 2 ^ 6
    Const x2 As Byte = 2 ^ 7
    Const x3 As Byte = x1 Or x2
    
    Const y1 As Byte = 2 ^ 3
    Const y2 As Byte = 2 ^ 4
    Const y3 As Byte = y1 Or y2
    Const y4 As Byte = 2 ^ 5
    Const y5 As Byte = y4 Or y1
    Const y6 As Byte = y4 Or y2
    Const y7 As Byte = y4 Or y4
    
    Const z1 As Byte = 2 ^ 0
    Const z2 As Byte = 2 ^ 1
    Const z3 As Byte = y1 Or y2
    Const z4 As Byte = 2 ^ 2
    Const z5 As Byte = z4 Or z1
    Const z6 As Byte = z4 Or z2
    Const z7 As Byte = z4 Or z4
    
    Const p1 As Byte = 2 ^ 4
    Const p2 As Byte = 2 ^ 5
    Const p3 As Byte = p1 Or p2
    
    Dim Opcode As Byte
End Function

