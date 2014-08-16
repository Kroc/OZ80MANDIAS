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

'--------------------------------------------------------------------------------------

'http://z80.info/decoding.htm
    
' .---------------.
' |7|6|5|4|3|2|1|0|
' '---------------'
' | x |  y  |  z  |
' '---------------'
'     | p |q|
'     '-----'

Private Enum OZ80_OPCODE
    X1 = 2 ^ 6
    X2 = 2 ^ 7
    X3 = X1 Or X2
    
    Y1 = 2 ^ 3
    Y2 = 2 ^ 4
    Y3 = Y1 Or Y2
    Y4 = 2 ^ 5
    Y5 = Y4 Or Y1
    Y6 = Y4 Or Y2
    Y7 = Y4 Or Y3
    
    Z1 = 2 ^ 0
    Z2 = 2 ^ 1
    Z3 = Y1 Or Y2
    Z4 = 2 ^ 2
    Z5 = Z4 Or Z1
    Z6 = Z4 Or Z2
    Z7 = Z4 Or Z4
    
    P1 = 2 ^ 4
    P2 = 2 ^ 5
    P3 = P1 Or P2

    Q1 = Y1
End Enum

'Our routine for comparing parameters is going to be used hundreds of times, _
 so for clarity, we'll use some class-level copies to share the input parameters _
 with the comparison routine
Private Param1Test As oz80Param
Private Param2Test As oz80Param
Private Param3Test As oz80Param

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
Friend Function AddInstruction( _
    ByRef Instruction As OZ80_TOKEN, _
    ByRef Param1 As oz80Param, _
    ByRef Param2 As oz80Param, _
    ByRef Param3 As oz80Param _
) As OZ80_ERROR
    'OPCODE MAP: _
     Here's a complete list of the Z80 instruction set and the opcodes used for each _
     addressing type.
     
     
    'http://www.smspower.org/Development/InstructionSet
    
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
    
    'Our routine for comparing parameters is going to be used hundreds of times, _
     so for clarity, we'll use some class-level copies to share the input parameters _
     with the comparison routine
    Let Param1Test = Param1
    Let Param2Test = Param2
    Let Param3Test = Param3
    
    'See, BASIC does have a good use!
    On Instruction - ([_TOKEN_INSTRUCTIONS_BEGIN]) GoTo _
        iADC, iADD, iAND, iBIT, iCALL, iCCF, iCP, iCPD, iCPDR, iCPI, iCPIR, iCPL, _
        iDAA, iDEC, iDI, iDJNZ, iEI, iEX, iEXX, iHALT, iIM, iIN, iINC, iIND, iINDR, _
        iINI, iINIR, iJP, iJR, iLD, iLDD, iLDDR, iLDI, iLDIR, iNEG, iNOP, iOR, iOUT, _
        iOUTD, iOTDR, iOUTI, iOTIR, iPOP, iPUSH, iRES, iRET, iRETI, iRETN, iRLA, iRL, _
        iRLCA, iRLC, iRLD, iRRA, iRR, iRRCA, iRRC, iRRD, iRST, iSBC, iSCF, iSET, _
        iSLA, iSRA, iSLL, iSRL, iSUB, iXOR
    Stop
    
    'http://z80.info/decoding.htm
    Dim Opcode As Integer
    
iADC:
    'Add with Carry _
     ..................................................................................
    '`ADC A, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y1 Or OpcodeZ_Register(Param2.Mask)
        
    '`ADC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y1 Or Z6
        
    '`ADC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        'IX/IY is not supported here due to the ED-prefix
        Let Opcode = &HED00 Or X1 Or OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z2
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iADD:
    'Add: _
     ..................................................................................
    '`ADD HL, BC|DE|HL|SP`
    If TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z1
    
    '`ADD A, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REG_A, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeZ_Register(Param2.Mask)
    
    '`ADD A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iAND:
    'Bitwise AND: _
     ..................................................................................
    '`AND A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y4 Or OpcodeZ_Register(Param1.Mask)
    
    '`AND $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y4 Or Z6
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iBIT:
    'Test a bit: _
     ..................................................................................
    '`BIT 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or X1 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iCALL:
    'Call Procedure: _
     ..................................................................................
    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
    If TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Param1.Token) Or Z4
    
    '`CALL $16` - absolute call
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Q1 Or Z5
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iCCF:
    '..................................................................................
    '`CCF`
    If TestParams() Then
        Let Opcode = Y7 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iCP:
    '..................................................................................
    '`CP A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y7 Or OpcodeZ_Register(Param1.Mask)
    
    '`CP $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y7 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iCPD:
iCPDR:
iCPI:
iCPIR:
    
iCPL:
    'Compliment: Flip the bits of the Accumulator _
     ..................................................................................
    '`CPL`
    If TestParams() Then
        Let Opcode = Y5 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iDAA:
    'Decimal Adjust Accumulator: _
     ..................................................................................
    '`DAA`
    If TestParams() Then
        Let Opcode = Y4 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iDEC:
    'Decrement: _
     ..................................................................................
    '`DEC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z3
    
    '`DEC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Param1.Token) Or Z5
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iDI:
    'Disable Interrupts: _
     ..................................................................................
    If TestParams() Then
        Let Opcode = X3 Or Y6 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iDJNZ:
    'Decrement, Jump if Not Zero: _
     ..................................................................................
    If TestParams(MASK_VAL) Then Let Opcode = Y2: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
     
iEI:
    'Enable Interrupts: _
     ..................................................................................
    If TestParams() Then
        Let Opcode = X3 Or Y7 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iEX:
    'Exchange: _
     ..................................................................................
    '`EX AF AF'`
    If TestParams(MASK_REG_AF, MASK_REG_AF) Then
        Let Opcode = Y1
    
    '`EX (SP), HL|IX|IY`
    ElseIf TestParams(MASK_MEM_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or Y4 Or Z3
    
    '`EX DE, HL`
    ElseIf TestParams(MASK_REG_DE, MASK_REG_HL) Then
        Let Opcode = X3 Or Y5 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iEXX:
    'Exchange All: _
     ..................................................................................
    If TestParams() Then
        Let Opcode = X3 Or Q1 Or P1 Or Z1
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iHALT:
    'Halt: Stop the processor and wait for interrupts _
     ..................................................................................
    '`HALT`
    If TestParams() Then
        Let Opcode = X1 Or Y6 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iIM:

iIN:
    'Read from port: _
     ..................................................................................
    '`IN A, ($8)`
    If TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = X3 Or Y3 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iINC:
    'Increment: _
     ..................................................................................
    '`INC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Z3
    
    '`INC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL) Then
        Let Opcode = OpcodeY_Register(Param1.Token) Or Z4
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iIND:
iINDR:
iINI:
iINIR:

iJP:
    'Jump: _
     ..................................................................................
    '`JP (HL|IX|IY)` or `JP HL|IX|IY` - indirect absolute jump
    If TestParams(MASK_REGS_HL_IXY Or MASK_MEM_HLIXY) Then
        Let Opcode = X3 Or P2 Or Q1 Or Z1
    
    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
    ElseIf TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Param1.Token) Or Z2
    
    '`JP $16` - absolute jump
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Z3
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iJR:
    'Jump Relative: _
     ..................................................................................
    If TestParams(MASK_VAL) Then
        Let Opcode = Y3
    ElseIf TestParams(MASK_FLAGS_CZ, MASK_VAL) Then
        Let Opcode = OpcodeY_Condition(Param1.Token) + Y4
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iLD:
    'The Mother-Load: _
     ..................................................................................
    '`LD BC|DE|HL|SP|IX|IY, $16`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY, MASK_VAL) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Z1
    
    '`LD (BC), A`
    ElseIf TestParams(MASK_MEM_BC, MASK_REG_A) Then
        Let Opcode = Z2
        
    '`LD (DE), A`
    ElseIf TestParams(MASK_MEM_DE, MASK_REG_A) Then
        Let Opcode = P1 Or Z2
    
    '`LD A, (BC)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_BC) Then
        Let Opcode = Q1 Or Z2
    
    '`LD A, (DE)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_DE) Then
        Let Opcode = P1 Or Q1 Or Z2
    
    '`LD ($16), HL`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REG_HL) Then
        Let Opcode = P2 Or Z2
    
    '`LD ($16), A`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = P3 Or Z2
    
    '`LD HL, ($16)`
    ElseIf TestParams(MASK_REG_HL, MASK_MEM_VAL) Then
        Let Opcode = P2 Or Q1 Or Z2
    
    '`LD A, ($16)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = P3 Or Q1 Or Z2
    
    '`LD A|B|C|D|E|H|L|(HL|IX+$8|IY+$8), $8`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY, MASK_VAL) Then
        Let Opcode = OpcodeY_Register(Param1.Token) Or Z6
    
    '`LD A|B|C|D|E|H|L|(HL|IX+$8|IY+$8), A|B|C|D|E|H|L`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X1 Or OpcodeY_Register(Param1.Token) Or OpcodeZ_Register(Param2.Token)
        
    '`LD A|B|C|D|E|H|L, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X1 Or OpcodeY_Register(Param1.Token) Or OpcodeZ_Register(Param2.Token)
    
    '`LD SP, HL|IX|IY`
    ElseIf TestParams(MASK_REG_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P3 Or Q1 Or Z1
        
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iLDD:
iLDDR:
iLDI:
iLDIR:
iNEG:

iNOP:
    'No Operation: _
     ..................................................................................
    '`NOP`
    If TestParams() Then
        Let Opcode = 0
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iOR:
    'Bitwise OR: _
     ..................................................................................
    '`OR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y6 Or OpcodeZ_Register(Param1.Mask)
    
    '`OR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y6 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iOUT:
    'Write to port: _
     ..................................................................................
    '`OUT ($8), A`
    If TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = X3 Or Y2 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iOUTD:
iOTDR:
iOUTI:
iOTIR:

iPOP:
    'Fetch from stack: _
     ..................................................................................
    '`POP AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z1
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iPUSH:
    'Push a register on to the stack: _
     ..................................................................................
    '`PUSH AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z5
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRES:
    'Reset a bit: _
     ..................................................................................
    '`RES 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or X2 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRET:
    'Return: _
     ..................................................................................
    '`RET NZ|Z|NC|C|PO|PE|P|M` - conditional return
    If TestParams(MASK_FLAGS) Then
        Let Opcode = X3 Or OpcodeY_Condition(Param1.Token)
    
    '`RET`
    ElseIf TestParams() Then
        Let Opcode = X3 Or Q1 Or Z1
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iRETI:
iRETN:

iRLA:
    '..................................................................................
    '`RLA`
    If TestParams() Then
        Let Opcode = Y2 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRL:
    '..................................................................................
    '`RL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y2 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iRLCA:
    '..................................................................................
    '`RLCA`
    If TestParams() Then
        Let Opcode = Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRLC:
    '..................................................................................
    '`RLC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRLD:

iRRA:
    '..................................................................................
    '`RRA`
    If TestParams() Then
        Let Opcode = Y3 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRR:
    '..................................................................................
    '`RR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y3 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iRRCA:
    '..................................................................................
    '`RRCA`
    If TestParams() Then
        Let Opcode = Y1 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iRRC:
    '..................................................................................
    '`RRC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y1 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iRRD:

iRST:
    '"Reset" - Fire specific interrupt vector: _
     ..................................................................................
    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
    If TestParams(MASK_VAL) Then
        'TODO: Need to validate the number is in range at some point in the chain
        Let Opcode = X3 Or OpcodeY_RST(Param1.Value) Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iSBC:
    '..................................................................................
    '`SBC A, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y3 Or OpcodeZ_Register(Param2.Mask)
    
    '`SBC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y3 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iSCF:
    '..................................................................................
    '`SCF`
    If TestParams() Then
        Let Opcode = Y6 Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iSET:
    'Set a bit: _
     ..................................................................................
    '`SET 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or X3 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iSLA:
    '..................................................................................
    '`SLA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y4 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iSRA:
    '..................................................................................
    '`SRA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y5 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iSLL:
    '..................................................................................
    '`SLL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y6 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iSRL:
    '..................................................................................
    '`SRL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00 Or Y7 Or OpcodeZ_Register(Param1.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish

iSUB:
    'Subtract: _
     ..................................................................................
    '`SUB A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y2 Or OpcodeZ_Register(Param1.Mask)
    
    '`SUB $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y2 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    GoTo Finish
    
iXOR:
    'Bitwise XOR: _
     ..................................................................................
    '`XOR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y5 Or OpcodeZ_Register(Param1.Mask)
    
    '`XOR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y5 Or Z6
    Else
        'ERROR: Wrong parameters for Z80 instruction
        Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS: Exit Function
    End If
    
Finish:
    
End Function

'OpcodeY_Alu : Converts arithmetic instructions to bits 3, 4 & 5 of opcode _
 ======================================================================================
Private Function OpcodeY_Alu( _
    ByRef Instruction As OZ80_TOKEN _
) As Integer
    If Instruction = TOKEN_Z80_ADD Then
        Exit Function
    ElseIf Instruction = TOKEN_Z80_ADC Then
        Let OpcodeY_Alu = Y1
    ElseIf Instruction = TOKEN_Z80_SUB Then
        Let OpcodeY_Alu = Y2
    ElseIf Instruction = TOKEN_Z80_SBC Then
        Let OpcodeY_Alu = Y3
    ElseIf Instruction = TOKEN_Z80_AND Then
        Let OpcodeY_Alu = Y4
    ElseIf Instruction = TOKEN_Z80_XOR Then
        Let OpcodeY_Alu = Y5
    ElseIf Instruction = TOKEN_Z80_OR Then
        Let OpcodeY_Alu = Y6
    ElseIf Instruction = TOKEN_Z80_CP Then
        Let OpcodeY_Alu = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_Bit : Convert a bit number into the opcode Y bits _
 ======================================================================================
Private Function OpcodeY_Bit( _
    ByRef Value As Double _
) As Integer
    'TODO: Cannot allow indefinite number during phase 3
    If Value = OZ80_INDEFINITE Then
        'For now, treat as bit 0
        Exit Function
    ElseIf Value = 0 Then
        Exit Function
    ElseIf Value = 1 Then
        Let OpcodeY_Bit = Y1
    ElseIf Value = 2 Then
        Let OpcodeY_Bit = Y2
    ElseIf Value = 3 Then
        Let OpcodeY_Bit = Y3
    ElseIf Value = 4 Then
        Let OpcodeY_Bit = Y4
    ElseIf Value = 5 Then
        Let OpcodeY_Bit = Y5
    ElseIf Value = 6 Then
        Let OpcodeY_Bit = Y6
    ElseIf Value = 7 Then
        Let OpcodeY_Bit = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_Condition : Convert a condition flag to bits 3, 4 & 5 of opcode _
 ======================================================================================
Private Function OpcodeY_Condition( _
    ByRef Condition As OZ80_TOKEN _
) As Integer
    If Condition = TOKEN_Z80_NZ Then
        Exit Function
    ElseIf Condition = TOKEN_Z80_Z Then
        Let OpcodeY_Condition = Y1
    ElseIf Condition = TOKEN_Z80_NC Then
        Let OpcodeY_Condition = Y2
    ElseIf Condition = TOKEN_Z80_C Then
        Let OpcodeY_Condition = Y3
    ElseIf Condition = TOKEN_Z80_PO Then
        Let OpcodeY_Condition = Y4
    ElseIf Condition = TOKEN_Z80_PE Then
        Let OpcodeY_Condition = Y5
    ElseIf Condition = TOKEN_Z80_P Then
        Let OpcodeY_Condition = Y6
    ElseIf Condition = TOKEN_Z80_M Then
        Let OpcodeY_Condition = Y7
    Else
        Stop
    End If
End Function

'OpcodeP_RegisterPair : Converts 16-bit register tokens to bits 4 & 5 of the opcode _
 ======================================================================================
Private Function OpcodeP_RegisterPair( _
    ByRef Register As OZ80_TOKEN _
) As Integer
    If Register = TOKEN_Z80_BC Then
        Exit Function
    ElseIf Register = TOKEN_Z80_DE Then
        Let OpcodeP_RegisterPair = P1
    ElseIf Register = TOKEN_Z80_HL Then
        Let OpcodeP_RegisterPair = P2
    ElseIf Register = TOKEN_Z80_SP Or Register = TOKEN_Z80_AF Then
        Let OpcodeP_RegisterPair = P3
    Else
        Stop
    End If
End Function

'OpcodeY_Register : Sets the Y-bits (3, 4 & 5) of the opcode based on the register _
 ======================================================================================
Private Function OpcodeY_Register( _
    ByRef Mask As OZ80_MASK _
) As Integer
    If Mask = MASK_REG_B Then
        Exit Function
    ElseIf Mask = MASK_REG_C Then
        Let OpcodeY_Register = Y1
    ElseIf Mask = MASK_REG_D Then
        Let OpcodeY_Register = Y2
    ElseIf Mask = MASK_REG_E Then
        Let OpcodeY_Register = Y3
    ElseIf Mask = MASK_REG_H Then
        Let OpcodeY_Register = Y4
    ElseIf Mask = MASK_REG_L Then
        Let OpcodeY_Register = Y5
    ElseIf Mask = MASK_MEM_HLIXY Then
        Let OpcodeY_Register = Y6
    ElseIf Mask = MASK_REG_A Then
        Let OpcodeY_Register = Y7
    Else
        Stop
    End If
End Function

'OpcodeZ_Register : Sets the Z-bits (0, 1 & 2) of the opcode based on the register _
 ======================================================================================
Private Function OpcodeZ_Register( _
    ByRef Mask As OZ80_MASK _
) As Integer
    If Mask = MASK_REG_B Then
        Exit Function
    ElseIf Mask = MASK_REG_C Then
        Let OpcodeZ_Register = Z1
    ElseIf Mask = MASK_REG_D Then
        Let OpcodeZ_Register = Z2
    ElseIf Mask = MASK_REG_E Then
        Let OpcodeZ_Register = Z3
    ElseIf Mask = MASK_REG_H Then
        Let OpcodeZ_Register = Z4
    ElseIf Mask = MASK_REG_L Then
        Let OpcodeZ_Register = Z5
    ElseIf Mask = MASK_MEM_HLIXY Then
        Let OpcodeZ_Register = Z6
    ElseIf Mask = MASK_REG_A Then
        Let OpcodeZ_Register = Z7
    Else
        Stop
    End If
End Function

'OpcodeY_Rot : Set opcode for shift / rotate instructions _
 ======================================================================================
Private Function OpcodeY_Rot( _
    ByRef Instruction As OZ80_TOKEN _
) As Integer
    If Instruction = TOKEN_Z80_RLC Then
        Exit Function
    ElseIf Instruction = TOKEN_Z80_RRC Then
        Let OpcodeY_Rot = Y1
    ElseIf Instruction = TOKEN_Z80_RL Then
        Let OpcodeY_Rot = Y2
    ElseIf Instruction = TOKEN_Z80_RR Then
        Let OpcodeY_Rot = Y3
    ElseIf Instruction = TOKEN_Z80_SLA Then
        Let OpcodeY_Rot = Y4
    ElseIf Instruction = TOKEN_Z80_SRA Then
        Let OpcodeY_Rot = Y5
    ElseIf Instruction = TOKEN_Z80_SLL Then
        Let OpcodeY_Rot = Y6
    ElseIf Instruction = TOKEN_Z80_SRL Then
        Let OpcodeY_Rot = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_RST : Set Opcode Y bits for an RST instruction address _
 ======================================================================================
Private Function OpcodeY_RST(ByRef Value As Double) As Integer
    'TODO: Cannot allow indefinite number during phase 3
    If Value = OZ80_INDEFINITE Then
        'For now, treat as `RST 0`
        Exit Function
    ElseIf Value = &H0 Then
        Exit Function
    ElseIf Value = &H8 Then
        Let OpcodeY_RST = Y1
    ElseIf Value = &H10 Then
        Let OpcodeY_RST = Y2
    ElseIf Value = &H18 Then
        Let OpcodeY_RST = Y3
    ElseIf Value = &H20 Then
        Let OpcodeY_RST = Y4
    ElseIf Value = &H28 Then
        Let OpcodeY_RST = Y5
    ElseIf Value = &H30 Then
        Let OpcodeY_RST = Y6
    ElseIf Value = &H38 Then
        Let OpcodeY_RST = Y7
    Else
        Stop
    End If
End Function

'TestParams : Check if parameters are of the right type desired _
 ======================================================================================
Private Function TestParams( _
    Optional ByRef Test1 As OZ80_MASK, _
    Optional ByRef Test2 As OZ80_MASK, _
    Optional ByRef Test3 As OZ80_MASK _
) As Boolean
    'Why call the function with no parameters? It's a consistent way to check when _
     there should be no parameters for an instruction!
    
    'If a parameter exists, but there shouldn't be, return False
    If (Test1 = 0) And (Param1Test.Mask <> 0) Then Exit Function
    'If there is no first parameter, pass True
    If (Test1 = 0) And (Param1Test.Mask = 0) Then GoTo Finish
    
    '1st Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Is the 'memory reference' bit the same in the parameter as in the test?
    If (Param1Test.Mask And [_MASK_MEM]) Xor (Test1 And [_MASK_MEM]) Then Exit Function
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L _
     NOTE: The memory reference bit is removed to avoid false positives where the _
           memory reference bit is present, but the rest of the mask does not match
    If ( _
        (Param1Test.Mask And Not [_MASK_MEM]) And (Test1 And Not [_MASK_MEM]) _
    ) = 0 Then Exit Function
    
    'If there shouldn't be a second parameter, return False
    If (Test2 = 0) And (Param2Test.Mask <> 0) Then Exit Function
    'If there is no second parameter, the test has passed
    If (Test2 = 0) And (Param2Test.Mask = 0) Then GoTo Finish
    
    '2nd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Is the 'memory reference' bit the same in the parameter as in the test?
    If (Param2Test.Mask And [_MASK_MEM]) Xor (Test2 And [_MASK_MEM]) Then Exit Function
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L _
     NOTE: The memory reference bit is removed to avoid false positives where the _
           memory reference bit is present, but the rest of the mask does not match
    If ( _
        (Param2Test.Mask And Not [_MASK_MEM]) And (Test2 And Not [_MASK_MEM]) _
    ) = 0 Then Exit Function
    
    'If there shouldn't be a third parameter, return False
    If (Test3 = 0) And (Param3Test.Mask <> 0) Then Exit Function
    'If there is no third parameter, the test has passed
    If (Test3 = 0) And (Param3Test.Mask = 0) Then GoTo Finish
    
    '3rd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Is the 'memory reference' bit the same in the parameter as in the test?
    If (Param3Test.Mask And [_MASK_MEM]) Xor (Test3 And [_MASK_MEM]) Then Exit Function
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L _
     NOTE: The memory reference bit is removed to avoid false positives where the _
           memory reference bit is present, but the rest of the mask does not match
    If ( _
        (Param3Test.Mask And Not [_MASK_MEM]) And (Test3 And Not [_MASK_MEM]) _
    ) = 0 Then Exit Function
    
Finish:
    Let TestParams = True
End Function
