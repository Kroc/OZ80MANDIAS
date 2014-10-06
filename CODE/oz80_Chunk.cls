VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_Chunk"
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
'CLASS :: oz80_Chunk

'Code and/or data is assembled in separate chunks (this class) which are arranged _
 according to Section layout before being written out to the ROM. _
 i.e. each `PROC` / `TABLE` will be its own chunk

'This class helps keep the complexities of the opcode mapping private from the _
 assembler. It'll also make supporting other CPUs easier if we ever get to that

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'Which section the chunk belongs to. This will be the CRC32 of the Section name, _
 as stored in oz80_Assembler. Value will be 0 for no particular section
Public Section As Long

'The absolute position of the chunk in the ROM. _
 This will be assigned after the first assembly phase when all chunk sizes are known _
 and they can be placed in their given sections
Public Address As Long

'If the Chunk is an Interrupt Procedure its fixed address will be specified here, _
 which can be 0 (i.e. `START`), so the class will default this to -1
Public Interrupt As Long

'The assembled code/data goes here
Private Chunk() As Byte

Private My_Size As Long                 'Length of the chunk in bytes
Private My_CyclesMin As Long            'Minimum number of cycles in this chunk
Private My_CyclesMax As Long            'Maximum number of cycles in this chunk

'Each chunk starts at this size (many chunks will be small Z80 routines). _
 If this buffer size is filled, it will be doubled. Once the chunk is "finalised", _
 the buffer will be cut down to the size of the actual contents
Private Const BUFFER_SIZE As Long = 1024
Private BufferSize As Long

'--------------------------------------------------------------------------------------

'Z80 opcodes are made up of various bit patterns that can be broken down into the _
 following layout: (this information from <z80.info/decoding.htm>)

'             .---------------.
' Bit Number: |7|6|5|4|3|2|1|0|
'             '---------------'
' Component:  | X |  Y  |  Z  |
'             '---------------'
'                 | P |Q|
'                 '-----'

'Here we encode these component's bit numbers so we can easily build opcodes. _
 A set of functions at the bottom of this class translate instructions / registers _
 into various component parts

Private Enum OZ80_OPCODE
    X0 = 0
    X1 = 2 ^ 6
    X2 = 2 ^ 7
    X3 = X1 Or X2
    
    Y0 = 0
    Y1 = 2 ^ 3
    Y2 = 2 ^ 4
    Y3 = Y1 Or Y2
    Y4 = 2 ^ 5
    Y5 = Y4 Or Y1
    Y6 = Y4 Or Y2
    Y7 = Y4 Or Y3
    
    P0 = 0
    P1 = 2 ^ 4
    P2 = 2 ^ 5
    P3 = P1 Or P2

    Q0 = 0
    Q1 = Y1
    
    Z0 = 0
    Z1 = 2 ^ 0
    Z2 = 2 ^ 1
    Z3 = Z1 Or Z2
    Z4 = 2 ^ 2
    Z5 = Z4 Or Z1
    Z6 = Z4 Or Z2
    Z7 = Z4 Or Z3
End Enum

'A small lookup table for converting bit-numbers in instructions to opcode-components
Private OpcodeY_Bit(0 To 7) As Long

'Our routine for comparing parameters is going to be used hundreds of times, _
 so for clarity, we'll use some class-level copies to share the input parameters _
 with the comparison routine
Private Param1Test As oz80Param
Private Param2Test As oz80Param
Private Param3Test As oz80Param

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output
Event Message( _
    ByRef LogLevel As OZ80_LOG, ByRef Text As String _
)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise the space where the assembled bytes go. _
     This chunk of bytes will be copied into the ROM once the location has been chosen
    Let BufferSize = BUFFER_SIZE
    ReDim Chunk(0 To BufferSize) As Byte
    
    'Interrupt Procedures can be at $0000, so default this to -1
    Let Me.Interrupt = -1
    
    'Populate the lookup table for converting bit numbers in z80 instructions
    Let OpcodeY_Bit(0) = Y0
    Let OpcodeY_Bit(1) = Y1
    Let OpcodeY_Bit(2) = Y2
    Let OpcodeY_Bit(3) = Y3
    Let OpcodeY_Bit(4) = Y4
    Let OpcodeY_Bit(5) = Y5
    Let OpcodeY_Bit(6) = Y6
    Let OpcodeY_Bit(7) = Y7
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Chunk
    Erase OpcodeY_Bit
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddByte : Add a byte to the assembled output and move along _
 ======================================================================================
Public Sub AddByte(ByRef Value As Long)
    'Add the byte to the ROM and move to the next byte
    Let Chunk(My_Size) = Value And &HFF
    Let My_Size = My_Size + 1
    
    'If the chunk's buffer is full then double the buffer size to compensate. _
     This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, 16 KB
    If My_Size > BufferSize Then
        Let BufferSize = BufferSize * 2
        ReDim Preserve Chunk(0 To BufferSize) As Byte
    End If
End Sub

'AddWord : Add a 2-byte value, in little-endian order _
 ======================================================================================
Public Sub AddWord(ByRef Value As Long)
    'Add the low-byte first
    Call AddByte(Value And &HFF)
    'And the high-byte
    Call AddByte((Value And &HFF00&) \ &H100)
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
    '                               (hl|ix+$8|iy+$8)    rr 8E ??            7
    '                               $8                  CE ??               7
    '       hl                      bc|de|hl|sp         ED 4A+rr            15
                                
    'ADD    a                       a|b|c|d|e|h|l       80+r                4
    '                               (hl|ix+$8|iy+$8)    rr 86 ??            7
    '                               $8                  C6 ??               7
    '       hl                      bc|de|hl|sp         09+rr               11
    '       ix                      bc|de|ix|sp         rr 09+rr            11
    '       iy                      bc|de|iy|sp         rr 09+rr            11
    
    'AND    a|b|c|d|e|h|l                               A0+r
    '       (hl|ix+$8|iy+$8)                            rr A6 ??
    '       $8                                          E6 ??
    
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
    '       $8                                          F6 ??
    
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
    
    'SLL    a|b|c|d|e|h|l                               CB 30+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 36
    
    'SRA    a|b|c|d|e|h|l                               CB 28+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 2E
    
    'SRL    a|b|c|d|e|h|l                               CB 38+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 3E
    
    'SUB    a|b|c|d|e|h|l                               90+r
    '       (hl|ix+$8|iy+$8)                            rr 96 ??
    '       $8                                          D6 ??
    
    'XOR    a|b|c|d|e|h|l                               A8+r
    '       (hl|ix+$8|iy+$8)                            rr AE ??
    '       $8                                          EE ??
    
    '----------------------------------------------------------------------------------
    Dim Opcode As Long
    
    'Since the numerical value that accompanies some instructions could be in _
     param 1 or 2, we need to record which to use at the end. This excludes _
     the extra value that goes with IX/IY which is handled separately, _
     e.g. when using `LD (IX+$8), $8`
    Dim ValueParam As oz80Param
    'Whether the opcode supports 8 or 16-bit operands
    Dim ValueBits As OZ80_MASK
    
    'Use of IX / IY requires a prefix byte ($DD/$FD) before the opcode. _
     We need to record which parameter is the IX/IY reference, along with its value
    Dim ParamIXY As oz80Param
    'NOTE: Parameter three is only used for some undocumented instructions and _
     will not be IX/IY (though IX/IY can be in previous parameters)
    If (Param1.Mask And MASK_MEM_HLIXY) <> 0 _
        Then If (Param1.Token = TOKEN_Z80_IX) Or (Param1.Token = TOKEN_Z80_IY) _
            Then Let ParamIXY = Param1
    If (Param2.Mask And MASK_MEM_HLIXY) <> 0 _
        Then If (Param2.Token = TOKEN_Z80_IX) Or (Param2.Token = TOKEN_Z80_IY) _
            Then Let ParamIXY = Param2
    
    'Our routine for comparing parameters is going to be used hundreds of times, _
     so for clarity, we'll use some class-level references to share the input _
     parameters with the comparison routine
    Let Param1Test = Param1
    Let Param2Test = Param2
    Let Param3Test = Param3
    
    'See, BASIC does have a good use!
    On (Instruction - [_TOKEN_INSTRUCTIONS_BEGIN]) GoTo _
        iADC, iADD, iAND, iBIT, iCALL, iCCF, iCP, iCPD, iCPDR, iCPI, iCPIR, iCPL, _
        iDAA, iDEC, iDI, iDJNZ, iEI, iEX, iEXX, iHALT, iIM, iIN, iINC, iIND, iINDR, _
        iINI, iINIR, iJP, iJR, iLD, iLDD, iLDDR, iLDI, iLDIR, iNEG, iNOP, iOR, iOUT, _
        iOUTD, iOTDR, iOUTI, iOTIR, iPOP, iPUSH, iRES, iRET, iRETI, iRETN, iRLA, iRL, _
        iRLCA, iRLC, iRLD, iRRA, iRR, iRRCA, iRRC, iRRD, iRST, iSBC, iSCF, iSET, _
        iSLA, iSRA, iSLL, iSRL, iSUB, iXOR
    
    'TODO: ERROR: completely invalid token number!!
    Stop
    
iADC:
    'Add with Carry _
     ..................................................................................
    '`ADC A, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y1 Or OpcodeZ_Register(Param2.Mask)
        
    '`ADC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y1 Or Z6
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL8
        
    '`ADC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        'IX/IY is not supported here due to the ED-prefix
        Let Opcode = &HED00& Or X1 Or Q1 Or Z2 Or OpcodeP_RegisterPair(Param2.Token)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iADD:
    'Add: _
     ..................................................................................
    '`ADD HL, BC|DE|HL|SP`
    If TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1
    
    '`ADD IX, BC|DE|IX|SP`
    ElseIf TestParams(MASK_REG_IX, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IX Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IX"), _
         we still need to validate the params to ensure HL & IX are not mixed
        Let Opcode = OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1
        Let ParamIXY = Param1
    
    '`ADD IY, BC|DE|IY|SP`
    ElseIf TestParams(MASK_REG_IY, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IY Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IY"), _
         we still need to validate the params to ensure HL & IY are not mixed
        Let Opcode = OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1
        Let ParamIXY = Param1
    
    '`ADD A, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REG_A, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeZ_Register(Param2.Mask)
    
    '`ADD A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Z6
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
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
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iBIT:
    'Test a bit: _
     ..................................................................................
    '`BIT 0-7, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or X1 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask)
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iCALL:
    'Call Procedure: _
     ..................................................................................
    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
    If TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Param1.Token) Or Z4
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL16
    
    '`CALL $16` - absolute call
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Q1 Or Z5
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL16
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iCCF:
    '..................................................................................
    If TestParams() Then Let Opcode = Y7 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCP:
    '..................................................................................
    '`CP A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y7 Or OpcodeZ_Register(Param1.Mask)
    
    '`CP $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y7 Or Z6
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iCPD:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y5 Or Z1: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPDR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y7 Or Z1: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iCPI:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y4 Or Z1: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPIR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y6 Or Z1: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPL:
    'Compliment: Flip the bits of the Accumulator _
     ..................................................................................
    If TestParams() Then Let Opcode = Y5 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDAA:
    'Decimal Adjust Accumulator: _
     ..................................................................................
    If TestParams() Then Let Opcode = Y4 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iDEC:
    'Decrement: _
     ..................................................................................
    '`DEC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z3
        Let ParamIXY = Param1
    
    '`DEC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Param1.Mask) Or Z5
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iDI:
    'Disable Interrupts: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Y6 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDJNZ:
    'Decrement, Jump if Not Zero: _
     ..................................................................................
    If TestParams(MASK_VAL) Then
        Let Opcode = Y2
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8       'TODO: convert 16-bit address to 8-bit disp.
        GoTo Finish
    End If
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
     
iEI:
    'Enable Interrupts: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Y7 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iEX:
    'Exchange: _
     ..................................................................................
    '`EX AF AF'`
    If TestParams(MASK_REG_AF, MASK_REG_AF) Then
        Let Opcode = Y1
    
    '`EX (SP), HL|IX|IY`
    ElseIf TestParams(MASK_MEM_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or Y4 Or Z3
        Let ParamIXY = Param2
    
    '`EX DE, HL`
    ElseIf TestParams(MASK_REG_DE, MASK_REG_HL) Then
        Let Opcode = X3 Or Y5 Or Z3
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iEXX:
    'Exchange All: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Q1 Or P1 Or Z1: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iHALT:
    'Halt: Stop the processor and wait for interrupts _
     ..................................................................................
    If TestParams() Then Let Opcode = X1 Or Y6 Or Z6: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iIM:
    'Interrupt Mode: _
     ..................................................................................
    '`IM 0|1|2`
    If TestParams(MASK_VAL) Then
        'An indefinite value is assumed to be zero. _
         The assembler will ensure that there are no indefinite values remaining _
         on the second parsing of the source code
        If Param1.Value = OZ80_INDEFINITE Then
            Let Opcode = &HED00& Or X1 Or Z6
        ElseIf Param1.Value = 0 Then
            Let Opcode = &HED00& Or X1 Or Z6
        ElseIf Param1.Value = 1 Then
            Let Opcode = &HED00& Or X1 Or Y2 Or Z6
        ElseIf Param1.Value = 2 Then
            Let Opcode = &HED00& Or X1 Or Y3 Or Z6
        Else
            'ERROR: Wrong parameters for Z80 instruction
            GoTo InvalidParams
        End If
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iIN:
    'Read from port: _
     ..................................................................................
    '`IN A, ($8)`
    If TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = X3 Or Y3 Or Z3
        Let ValueParam = Param2
        Let ValueBits = MASK_MEM_VAL8
    
    '`IN A|B|C|D|E|H|L, (C)`
    ElseIf TestParams(MASK_REGS_ABCDEHL, MASK_MEM_BC) Then
        Let Opcode = &HED00& Or X1 Or OpcodeY_Register(Param1.Mask)
    
    '`IN (C)`
    ElseIf TestParams(MASK_MEM_BC) Then
        Let Opcode = &HED00& Or X1 Or Y6
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iINC:
    'Increment: _
     ..................................................................................
    '`INC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Z3
        Let ParamIXY = Param1
    
    '`INC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Param1.Mask) Or Z4
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iIND:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y5 Or Z2: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINDR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y7 Or Z2: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iINI:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y4 Or Z2: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINIR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y6 Or Z2: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iJP:
    'Jump: _
     ..................................................................................
    '`JP (HL|IX|IY)` or `JP HL|IX|IY` - indirect absolute jump
    If TestParams(MASK_REGS_HL_IXY Or MASK_MEM_HLIXY) Then
        Let Opcode = X3 Or P2 Or Q1 Or Z1
    
    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
    ElseIf TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Param1.Token) Or Z2
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL16
    
    '`JP $16` - absolute jump
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Z3
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL16
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iJR:
    'Jump Relative: _
     ..................................................................................
    '`JR $8` - relative immediate jump
    If TestParams(MASK_VAL) Then
        Let Opcode = Y3
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    
    '`JR NZ|Z|NC|C, $8` - conditional relative immediate jump
    ElseIf TestParams(MASK_FLAGS_CZ, MASK_VAL) Then
        Let Opcode = OpcodeY_Condition(Param1.Token) + Y4
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iLD:
    'The Mother-Load: _
     ..................................................................................
    '`LD BC|DE|HL|SP|IX|IY, $16`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY, MASK_VAL) Then
        Let Opcode = OpcodeP_RegisterPair(Param1.Token) Or Z1
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL16
        Let ParamIXY = Param1
    
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
    
    '`LD ($16), HL|IX|IY`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_HL_IXY) Then
        Let Opcode = P2 Or Z2
        Let ValueParam = Param1
        Let ValueBits = MASK_MEM_VAL16
        Let ParamIXY = Param2
    
    '`LD ($16), A`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = P3 Or Z2
        Let ValueParam = Param1
        Let ValueBits = MASK_MEM_VAL16
    
    '`LD HL|IX|IY, ($16)`
    ElseIf TestParams(MASK_REGS_HL_IXY, MASK_MEM_VAL) Then
        Let Opcode = P2 Or Q1 Or Z2
        Let ParamIXY = Param1
        Let ValueParam = Param2
        Let ValueBits = MASK_MEM_VAL16
    
    '`LD A, ($16)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = P3 Or Q1 Or Z2
        Let ValueParam = Param2
        Let ValueBits = MASK_MEM_VAL16
    
    '`LD A|B|C|D|E|H|L|(HL|IX+$8|IY+$8), $8`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY, MASK_VAL) Then
        Let Opcode = OpcodeY_Register(Param1.Mask) Or Z6
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL8
        Let ParamIXY = Param1
    
    '`LD A|B|C|D|E|H|L|(HL|IX+$8|IY+$8), A|B|C|D|E|H|L`
    ElseIf TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask)
        
    '`LD A|B|C|D|E|H|L, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEHL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask)
    
    '`LD SP, HL|IX|IY`
    ElseIf TestParams(MASK_REG_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P3 Or Q1 Or Z1
        Let ParamIXY = Param2
        
    '`LD ($16), BC|DE|SP`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_BC_DE_SP) Then
        Let Opcode = &HED00& Or X1 Or OpcodeP_RegisterPair(Param2.Token) Or Z3
        Let ValueParam = Param1
        Let ValueBits = MASK_MEM_VAL16
        Let ParamIXY = Param2
    
    '`LD BC|DE|SP, ($16)`
    ElseIf TestParams(MASK_REGS_BC_DE_SP, MASK_MEM_VAL) Then
        Let Opcode = &HED00& Or X1 Or OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z3
        Let ValueParam = Param2
        Let ValueBits = MASK_MEM_VAL16
        
    '`LD I, A`
    ElseIf TestParams(MASK_REG_I, MASK_REG_A) Then
        Let Opcode = &HED00& Or X1 Or Z7
    
    '`LD R, A`
    ElseIf TestParams(MASK_REG_R, MASK_REG_A) Then
        Let Opcode = &HED00& Or X1 Or Y1 Or Z7
    
    '`LD A, I`
    ElseIf TestParams(MASK_REG_A, MASK_REG_I) Then
        Let Opcode = &HED00& Or X1 Or Y2 Or Z7
    
    '`LD A, R`
    ElseIf TestParams(MASK_REG_A, MASK_REG_R) Then
        Let Opcode = &HED00& Or X1 Or Y3 Or Z7
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iLDD:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y5: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iLDDR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iLDI:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y4: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iLDIR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y6: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iNEG:
    'Negate the Accumulator: _
     ..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X1 Or Z4: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iNOP:
    'No Operation: _
     ..................................................................................
    '`NOP` - Opcode is 0
    If TestParams() Then GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOR:
    'Bitwise OR: _
     ..................................................................................
    '`OR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y6 Or OpcodeZ_Register(Param1.Mask)
    
    '`OR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y6 Or Z6
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iOUT:
    'Write to port: _
     ..................................................................................
    '`OUT ($8), A`
    If TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = X3 Or Y2 Or Z3
        Let ValueParam = Param1
        Let ValueBits = MASK_MEM_VAL8
        
    '`OUT (C), A|B|C|D|E|H|L`
    ElseIf TestParams(MASK_MEM_BC, MASK_REGS_ABCDEHL) Then
        Let Opcode = &HED00& Or X1 Or OpcodeY_Register(Param2.Mask) Or Z1
    
    '`OUT (C), 0`
    ElseIf TestParams(MASK_MEM_BC, MASK_VAL) Then
        Let Opcode = &HED00& Or X1 Or Y6 Or Z1
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iOUTD:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y5 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTDR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y7 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iOUTI:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y4 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTIR:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X2 Or Y6 Or Z3: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iPOP:
    'Fetch from stack: _
     ..................................................................................
    '`POP AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z1
        Let ParamIXY = Param1
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iPUSH:
    'Push a register on to the stack: _
     ..................................................................................
    '`PUSH AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z5
        Let ParamIXY = Param1
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iRES:
    'Reset a bit: _
     ..................................................................................
    '`RES 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or X2 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask)
    
    '`RES 0-7, (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RES memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or X2 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param3.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
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
        GoTo InvalidParams
    End If
    GoTo Finish

iRETI:
    'Return from Interrupt: _
     ..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X1 Or Y1 Or Z5: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRETN:
    'Return from NMI: _
     ..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X1 Or Z5: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y2 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRL:
    '..................................................................................
    '`RL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y2 Or OpcodeZ_Register(Param1.Mask)
    
    '`RL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RL memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y2 Or OpcodeZ_Register(Param2.Mask)
        
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iRLCA:
    '..................................................................................
    If TestParams() Then Let Opcode = Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLC:
    '..................................................................................
    '`RLC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or OpcodeZ_Register(Param1.Mask)
    
    '`RLC (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RLC memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iRLD:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X1 Or Y5 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRRA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y3 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRR:
    '..................................................................................
    '`RR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y3 Or OpcodeZ_Register(Param1.Mask)
    
    '`RR (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RR memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y3 Or OpcodeZ_Register(Param2.Mask)
        
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iRRCA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y1 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRRC:
    '..................................................................................
    '`RRC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y1 Or OpcodeZ_Register(Param1.Mask)
    
    '`RRC (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RRC memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y1 Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iRRD:
    '..................................................................................
    If TestParams() Then Let Opcode = &HED00& Or X1 Or Y4 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRST:
    '"Reset" - Fire specific interrupt vector: _
     ..................................................................................
    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
    If TestParams(MASK_VAL) Then
        'TODO: Need to validate the number is in range at some point in the chain
        Let Opcode = X3 Or OpcodeY_RST(Param1.Value) Or Z7
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
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
        Let ValueParam = Param2
        Let ValueBits = MASK_VAL8
    
    '`SBC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = &HED00& Or X1 Or OpcodeP_RegisterPair(Param2.Token) Or Z2
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iSCF:
    '..................................................................................
    If TestParams() Then Let Opcode = Y6 Or Z7: GoTo Finish
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iSET:
    'Set a bit: _
     ..................................................................................
    '`SET 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or X3 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask)
    
    '`SET 0-7, (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SET memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or X3 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param3.Mask)
        
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iSLA:
    '..................................................................................
    '`SLA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y4 Or OpcodeZ_Register(Param1.Mask)
    
    '`SLA (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SLA memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y4 Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish

iSRA:
    '..................................................................................
    '`SRA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y5 Or OpcodeZ_Register(Param1.Mask)
    
    '`SRA (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SRA memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y5 Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iSLL:
    '..................................................................................
    '`SLL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y6 Or OpcodeZ_Register(Param1.Mask)
    
    '`SLL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SLL memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y6 Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    GoTo Finish
    
iSRL:
    '..................................................................................
    '`SRL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = &HCB00& Or Y7 Or OpcodeZ_Register(Param1.Mask)
        
    '`SRL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SRL memory and copy to register
    ElseIf TestParams(MASK_MEM_HLIXY, MASK_REGS_ABCDEHL) Then
        'TODO: Can't use HL here!
        Let Opcode = &HCB00& Or Y7 Or OpcodeZ_Register(Param2.Mask)
    
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
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
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
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
        Let ValueParam = Param1
        Let ValueBits = MASK_VAL8
    Else
        'ERROR: Wrong parameters for Z80 instruction
        GoTo InvalidParams
    End If
    
Finish:
    '----------------------------------------------------------------------------------
    'Assemble the opcode & value into the chunk
    '----------------------------------------------------------------------------------
    'If this is an instruction with a value (excluding IX/IY memory references), _
     then check that the value is within the 8 or 16-bit limit of the instruction
    If ValueParam.Token = TOKEN_NUMBER Then
        'If a 16-bit number is used when the instruction only supports 8-bits...
        If (ValueParam.Mask And ValueBits) = 0 Then
            'ERROR: Can't use a 16-bit number with this instruction
            Let AddInstruction = OZ80_ERROR_OVERFLOW_Z80
            Exit Function
        End If
    
    'If this is an IX/IY memory reference, check that it's 8-bit
    ElseIf (ValueParam.Token = TOKEN_Z80_IX) Or (ValueParam.Token = TOKEN_Z80_IY) Then
        If ValueParam.Value > &HFF Then
            'ERROR: IX/IY memory references must use an 8-bit value
            Let AddInstruction = OZ80_ERROR_OVERFLOW_Z80
            Exit Function
        End If
    End If
    
    '..................................................................................
    
    Dim Prefix As Long
    Dim LogMsg As String
    
    'TODO: All this string concatenation is *SLOW* _
           Need to find a way to do logging fast
    Let LogMsg = "$" & oz80.HexStr16(My_Size) & ": "
    
    'If IX or IY are used, add their prefix byte _
     (we remember if this happens to be able to look for the $DD/$FD+$CB special case)
    If ParamIXY.Token = TOKEN_Z80_IX Then
        Call AddByte(&HDD)
        Let Prefix = &HDD
        Let LogMsg = LogMsg & " DD"
    ElseIf ParamIXY.Token = TOKEN_Z80_IY Then
        Call AddByte(&HFD)
        Let Prefix = &HFD
        Let LogMsg = LogMsg & " FD"
    Else
        Let LogMsg = LogMsg & " ··"
    End If
    
    'Is it a double opcode? (i.e. prefixed with $CB/$ED)
    If Opcode > &HFF Then '............................................................
        'TODO:
        If Prefix And (Opcode And &HFF00&) = &HCB00& Then
            'write the first half of the opcode ($CB)
            Call AddByte(&HCB)
            'write the displacement byte
            Call AddByte(ValueParam.Value And &HFF)
            'write the second half of the opcode
            Call AddByte(Opcode And &HFF)
            Let LogMsg = LogMsg _
                & " CB " & oz80.HexStr8(ParamIXY.Value) _
                & " " & oz80.HexStr8(Opcode) & " ·· ··"
        Else
            'write the first half of the opcode
            Call AddByte((Opcode And &HFF00&) \ &H100)
            'Write the second half of the opcode
            Call AddByte(Opcode And &HFF)
            
            Let LogMsg = LogMsg _
                & " " & oz80.HexStr8((Opcode And &HFF00&) \ &H100) _
                & " ·· " & oz80.HexStr8(Opcode And &HFF)
            
            'Value?
            If ValueParam.Token = TOKEN_NUMBER Then
                Call AddWord(ValueParam.Value)
                Let LogMsg = LogMsg _
                & " " & oz80.HexStr8((ValueParam.Value And &HFF00&) \ &H100) _
                & " " & oz80.HexStr8(ValueParam.Value And &HFF)
            Else
                Let LogMsg = LogMsg & " ·· ··"
            End If
        End If
        
    'Single opcode:
    Else '.............................................................................
        'Add the opcode to the output
        Call AddByte(Opcode)
        Let LogMsg = LogMsg & " " & oz80.HexStr8(Opcode)
        'An IX/IY memory reference (with value)
        If (ParamIXY.Mask = MASK_MEM_HLIXY) And (Prefix <> 0) Then
            Call AddByte(ParamIXY.Value)
            Let LogMsg = LogMsg & " " & oz80.HexStr8(ParamIXY.Value) & " ··"
        Else
            Let LogMsg = LogMsg & " ·· ··"
        End If
        'Is there an associated value?
        If ValueParam.Token = TOKEN_NUMBER Then
            'Is it 8 or 16 bits?
            If (ValueBits = MASK_VAL8) Or (ValueBits = MASK_MEM_VAL8) Then
                'Straight-forward 8-bit number
                Call AddByte(ValueParam.Value)
                Let LogMsg = LogMsg & " " & oz80.HexStr8(ValueParam.Value) _
                            & " ··"
            Else
                'Write the 16-bit number as little-endian
                Call AddWord(ValueParam.Value)
                Let LogMsg = LogMsg _
                    & " " & oz80.HexStr8(ValueParam.Value And &HFF) _
                    & " " & oz80.HexStr8((ValueParam.Value And &HFF00&) \ &H100)
            End If
        Else
            Let LogMsg = LogMsg & " ·· ··"
        End If
    End If
    
    Let LogMsg = LogMsg & " | " & oz80.TokenName(Instruction)
    If Param1.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param1)
    If Param2.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param2)
    If Param3.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param3)
    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    Exit Function
    
InvalidParams:
    '----------------------------------------------------------------------------------
    'ERROR: Wrong parameters for Z80 instruction
    Let AddInstruction = OZ80_ERROR_INVALID_Z80PARAMS
End Function

'OpcodeP_RegisterPair : Converts 16-bit register tokens to bits 4 & 5 of the opcode _
 ======================================================================================
Private Function OpcodeP_RegisterPair( _
    ByRef Register As OZ80_TOKEN _
) As Long
    If Register = TOKEN_Z80_BC Then
        Exit Function
    ElseIf Register = TOKEN_Z80_DE Then Let OpcodeP_RegisterPair = P1
    ElseIf Register = TOKEN_Z80_HL Then Let OpcodeP_RegisterPair = P2
    ElseIf Register = TOKEN_Z80_IX Then Let OpcodeP_RegisterPair = P2
    ElseIf Register = TOKEN_Z80_IY Then Let OpcodeP_RegisterPair = P2
    ElseIf Register = TOKEN_Z80_SP Then Let OpcodeP_RegisterPair = P3
    ElseIf Register = TOKEN_Z80_AF Then Let OpcodeP_RegisterPair = P3
    Else
        Stop
    End If
End Function

'OpcodeY_Condition : Convert a condition flag to bits 3, 4 & 5 of opcode _
 ======================================================================================
Private Function OpcodeY_Condition( _
    ByRef Condition As OZ80_TOKEN _
) As Long
    If Condition = TOKEN_Z80_NZ Then
        Exit Function
    ElseIf Condition = TOKEN_Z80_Z Then Let OpcodeY_Condition = Y1
    ElseIf Condition = TOKEN_Z80_NC Then Let OpcodeY_Condition = Y2
    ElseIf Condition = TOKEN_Z80_C Then Let OpcodeY_Condition = Y3
    ElseIf Condition = TOKEN_Z80_PO Then Let OpcodeY_Condition = Y4
    ElseIf Condition = TOKEN_Z80_PE Then Let OpcodeY_Condition = Y5
    ElseIf Condition = TOKEN_Z80_P Then Let OpcodeY_Condition = Y6
    ElseIf Condition = TOKEN_Z80_M Then Let OpcodeY_Condition = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_Register : Sets the Y-bits (3, 4 & 5) of the opcode based on the register _
 ======================================================================================
Private Function OpcodeY_Register( _
    ByRef Mask As OZ80_MASK _
) As Long
    If Mask = MASK_REG_B Then
        Exit Function
    ElseIf Mask = MASK_REG_C Then Let OpcodeY_Register = Y1
    ElseIf Mask = MASK_REG_D Then Let OpcodeY_Register = Y2
    ElseIf Mask = MASK_REG_E Then Let OpcodeY_Register = Y3
    ElseIf Mask = MASK_REG_H Then Let OpcodeY_Register = Y4
    ElseIf Mask = MASK_REG_L Then Let OpcodeY_Register = Y5
    ElseIf Mask = MASK_MEM_HLIXY Then Let OpcodeY_Register = Y6
    ElseIf Mask = MASK_REG_A Then Let OpcodeY_Register = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_Rot : Set opcode for shift / rotate instructions _
 ======================================================================================
Private Function OpcodeY_Rot( _
    ByRef Instruction As OZ80_TOKEN _
) As Long
    If Instruction = TOKEN_Z80_RLC Then
        Exit Function
    ElseIf Instruction = TOKEN_Z80_RRC Then Let OpcodeY_Rot = Y1
    ElseIf Instruction = TOKEN_Z80_RL Then Let OpcodeY_Rot = Y2
    ElseIf Instruction = TOKEN_Z80_RR Then Let OpcodeY_Rot = Y3
    ElseIf Instruction = TOKEN_Z80_SLA Then Let OpcodeY_Rot = Y4
    ElseIf Instruction = TOKEN_Z80_SRA Then Let OpcodeY_Rot = Y5
    ElseIf Instruction = TOKEN_Z80_SLL Then Let OpcodeY_Rot = Y6
    ElseIf Instruction = TOKEN_Z80_SRL Then Let OpcodeY_Rot = Y7
    Else
        Stop
    End If
End Function

'OpcodeY_RST : Set Opcode Y bits for an RST instruction address _
 ======================================================================================
Private Function OpcodeY_RST( _
    ByRef Value As Long _
) As Long
    If Value = &H0 Then
        Exit Function
    ElseIf Value = &H8 Then Let OpcodeY_RST = Y1
    ElseIf Value = &H10 Then Let OpcodeY_RST = Y2
    ElseIf Value = &H18 Then Let OpcodeY_RST = Y3
    ElseIf Value = &H20 Then Let OpcodeY_RST = Y4
    ElseIf Value = &H28 Then Let OpcodeY_RST = Y5
    ElseIf Value = &H30 Then Let OpcodeY_RST = Y6
    ElseIf Value = &H38 Then Let OpcodeY_RST = Y7
    Else
        Stop
    End If
End Function

'OpcodeZ_Register : Sets the Z-bits (0, 1 & 2) of the opcode based on the register _
 ======================================================================================
Private Function OpcodeZ_Register( _
    ByRef Mask As OZ80_MASK _
) As Long
    If Mask = MASK_REG_B Then
        Exit Function
    ElseIf Mask = MASK_REG_C Then Let OpcodeZ_Register = Z1
    ElseIf Mask = MASK_REG_D Then Let OpcodeZ_Register = Z2
    ElseIf Mask = MASK_REG_E Then Let OpcodeZ_Register = Z3
    ElseIf Mask = MASK_REG_H Then Let OpcodeZ_Register = Z4
    ElseIf Mask = MASK_REG_L Then Let OpcodeZ_Register = Z5
    ElseIf Mask = MASK_MEM_HLIXY Then Let OpcodeZ_Register = Z6
    ElseIf Mask = MASK_REG_A Then Let OpcodeZ_Register = Z7
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
     there *should* be no parameters for an instruction!
    
    'Are we testing a first parameter?
    If (Test1 = 0) Then
        'If a parameter exists, but there shouldn't be, return False
        If (Param1Test.Mask <> 0) Then Exit Function
        'If there is no first parameter, pass True
        If (Param1Test.Mask = 0) Then GoTo Finish
    End If
    
    '1st Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param1Test.Mask And Test1) = 0 Then Exit Function
    
    'Are we testing a second parameter?
    If (Test2 = 0) Then
        'If there shouldn't be a second parameter, return False
        If (Param2Test.Mask <> 0) Then Exit Function
        'If there is no second parameter, the test has passed
        If (Param2Test.Mask = 0) Then GoTo Finish
    End If
    
    '2nd Parameter Test: _
     ----------------------------------------------------------------------------------
        
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param2Test.Mask And Test2) = 0 Then Exit Function
    
    'If there shouldn't be a third parameter, return False
    If (Test3 = 0) And (Param3Test.Mask <> 0) Then Exit Function
    'If there is no third parameter, the test has passed
    If (Test3 = 0) And (Param3Test.Mask = 0) Then GoTo Finish
    
    '3rd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param3Test.Mask And Test3) = 0 Then Exit Function
    
Finish:
    Let TestParams = True
End Function
