VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_Assembler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: oz80_Assembler

'/// API //////////////////////////////////////////////////////////////////////////////

'Undocumented VB6 APIs to write a number of bytes to memory. _
 We'll be using this to manually create some Double values
Private Declare Sub api_PutMem1 Lib "msvbvm60" Alias "PutMem1" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Byte _
)
Private Declare Sub api_PutMem2 Lib "msvbvm60" Alias "PutMem2" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Integer _
)
Private Declare Sub api_PutMem4 Lib "msvbvm60" Alias "PutMem2" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Long _
)

'/// DEFINITIONS //////////////////////////////////////////////////////////////////////

'Error number!
Private My_Error As OZ80_ERROR

'During class initialisation we manually create special 64-bit numbers _
 in these values, which cannot be written out natively in VB6
Private INDEFINITE As Double
Private NUMBER_MAX As Double
Private NUMBER_MIN As Double
'The min/max numbers that can be multiplied/divided by 1'000 or 1'024 without overflow
Private NUMBER_MIN_K As Double
Private NUMBER_MAX_K As Double
Private NUMBER_MIN_KB As Double
Private NUMBER_MAX_KB As Double

'Logging: _
 --------------------------------------------------------------------------------------

'Formatting of output messages
Private Const LOG_VERSION       As String = "OZ80MANDIAS v{$}.{$},{$}"
Private Const LOG_PHASE1        As String = "Phase 1: Structural Pass"
Private Const LOG_PHASE2        As String = "Phase 2: Layout ROM"
Private Const LOG_PHASE3        As String = "Phase 3: Full Assemble"
Private Const LOG_PHASE4        As String = "Phase 4: Post Processing"
Private Const LOG_ASM_COMPLETE  As String = "Assembly Complete"
Private Const LOG_TIME          As String = "Time: {$}s"
Private Const LOG_Z80           As String = "{$} | {$} {$} {$} {$} {$} {$} | {$} " & vbTab & "{$}{$}{$}"
Private Const LOG_Z80_CB        As String = "{$} | {$} CB {$} {$} ·· ·· | {$} " & vbTab & "{$}{$}{$}"
Private Const LOG_Z80_VAL8      As String = "{$} | {$} {$} {$} ·· {$} ·· | {$} " & vbTab & "{$}{$}{$}"
Private Const LOG_Z80_VAL16     As String = "{$} | {$} {$} ·· ·· {$} {$} | {$} " & vbTab & "{$}{$}"
Private Const LOG_DATA          As String = "{$} | {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} | {$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}"
Private Const LOG_SECTION       As String = "SECTION {$}"

'String constants to save these being temporarily created and destroyed multiple times
Private Const STR_SMS           As String = "SMS"
Private Const STR_START         As String = "START"

'To save having to create and destroy short strings all the time, _
 we'll keep this one around for formatting log messages
Private LogMsg As bluString

'Which type of data the Chunk is recording (Z80 assembly or Data). _
 This is used to produce the disassembly log
Private ChunkMode As CHUNK_MODE
Private Enum CHUNK_MODE
    MODE_UNSET
    MODE_Z80                            'When adding Z80 instructions
    MODE_DATA                           'When adding data
End Enum

'When generating the disassembly log, we don't show data as 1-byte per line, _
 but render it in typical bin-hex format of 16 hex-bytes followed by ASCII. _
 We therefore need to remember up to 16 bytes as we go before we can output this
Private ChunkData(0 To 15) As String
Private ChunkDataIndex As Long
'The address of where the line of data begins
Private ChunkDataAddr As String * 4

'Parsing: _
 --------------------------------------------------------------------------------------
'Each source code file is broken down into a machine-readable representation _
 ('oz80_File') which we store here. When assembling one file and we come _
 across an `INCLUDE` statement, another File instance gets added here
Private Files() As oz80_File
Private FileCount As Long
'The current source file being processed, attached such that we can _
 listen to the message and error events from it
Private FileIndex As Long
Private WithEvents File As oz80_File
Attribute File.VB_VarHelpID = -1

'Which target system is selected, e.g. SEGA Master System
Private My_System As OZ80_SYSTEM

'The database of named item values. This includes Constants, RAM Names and Hash values
Private Values As Scripting.Dictionary

'Sections need to be defined first before they can be used
Private SectionSlots As Scripting.Dictionary
'For normalising Label names, we note which Section is in play
Private CurrentSection As String

'The database of Label locations (in ROM) and their properties. _
 These are separate from the other named items since we keep the _
 Label addresses for the second assembly pass
Private Labels As Scripting.Dictionary
'For normalising Property names, we note which Label is in play
Private CurrentLabel As String

'Z80 Parameters: _
 --------------------------------------------------------------------------------------
'In order to compare the hundreds of permutations of parameters for Z80 instructions, _
 we assign each parameter type a single bit. We can thus check very quickly if a given _
 parameter falls within an allowed list of accepted types

Public Enum OZ80_MASK
    
    [_MASK_REGS_BEGIN] = 1
    MASK_REG_B = 2 ^ 0
    MASK_REG_C = 2 ^ 1
    MASK_REG_D = 2 ^ 2
    MASK_REG_E = 2 ^ 3
    MASK_REG_H = 2 ^ 4
    MASK_REG_L = 2 ^ 5
    MASK_MEM_HL = 2 ^ 6
    MASK_REG_A = 2 ^ 7
    
    'The presence of an IX/IY prefix on the opcode changes H/L to IXH/IYH or IXL/IYL
     'respectively, but only on instructions that use single byte opcodes.
     'This is officially undocumented, but obviously fair game for old systems
    MASK_REG_IXH = 2 ^ 8
    MASK_REG_IXL = 2 ^ 9
    MASK_REG_IYH = 2 ^ 10
    MASK_REG_IYL = 2 ^ 11
    
    'The presence of an IX/IY prefix on the opcode changes a memory reference "[HL]"
     'to IX/IY, with an offset value e.g. "[IX+$8]"
    MASK_MEM_IX = 2 ^ 12
    MASK_MEM_IY = 2 ^ 13
    [_MASK_REGS_END] = MASK_MEM_IY
    
    'A couple of undocumented instructions allow for IX/IY memory references,
     'but not the standard "[HL]" reference
    MASK_MEM_IXY = MASK_MEM_IX Or MASK_MEM_IY
    'And this is the common "[HL|IX+$8|IY+$8]" form that is used often throughout
    MASK_MEM_HLIXY = MASK_MEM_HL Or MASK_MEM_IXY
    
    'The main 8-bit registers are a common instruction parameter
    MASK_REGS_ABCDEHL = MASK_REG_A Or MASK_REG_B Or MASK_REG_C Or MASK_REG_D Or MASK_REG_E Or MASK_REG_E Or MASK_REG_H Or MASK_REG_L
    'The Z80 clumps HL/IX & IY memory references together with 8-bit registers when
     'building opcodes, i.e. "A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]"
    MASK_REGS_ABCDEHL_MEM_HLIXY = MASK_REGS_ABCDEHL Or MASK_MEM_HLIXY
    'The use of the IX/IY prefix turns H/L into IXH/IXL/IYH/IYL in many instances
    MASK_REGS_IXHL = MASK_REG_IXH Or MASK_REG_IXL
    MASK_REGS_IYHL = MASK_REG_IYH Or MASK_REG_IYL
    MASK_REGS_IXYHL = MASK_REGS_IXHL Or MASK_REGS_IYHL
    MASK_REGS_ABCDEIXYHL_MEM_HLIXY = MASK_REGS_ABCDEHL_MEM_HLIXY Or MASK_REGS_IXYHL
    
    'Very uncommon 8-bit registers
    MASK_REG_I = 2 ^ 14                 'Interrupt register
    MASK_REG_R = 2 ^ 15                 'Refresh register, pseudo-random
    
    'The 16-bit register pairs
    MASK_REG_AF = 2 ^ 16                'The Accumulator and the processor Flags
    MASK_REG_BC = 2 ^ 17                'Registers B & C
    MASK_REG_DE = 2 ^ 18                'Registers D & E
    MASK_REG_HL = 2 ^ 19                'Registers H & L
    MASK_REG_SP = 2 ^ 20                'Stack Pointer
    
    MASK_REG_IX = 2 ^ 21
    MASK_REG_IY = 2 ^ 22
    
    MASK_REGS_BC_DE_SP = MASK_REG_BC Or MASK_REG_DE Or MASK_REG_SP
    'Some instructions accept BC/DE/HL/SP, but not IX & IY due to existing prefixes
    MASK_REGS_BC_DE_HL_SP = MASK_REGS_BC_DE_SP Or MASK_REG_HL
    
    'HL, IX & IY are synonymous as they use an opcode prefix to determine which
    MASK_REGS_HL_IXY = MASK_REG_HL Or MASK_REG_IX Or MASK_REG_IY
    'PUSH / POP allow AF but not SP
    MASK_REGS_AF_BC_DE_HL_IXY = MASK_REG_AF Or MASK_REG_BC Or MASK_REG_DE Or MASK_REGS_HL_IXY
    'The LD instruction can take most 16-bit registers
    MASK_REGS_BC_DE_HL_SP_IXY = MASK_REGS_BC_DE_HL_SP Or MASK_REG_IX Or MASK_REG_IY
    
    'Any immediate 8/16-bit Value,
     'e.g. `LD A, $FF`
    MASK_VAL = 2 ^ 23
    
    '..................................................................................
    
    'Register C & Flag C cannot be distinguished by the tokeniser (it isn't aware of
     'context) so they are treated as the same thing. Another bit covers NC/Z/NZ so
     'that these are not accidentally taken as Register C elsewhere
    MASK_FLAGS_CZ = MASK_REG_C Or (2 ^ 24)
    MASK_FLAGS_MP = (2 ^ 25)
    
    MASK_FLAGS = MASK_FLAGS_CZ Or MASK_FLAGS_MP
    
    '..................................................................................
    
    'The IN and OUT instructions can use port "C" (which is, in reality, BC)
    MASK_MEM_BC = 2 ^ 26
    MASK_MEM_DE = 2 ^ 27
    MASK_MEM_SP = 2 ^ 28
    
    MASK_MEM_VAL = 2 ^ 29
    
    '..................................................................................
    
    'This is a shorthand to check for any instance of IX/IY so that we can add the
     'relevant opcode prefix with the simplest of tests
    MASK_ANY_IX = MASK_REG_IX Or MASK_REG_IXH Or MASK_REG_IXL Or MASK_MEM_IX
    MASK_ANY_IY = MASK_REG_IY Or MASK_REG_IYH Or MASK_REG_IYL Or MASK_MEM_IY
    MASK_ANY_IXY = MASK_ANY_IX Or MASK_ANY_IY
    
    'This is shorthand to check for any 8/16-bit parameter Value
    MASK_ANY_VAL = MASK_VAL Or MASK_MEM_VAL
End Enum

Private Type oz80Param
    Mask As OZ80_MASK
    Token As OZ80_TOKEN
    Value As Long
End Type

'Z80 Opcodes: _
 --------------------------------------------------------------------------------------
'Z80 opcodes are made up of various bit patterns that can be broken down into the _
 following layout: (this information from <z80.info/decoding.htm>)

'             .---------------.
' Bit Number: |7|6|5|4|3|2|1|0|
'             '---------------'
'  Component: | X |  Y  |  Z  |
'             '---------------'
'                 | P |Q|
'                 '-----'

'Here we encode these component's bit numbers so we can easily build opcodes. _
 A set of look-up tables translate instructions / registers into component parts

Private Enum OZ80_OPCODE
    X0 = 0:         X1 = 2 ^ 6:     X2 = 2 ^ 7:     X3 = X1 Or X2
    
    Y0 = 0:         Y1 = 2 ^ 3:     Y2 = 2 ^ 4:     Y3 = Y1 Or Y2
    Y4 = 2 ^ 5:     Y5 = Y4 Or Y1:  Y6 = Y4 Or Y2:  Y7 = Y4 Or Y3
    
    P0 = 0:         P1 = 2 ^ 4:     P2 = 2 ^ 5:     P3 = P1 Or P2

    Q0 = 0:         Q1 = Y1
    
    Z0 = 0:         Z1 = 2 ^ 0:     Z2 = 2 ^ 1:     Z3 = Z1 Or Z2
    Z4 = 2 ^ 2:     Z5 = Z4 Or Z1:  Z6 = Z4 Or Z2:  Z7 = Z4 Or Z3
End Enum

'A look-up table to convert a token for a register into the bit mask for it
Private TokenRegister([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long

'Look-up tables for converting certain z80 parameters into opcode-components
Private OpcodeP_RegisterPair([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long
Private OpcodeY_Condition([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long
Private OpcodeY_RST(0 To &H38) As Long
Private OpcodeY_Bit(0 To 7) As Long
'Sure this is a waste of memory, but it is fast as we don't have to convert _
 bit-powers to linear numbers for many instructions being assembled
Private OpcodeY_Register([_MASK_REGS_BEGIN] To [_MASK_REGS_END]) As Long
Private OpcodeZ_Register([_MASK_REGS_BEGIN] To [_MASK_REGS_END]) As Long

'Our routine for comparing parameters is going to be used hundreds of times, _
 so for clarity, we'll use some class-level references to share the input _
 parameters with the comparison routine
Private Z80Param1 As oz80Param
Private Z80Param2 As oz80Param
Private Z80Param3 As oz80Param
'For quick emptying of the params
Private Z80ParamEmpty As oz80Param

'Whilst in the syntax `SLOT` uses a list (i.e. `SLOT 0, 1, 2`), we convert that into _
 a bit pattern to make it quick and easy to work with instead of iterating an array
Private Enum OZ80_SLOT
    SLOT0 = 2 ^ 0
    SLOT1 = 2 ^ 1
    SLOT2 = 2 ^ 2
End Enum

'Output: _
 --------------------------------------------------------------------------------------

Private Chunks() As oz80_Chunk
Private ChunkCount As Long

'The current chunk being assembled into. This is a reference into the Chunks array _
 so that we can speak to the current Chunk in scope consistently
Private Chunk As oz80_Chunk

'These will be references to the special Chunks at fixed positions within the ROM _
 (when laying out the ROM, we'll need to position these first)
Private Chunk_Header As oz80_Chunk
Private Chunk_Interrupt00 As oz80_Chunk
Private Chunk_Interrupt08 As oz80_Chunk
Private Chunk_Interrupt10 As oz80_Chunk
Private Chunk_Interrupt18 As oz80_Chunk
Private Chunk_Interrupt20 As oz80_Chunk
Private Chunk_Interrupt28 As oz80_Chunk
Private Chunk_Interrupt30 As oz80_Chunk
Private Chunk_Interrupt38 As oz80_Chunk
Private Chunk_Interrupt66 As oz80_Chunk
Private Chunk_Start As oz80_Chunk

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output. _
 A bluString is provided to make manipulating the message before displaying very easy
Event Message( _
    ByRef LogLevel As OZ80_LOG, _
    ByRef LogText As bluString _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByRef FilePath As String, _
    ByVal Number As OZ80_ERROR, _
    ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    'If too many possible tokens are created, things will screw up
    Debug.Assert OZ80_TOKEN.[_TOKEN_LAST] < 256
    
    'Manually create a 64-bit Double "Infinite" number, which cannot be written _
     implicitly in VB6. This method is based upon <stackoverflow.com/a/896292>. _
     We use this to mark numbers as "not yet defined" whilst still allowing full _
     64-Bit floating-point calculations
    Call api_PutMem1(VarPtr(INDEFINITE) + 6, 240)
    Call api_PutMem1(VarPtr(INDEFINITE) + 7, 127)
    'Manually create the largest possible 64-bit number
    Call api_PutMem2(VarPtr(NUMBER_MAX), &HFFFFFFFF)
    Call api_PutMem2(VarPtr(NUMBER_MAX) + 4, -1)
    Call api_PutMem1(VarPtr(NUMBER_MAX) + 6, 239)
    Call api_PutMem1(VarPtr(NUMBER_MAX) + 7, 127)
    'Manually create the largest negative 64-bit number
    Call api_PutMem2(VarPtr(NUMBER_MIN), &HFFFFFFFF)
    Call api_PutMem2(VarPtr(NUMBER_MIN) + 4, -1)
    Call api_PutMem1(VarPtr(NUMBER_MIN) + 6, 239)
    Call api_PutMem1(VarPtr(NUMBER_MIN) + 7, 255)
    'Define min/max numbers that can be multiplied by 1000/1024 without overflow
    Let NUMBER_MAX_K = NUMBER_MAX / 1000
    Let NUMBER_MIN_K = -NUMBER_MAX_K
    Let NUMBER_MAX_KB = NUMBER_MAX / 1024
    Let NUMBER_MIN_KB = -NUMBER_MAX_KB
    
    '----------------------------------------------------------------------------------
    
    'This is a re-usable temporary String for logging
    Set LogMsg = New bluString
    
    '----------------------------------------------------------------------------------
    
    'Populate the look-up table for converting register tokens into their bit masks
    Let TokenRegister(TOKEN_Z80_A) = MASK_REG_A
    Let TokenRegister(TOKEN_Z80_B) = MASK_REG_B
    Let TokenRegister(TOKEN_Z80_C) = MASK_REG_C
    Let TokenRegister(TOKEN_Z80_D) = MASK_REG_D
    Let TokenRegister(TOKEN_Z80_E) = MASK_REG_E
    Let TokenRegister(TOKEN_Z80_H) = MASK_REG_H
    Let TokenRegister(TOKEN_Z80_L) = MASK_REG_L
    Let TokenRegister(TOKEN_Z80_I) = MASK_REG_I
    Let TokenRegister(TOKEN_Z80_R) = MASK_REG_R
    Let TokenRegister(TOKEN_Z80_AF) = MASK_REG_AF
    Let TokenRegister(TOKEN_Z80_BC) = MASK_REG_BC
    Let TokenRegister(TOKEN_Z80_DE) = MASK_REG_DE
    Let TokenRegister(TOKEN_Z80_HL) = MASK_REG_HL
    Let TokenRegister(TOKEN_Z80_SP) = MASK_REG_SP
    Let TokenRegister(TOKEN_Z80_IX) = MASK_REG_IX
    Let TokenRegister(TOKEN_Z80_IXH) = MASK_REG_IXH
    Let TokenRegister(TOKEN_Z80_IXL) = MASK_REG_IXL
    Let TokenRegister(TOKEN_Z80_IY) = MASK_REG_IY
    Let TokenRegister(TOKEN_Z80_IYH) = MASK_REG_IYH
    Let TokenRegister(TOKEN_Z80_IYL) = MASK_REG_IYL
    Let TokenRegister(TOKEN_Z80_NC) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_Z) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_NZ) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_P) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PE) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PO) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_M) = MASK_FLAGS_MP
    
    'Populate the look-up tables for converting z80 parameters into opcodes
    Let OpcodeP_RegisterPair(TOKEN_Z80_BC) = P0
    Let OpcodeP_RegisterPair(TOKEN_Z80_DE) = P1
    Let OpcodeP_RegisterPair(TOKEN_Z80_HL) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_IX) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_IY) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_SP) = P3
    Let OpcodeP_RegisterPair(TOKEN_Z80_AF) = P3
    
    Let OpcodeY_Bit(0) = Y0
    Let OpcodeY_Bit(1) = Y1
    Let OpcodeY_Bit(2) = Y2
    Let OpcodeY_Bit(3) = Y3
    Let OpcodeY_Bit(4) = Y4
    Let OpcodeY_Bit(5) = Y5
    Let OpcodeY_Bit(6) = Y6
    Let OpcodeY_Bit(7) = Y7
    
    Let OpcodeY_Condition(TOKEN_Z80_NZ) = Y0
    Let OpcodeY_Condition(TOKEN_Z80_Z) = Y1
    Let OpcodeY_Condition(TOKEN_Z80_NC) = Y2
    Let OpcodeY_Condition(TOKEN_Z80_C) = Y3
    Let OpcodeY_Condition(TOKEN_Z80_PO) = Y4
    Let OpcodeY_Condition(TOKEN_Z80_PE) = Y5
    Let OpcodeY_Condition(TOKEN_Z80_P) = Y6
    Let OpcodeY_Condition(TOKEN_Z80_M) = Y7
    
    Let OpcodeY_Register(MASK_REG_B) = Y0
    Let OpcodeY_Register(MASK_REG_C) = Y1
    Let OpcodeY_Register(MASK_REG_D) = Y2
    Let OpcodeY_Register(MASK_REG_E) = Y3
    Let OpcodeY_Register(MASK_REG_H) = Y4
    Let OpcodeY_Register(MASK_REG_IXH) = Y4
    Let OpcodeY_Register(MASK_REG_IYH) = Y4
    Let OpcodeY_Register(MASK_REG_L) = Y5
    Let OpcodeY_Register(MASK_REG_IXL) = Y5
    Let OpcodeY_Register(MASK_REG_IYL) = Y5
    Let OpcodeY_Register(MASK_MEM_HL) = Y6
    Let OpcodeY_Register(MASK_MEM_IX) = Y6
    Let OpcodeY_Register(MASK_MEM_IY) = Y6
    Let OpcodeY_Register(MASK_REG_A) = Y7
    
    Let OpcodeY_RST(&H0) = Y0
    Let OpcodeY_RST(&H8) = Y1
    Let OpcodeY_RST(&H10) = Y2
    Let OpcodeY_RST(&H18) = Y3
    Let OpcodeY_RST(&H20) = Y4
    Let OpcodeY_RST(&H28) = Y5
    Let OpcodeY_RST(&H30) = Y6
    Let OpcodeY_RST(&H38) = Y7
    
    Let OpcodeZ_Register(MASK_REG_B) = Z0
    Let OpcodeZ_Register(MASK_REG_C) = Z1
    Let OpcodeZ_Register(MASK_REG_D) = Z2
    Let OpcodeZ_Register(MASK_REG_E) = Z3
    Let OpcodeZ_Register(MASK_REG_H) = Z4
    Let OpcodeZ_Register(MASK_REG_IXH) = Z4
    Let OpcodeZ_Register(MASK_REG_IYH) = Z4
    Let OpcodeZ_Register(MASK_REG_L) = Z5
    Let OpcodeZ_Register(MASK_REG_IXL) = Z5
    Let OpcodeZ_Register(MASK_REG_IYL) = Z5
    Let OpcodeZ_Register(MASK_MEM_HL) = Z6
    Let OpcodeZ_Register(MASK_MEM_IX) = Z6
    Let OpcodeZ_Register(MASK_MEM_IY) = Z6
    Let OpcodeZ_Register(MASK_REG_A) = Z7
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Call Reset
    
    'Free the look-up tables for opcode building
    Erase OpcodeP_RegisterPair, OpcodeY_Bit, OpcodeY_Condition, _
          OpcodeY_Register, OpcodeY_RST, OpcodeZ_Register
    Erase TokenRegister
    
    Set LogMsg = Nothing
End Sub

Private Sub Reset()
    'Dereference the currently attached Chunk
    Set Chunk = Nothing
    'Dereference the fixed Chunks so that the array can be emptied without _
     keeping these ones alive
    Set Chunk_Header = Nothing
    Set Chunk_Interrupt00 = Nothing
    Set Chunk_Interrupt08 = Nothing
    Set Chunk_Interrupt10 = Nothing
    Set Chunk_Interrupt18 = Nothing
    Set Chunk_Interrupt20 = Nothing
    Set Chunk_Interrupt28 = Nothing
    Set Chunk_Interrupt30 = Nothing
    Set Chunk_Interrupt38 = Nothing
    Set Chunk_Interrupt66 = Nothing
    Set Chunk_Start = Nothing
    'Now erase the list of Chunks
    Erase Chunks
    
    'Reset the value stores
    Set SectionSlots = Nothing: Set SectionSlots = New Scripting.Dictionary
    Set Values = Nothing:       Set Values = New Scripting.Dictionary
    Set Labels = Nothing:       Set Labels = New Scripting.Dictionary
    
    Set File = Nothing
    Erase Files
End Sub

'EVENT <File>_Error : Detailed error message from a File object
'======================================================================================
Private Sub File_Error( _
    ByVal Number As OZ80_ERROR, ByVal Line As Long, ByVal Col As Long _
)   'Pass this up to the controller, this class will shortly end processing. _
     The `RaiseError` function will select the standard error title from the number
    Call RaiseError(Number, Line, Col)
End Sub

'EVENT <File>_Message : Basic progress message from a File object
'======================================================================================
Private Sub File_Message( _
    ByRef LogLevel As OZ80_LOG, ByRef LogText As bluString _
)   'It's up to the controller to decide to display (or `Debug.Print` any messages), _
     we 're not going to unnecessarily spam their console
    RaiseEvent Message(LogLevel, LogText)
End Sub

'Error : Expose the internal error number
'======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property




'Assemble : Compile a Z80 binary from a given starting source code file
'======================================================================================
'FilePath       | Source code file to begin with
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Public Function Assemble( _
    ByVal FilePath As String _
) As OZ80_ERROR
    Dim Begin As Single
    Let Begin = Timer
    
    'Reset the error number
    Let My_Error = OZ80_ERROR_NONE
    
    'Register the events we'll be timing with the profiler
    With oz80.Profiler
        Call .RegisterEvent(EVENT_TOKENISE, "File.Tokenise")
        Call .RegisterEvent(EVENT_TOKENISE_READWORD, "File.Tokenise;ReadWord")
        Call .RegisterEvent(EVENT_TOKENISE_TOKENWORD, "File.Tokenise;TokenWord")
        Call .RegisterEvent(EVENT_FORMATTOKEN, "File;FormatToken")
        Call .RegisterEvent(EVENT_PROCESSZ80, "Assembler.ProcessZ80")
        Call .RegisterEvent(EVENT_PROCESSZ80_LOG, "Assembler.ProcessZ80;Log")
    End With
    
    'Announce ourselves
    Call RaiseMessage( _
        OZ80_LOG_INFO, LOG_VERSION, _
        App.Major, App.Minor, App.Revision _
    )
    
    'Clean up stuff from any previous assemble
    Call Reset
    
    '----------------------------------------------------------------------------------
    'Phase 1: First Pass & Assemble _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE1)
    
    'Start with parsing the given file; the assembler does not work with the original _
     text directly. If there was a problem parsing the file, an error message will _
     already have been raised by an event, we only need to exit
    
    'Now walk the File that's been created, assembling it as we go. _
     If an `INCLUDE` is occurred, it will be tokenised and assembled there and then
    If AssembleFile(FilePath) Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    'Phase 2: Layout the ROM _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE2)
    'Now that the size of each Chunk is known, we can arrange the Chunks in the ROM _
     space and determine each Label location
    If Layout() Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    'Phase 3: Second Parse & Assemble (now with ROM Addresses) _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE3)
    'TODO
    
    '----------------------------------------------------------------------------------
    'Phase 4: Post Processing, e.g. Checksum _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE4)
    'TODO
    
    'Finally, write the ROM to disk
'    Dim FileNumber As Integer: Let FileNumber = FreeFile
'    Open App.Path & "\ROM.sms" For Binary Access Write As #FileNumber
'    Put #FileNumber, , ROM
'    Close #FileNumber
    
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_ASM_COMPLETE)
    Call RaiseMessage(OZ80_LOG_INFO, LOG_TIME, Format$(Timer - Begin, "0.000"))
    
Finish:
    'Clean up
    Call Reset
    
    'Return the error code (also available in the `Error` property)
    Let Assemble = My_Error
End Function



'/// PARSING //////////////////////////////////////////////////////////////////////////
'/// Read in the source code and walk it

'AssembleFile : Tokenise a source code file and then begin assembling it
'======================================================================================
'FilePath       | File path to the source code file (".oz80")
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function AssembleFile( _
    ByRef FilePath As String _
) As OZ80_ERROR
    'Since we have a "scrrun.dll" reference for Dictionary object, this also gives _
     us the `FileSystemObject` to simplify a lot of mucky folder / file checking
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    '[1] Normalise the file path: _
     ----------------------------------------------------------------------------------
    'The file path handed to us might be just a file name and/or contain a relative _
     path (i.e. "..\"). We need to normalise this given path into an absolute path _
     (e.g. "C:\Folder\File.ext")
    Dim TestPath As New bluString
    
    'If the given path is already absolute, it does not need normalising. _
     NOTE: Upper/lower case will be auto-corrected and could throw a false negative
    If LCase$(FSO.GetAbsolutePathName(FilePath)) = LCase$(FilePath) Then
        'We'll skip ahead to checking the file exists
        Let TestPath.Text = FilePath
    Else
        'Path is relative! If no other file has been parsed yet _
         we will need to use the current directory as a base
        If File Is Nothing Then
            Let TestPath.Text = FSO.BuildPath(CurDir(), FilePath)
        Else
            'When another file has been parsed before, _
             we work relatively from that file's directory
            Let TestPath.Text = FSO.GetFile(File.Path).ParentFolder.Path
            'TODO: Error?
            Let TestPath.Text = FSO.BuildPath(TestPath.Text, FilePath)
        End If
    End If
    
    '[2] Check file exists: _
     ----------------------------------------------------------------------------------
    'ERROR: File Not Found
    If Not FSO.FileExists(TestPath.Text) Then
        Call RaiseError(OZ80_ERROR_FILE_NOTFOUND)
        GoTo Finish
    End If
    
    '[3] Tokenise the file: _
     ----------------------------------------------------------------------------------
    'Check to see if the file has already been tokenised
    Dim Index As Long, i As Long
    For i = 1 To FileCount
        If Files(i).Path = TestPath.Text Then Let Index = i: Exit For
    Next
    
    'If the file has already been tokenised, re-use it
    If Index <> 0 Then
        'Reference the existing File object
        Let FileIndex = Index
        Set File = Files(FileIndex)
    Else
        '*** TIMER BEGIN ***
        Call oz80.Profiler.EventBegin(EVENT_TOKENISE)
        
        'Create a File object to hold the machine representation of the text file; _
         the assembler doesn't work with the original text directly
        Set File = New oz80_File
        'Explode the source code file into tokens:
        'If the file failed to parse, an error code / message will already have _
         been raised via an event, we only need to exit
        If File.Tokenise(TestPath.Text) Then GoTo Finish
        'Add the File to the files collection
        Let FileCount = FileCount + 1
        ReDim Preserve Files(1 To FileCount) As oz80_File
        Let FileIndex = FileCount
        Set Files(FileIndex) = File
        
        Call oz80.Profiler.EventEnd(EVENT_TOKENISE)
        '*** TIMER END ***
    End If
    
    'The user will obviously want to begin at the start of the stream _
     rather than the end
    Call File.Rewind
    
    '[4] Assemble the file: _
     ----------------------------------------------------------------------------------
    'Now begin assembling the file. If an `INCLUDE` statement is encountered, _
     it will also be tokenised and assembled
    Call ScopeRoot
    
Finish:
    Set FSO = Nothing
    
    'If there were any error, an error event and message would already have been _
     raised, we only need to fold upwards and the parent will handle terminating the _
     assembling process early
    Let AssembleFile = My_Error
End Function

'This is a structured map of the OZ80 syntax: _
 The first column on the left is the name of a scope. For each scope, the columns to _
 the right represent the heirarchy of allowable items in that scope. Whatever is in _
 the third column is only allowed to follow what is in the second column. Items below _
 one another within a column represent more than one item being allowed within a scope.
'Another scope name within a column represents a nested scope. Items within brackets _
 show optional items

'NOTES: *   You can take `{text}` to mean any UCS-2 unicode point _
        *   `{operator}` can be any one of `+ - * / ^ \ & | << >>`

'Scope    : 1           2           3           4         ' Comment
'---------|-----------------------------------------------'----------------------------
'<value>  | {number}                                      '
'         | !{name}                                       ' "Constant"
'         | :{name}                                       ' "Label"
'         | .{name}                                       ' "Property"
'         | #{name}                                       ' "Object"
'         | $.{name}                                      ' "RAM Name"
'         | ::{name}                                      ' "Section"
'         | HI          <expr>                            ' High-8-bits of <expr>
'         | LO          <expr>                            ' Low-8-bits of <expr>
'         |                                               '
'<expr>   | <value>    [{operator}  <expr>]               '
'         | {           <expr>      }                     '
'         |                                               '
'<text>   | "           {text}      "                     '
'         |                                               '
'<list>   | <expr>     [,]         [<list>]               '
'         | <text>     [,]         [<list>]               '
'         |                                               '
'<params> | <param>    [,]         [<param>]              '
'         |                                               '
'<param>  | {register} [HELP        <text>]               '
'         | $.{name}   [HELP        <text>]               '
'         | {number}   [HELP        <text>]               '
'         |                                               '
'.........|...............................................'............................
'         |                                               '
'<root>   | DEF         <def>                             ' Define ...
'         |                                               '
'         | ENUM        !{name}    [<expr>]     <hash>    ' Enumerate Constants
'         |                                               '
'         | <if>                                          '
'         |                                               '
'         | INCLUDE     <text>                            '
'         |                                               '
'         | HASH        #{name}     <hash>                '
'         |                                               '
'         | PROC        :{name}    [HELP        <text>]   '
'         |                        [PARAMS      <params>] '
'         |                        [RETURN      <params>] '
'         |                        [INTERRUPT   <expr>]   '
'         |                         <chunk>               '
'         |                                               '
'         | TABLE       :{name}    [HELP        <text>]   '
'         |                         <chunk>               '
'         |                                               '
'<def>    | !{name}     [<expr>]                          ' Define Constant
'         | SECTION     ::{name}   [SLOT        <list>]   ' Define Section
'         | SYSTEM      <text>                            ' Define System
'         | RAM         $#{name}    <expr>      <expr>    ' Define RAM Bank
'         |                                               '
'<chunk>  | {           <code>      }                     '
'         |                                               '
'<code>   | .{name}                                       '
'         | BYTE        <list>                            '
'         | WORD        <list>                            '
'         | TRIP        <list>                            '
'         | LONG        <list>                            '
'         | <if>                                          '
'         | <asm>                                         '
'         |                                               '
'<if>     | IF          <expr>      <chunk>               '
'         |                        [ELSE        <if>]     '

'ScopeRoot : When we're not within the scope of anything else
'======================================================================================
Private Sub ScopeRoot()
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    Call RaiseMessage(OZ80_LOG_ACTION, "Assembling ""{$}""", File.Path)
    
Continue:
    'Define Constant / RAM Bank / Section / System:
    '----------------------------------------------------------------------------------
    If File.TokenKind = TOKEN_KEYWORD_DEF Then
        'Pass over the `DEF` keyword
        'ERROR: The file ended before the first parameter
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'What kind of definition is this?
        If File.TokenKind = TOKEN_CONST Then
            'Defining a constant
            Call DefineConst
        
        'Defining a RAM Bank
        ElseIf File.TokenKind = TOKEN_KEYWORD_RAM Then
            'Hand off to the function that will process this
            Call DefineRAMBank
            
        'Defining the target System
        ElseIf File.TokenKind = TOKEN_KEYWORD_SYSTEM Then
            'Hand off to the function that will process this
            Call DefineSystem
        
        'Defining a code Section
        ElseIf File.TokenKind = TOKEN_KEYWORD_SECTION Then
            'Hand off to the function that will process this
            Call DefineSection
            
        Else
            'ERROR: Unexpected definition
            'TODO: More specific error number/description
            Call RaiseError(OZ80_ERROR_EXPECTED_ROOT)
        End If
    
    'Define a Hash:
    '----------------------------------------------------------------------------------
    '`HASH #<hash> [ ... ]`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_HASH Then
        'Hand off to the function that will process this
        Call ProcessHash
    
    'Inlcude another source file:
    '----------------------------------------------------------------------------------
    '`INCLUDE <text>`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_INCLUDE Then
        'Hand off to the function that will process this
        Call ProcessInclude
    
    'Define a Procedure:
    '----------------------------------------------------------------------------------
    '`PROC :<label> { [...] }`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_PROC Then
        'Hand off to the function that will process this
        Call DefineProc
    
    'Create some RAM variables:
    '----------------------------------------------------------------------------------
    '`RAM [<expr>] [ ... ]`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_RAM Then
        'Hand off to the function that will process this
        Call DefineRAMHash
    
    'Define the starting Procedure:
    '----------------------------------------------------------------------------------
    '`START {...}`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_START Then
        'Hand off to the function that will process this
        Call DefineStart
        
    'Defing a data table:
    '----------------------------------------------------------------------------------
    '`TABLE :<name> {...}`
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_TABLE Then
        'Hand off to the function that will process this
        Call DefineTable
    
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
        Call RaiseError(OZ80_ERROR_EXPECTED_ROOT)
    End If
    
    If My_Error Then Exit Sub
    'As we loop, other functions will move the Token pointer forward (such as when _
     evaluating expressions). Check if we've hit the end of the file
    If File.EOF Then Exit Sub
    'As long as there is content left to process in the file, keep going
    GoTo Continue
End Sub

'DefineConst : Define a Constant value
'======================================================================================
'DEF        !<name>     [<expr>]
'======================================================================================
Private Sub DefineConst()
    '`ScopeRoot` has already verified that the current token is a Constant name; _
     check if the Constant has already been defined
    If Values.Exists(File.TokenText) Then
        'ERROR: You can't redefine Constants
        Call RaiseError(OZ80_ERROR_DUPLICATE_CONSTANT)
        Exit Sub
    End If
    
    'Remember the Constant's name before moving past it
    Dim ConstName As String
    Let ConstName = File.TokenText
    
    'The constant value is optional (it will default to 0), so the file can end here _
     (it's up to the parent to check if there's a unclosed scope problem)
    If Not File.NextToken Then
        'Define the Constant as zero
        Call Values.Add(ConstName, 0)
        Exit Sub
    End If
    
    'Check that this is an Expression
    If Not File.TokenIsExpression Then
        'ERROR: Expected an Expression to define the Constant
        Call RaiseError(OZ80_ERROR_EXPECTED_EXPRESSION)
        Exit Sub
    End If
    
    'Calculate the Expression's Value:
    '(If there was an error in the Expression, the error event _
      will already have been raised, we only need to fold upwards)
    Dim Value As Double
    If ParseExpression(Value) Then Exit Sub
    
    'Define the Constant:
    Call Values.Add(ConstName, Value)
End Sub

'ProcessHash : Process a Hash block
'======================================================================================
'HASH       #<name>     [ ... ]
'======================================================================================
Private Sub ProcessHash()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    Call File.NextToken
    
    If File.TokenKind = TOKEN_SQUARE_OPEN Then
        Let File.TokenIndex = File.TokenValue
        Call File.NextToken
    End If
End Sub

'ProcessInclude : Process an Include statement
'======================================================================================
'INCLUDE    <text>
'======================================================================================
Private Sub ProcessInclude()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Fetch the file path _
     (if there's an error doing this, fold upwards)
    Dim FilePath As bluString
    If GetText(FilePath) Then Exit Sub
    
    'Remember the current file being parsed
    Dim ReturnFile As Long, ReturnToken As Long
    Let ReturnFile = FileIndex
    Let ReturnToken = File.TokenIndex
    
    'Tokenise & assemble the included file
    If AssembleFile(FilePath.Text) Then Exit Sub
    
    'Reattach the previous file
    Let FileIndex = ReturnFile
    Set File = Files(FileIndex)
    Let File.TokenIndex = ReturnToken
End Sub

'DefineProc : Process a Procedure Chunk
'======================================================================================
'PROC   :<label>
'       [HELP       <text>]
'       [SECTION    ::<section>]    |   [INTERRUPT  <expr>]
'       [PARAMS     <params>]
'       [RETURN     <params>]
'{ [...] }
'======================================================================================
Private Sub DefineProc()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'ERROR: No Label name for Procedure
    If File.TokenKind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_PROC_NAME _
    ): GoTo ErrOut
    
    'Remember the String of the Label name, _
     we can't create the Label until the parameters are read
    Dim LabelName As String
    Let LabelName = File.TokenText
    
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'The `INTERRUPT` parameter is optional, but can be zero, so use 'undefined'
    Dim Interrupt As Double
    Let Interrupt = INDEFINITE
    
ReadParams:
    'Keep track of which parameters have been included or not
    Dim IsInterruptDefined As Boolean
    Dim IsSectionDefined As Boolean
    Dim IsParamsDefined As Boolean
    Dim IsReturnDefined As Boolean
    Dim IsHelpDefined As Boolean
    
    '----------------------------------------------------------------------------------
    '`SECTION` Parameter : Set which Section the Procedure belongs to _
     ----------------------------------------------------------------------------------
    If File.TokenKind = TOKEN_KEYWORD_SECTION Then
        'ERROR: The Section parameter has already been specified
        If IsSectionDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_SECTION _
        ): GoTo ErrOut
        
        'ERROR: Cannot co-exist with the `INTERRUPT` parameter
        If IsInterruptDefined Then Call RaiseError(OZ80_ERROR_INVALID_PROC_INTERRUPT): GoTo ErrOut
        
        'ERROR: File ended before the Section name?
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'ERROR: Not a Section name following the Section keyword
        If File.TokenKind <> TOKEN_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_SECTION_NAME _
        ): GoTo ErrOut
        
        'ERROR: Section has not been defined
        If Not Labels.Exists(File.TokenText) Then Call RaiseError( _
            OZ80_ERROR_INVALID_SECTION _
        ): GoTo ErrOut
        
        'Record the Section name
        Dim SectionName As String
        Let SectionName = File.TokenText
        Let IsSectionDefined = True
        
        'Pass over the Section name
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        GoTo ReadParams
    
    '----------------------------------------------------------------------------------
    '`HELP` Parameter : Provide a description for the procedure
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_HELP Then
        'ERROR: The `HELP` parameter has already been specified
        If IsParamsDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_HELP _
        ): GoTo ErrOut
         
        'Pass over the `HELP` keyword, the file cannot end here
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Fetch the text description (ignored for now)
        'TODO: The opening brace for the PROC is read as an Expression! _
               We need to change our use of brackets...
        'For now, just read and discard only text tokens
        Do
            'If not a text token, then continue
            If File.TokenKind <> TOKEN_TEXT Then Exit Do
            'Move over the token, the file ending is an error
            If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        Loop
        
        Let IsHelpDefined = True
        GoTo ReadParams
        
    '----------------------------------------------------------------------------------
    '`PARAMS` Parameter : Document which Z80 parameters the Procedure uses _
     ----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_PARAMS Then
        'The `PARAMS` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM Names, _
         and optional `HELP` descriptions for each
         
        'ERROR: The `PARAMS` parameter has already been specified
        If IsParamsDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_PARAMS _
        ): GoTo ErrOut
         
        'Pass over the `PARAMS` keyword, the file cannot end here
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        Do
            'Check for a valid `PARAMS` item
            If File.TokenIsRegister() _
            Or File.TokenKind = TOKEN_RAM _
            Then
                'Skip over the list item, these are not recorded just yet _
                 but will be in the future (for auto-documentation)
                'ERROR: File cannot end in the midst of the List
                If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Check for a `HELP` description for the list item
                If File.TokenKind = TOKEN_KEYWORD_HELP Then
                    'Pass over the `HELP` keyword, the file cannot end here
                    If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    'Fetch the text description (ignored for now)
                    'TODO: The opening brace for the PROC is read as an Expression! _
                           We need to change our use of brackets...
                    'For now, just read and discard only text tokens
                    Do
                        'If not a text token, then continue
                        If File.TokenKind <> TOKEN_TEXT Then Exit Do
                        'Move over the token, the file ending is an error
                        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    Loop
                End If
                
            'The list ends when we hit one of the other `PROC` parameters or the _
             opening brace. Anything else is considered a bad `PARAMS` list item
            ElseIf File.TokenKind = TOKEN_KEYWORD_HELP _
            Or File.TokenKind = TOKEN_KEYWORD_RETURN _
            Or File.TokenKind = TOKEN_KEYWORD_SECTION _
            Or File.TokenKind = TOKEN_BRACES_OPEN _
            Then
                'The list has come to the end, _
                 look for more parameters
                Let IsParamsDefined = True
                GoTo ReadParams
            Else
                'ERROR: Only Registers and RAM Names are allowed at the moment
                Call RaiseError(OZ80_ERROR_EXPECTED_PROC_PARAMS)
                GoTo ErrOut
            End If
        Loop
        
    '----------------------------------------------------------------------------------
    '`RETURN` Parameter : Document what the Procedure sets / returns _
     ----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_RETURN Then
        'The `RETURN` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM Names
         
        'ERROR: The `RETURN` parameter has already been specified
        If IsReturnDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_PARAMS _
        ): GoTo ErrOut
         
        'Pass over the `RETURN` keyword, the file cannot end here
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        Do
            'Check for a valid `RETURN` item
            If File.TokenIsRegister() _
            Or File.TokenKind = TOKEN_RAM _
            Then
                'Skip over the list item, these are not recorded just yet _
                 but will be in the future (for auto-documentation)
                'ERROR: File cannot end in the midst of the List
                If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Check for a `HELP` description for the list item
                If File.TokenKind = TOKEN_KEYWORD_HELP Then
                    'Pass over the `HELP` keyword, the file cannot end here
                    If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    'Fetch the text description (ignored for now)
                    'TODO: The opening brace for the PROC is read as an Expression! _
                           We need to change our use of brackets...
                    'For now, just read and discard only text tokens
                    Do
                        'If not a text token, then continue
                        If File.TokenKind <> TOKEN_TEXT Then Exit Do
                        'Move over the token, the file ending is an error
                        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    Loop
                End If
                
            'The list ends when we hit one of the other `PROC` parameters or the _
             opening brace. Anything else is considered a bad `RETURN` list item
            ElseIf File.TokenKind = TOKEN_KEYWORD_HELP _
            Or File.TokenKind = TOKEN_KEYWORD_PARAMS _
            Or File.TokenKind = TOKEN_KEYWORD_SECTION _
            Or File.TokenKind = TOKEN_BRACES_OPEN _
            Then
                'The list has come to the end, _
                 look for more parameters
                Let IsReturnDefined = True
                GoTo ReadParams
            Else
                'ERROR: Only Registers and RAM Names are allowed at the moment
                Call RaiseError(OZ80_ERROR_EXPECTED_PROC_RETURN)
                GoTo ErrOut
            End If
        Loop
        
    '----------------------------------------------------------------------------------
    '`INTERRUPT` Parameter : Define a fixed-location Interrupt Procedure _
     ----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_INTERRUPT Then
        'The `INTERRUPT` parameter fixes a Procedure to a specific location in the _
         ROM intended for hardware interrupts. This parameter cannot co-exist with _
         the `SECTION` parameter
        
        'ERROR: The `INTERRUPT` parameter has already been specified
        If IsInterruptDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_INTERRUPT _
        ): GoTo ErrOut
        
        'ERROR: Cannot co-exist with the `SECTION` parameter
        If IsSectionDefined Then Call RaiseError(OZ80_ERROR_INVALID_PROC_INTERRUPT): GoTo ErrOut
        
        'ERROR: File cannot end before the parameter Expression!
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Is this an Expression at all?
        If Not File.TokenIsExpression Then Call RaiseError(OZ80_ERROR_EXPECTED): GoTo ErrOut
        
        'Calculate the Expression:
        '(If there was an error in the Expression, the error event _
          will already have been raised, we only need to fold upwards)
        If ParseExpression(Interrupt) Then GoTo ErrOut
        
        'An indefinite Value cannot be used for an Interrupt address, we will only be _
         able to define the indefinite Values (i.e. Label locations in ROM) once _
         the whole code is parsed and we use the Sections to layout the ROM
        If Interrupt = INDEFINITE Then Call RaiseError(OZ80_ERROR_INDEFINITE): GoTo ErrOut
        
        'Validate the Value, only certain Interrupt addresses are allowed
        If Not ((Interrupt = &H0) Or (Interrupt = &H8) Or (Interrupt = &H10) _
            Or (Interrupt = &H18) Or (Interrupt = &H20) Or (Interrupt = &H28) _
            Or (Interrupt = &H30) Or (Interrupt = &H38) Or (Interrupt = &H66)) _
        Then Call RaiseError(OZ80_ERROR_INVALID_INTERRUPT): GoTo ErrOut
        
        Let IsInterruptDefined = True
        GoTo ReadParams
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the Chunk
    If File.TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): GoTo ErrOut
    
    'The chunk cannot be left open
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'When we assemble the Procedure, we need to be able to normalise _
     the Property names to their Section+Label
    Let CurrentSection = SectionName
    Let CurrentLabel = LabelName
    
    'Now define the Label name; if a Section was given, it needs to be prepended
    Dim FullLabelName As String
    Let FullLabelName = CurrentSection & CurrentLabel
    
    'You can't name two Chunks the same thing
    If ChunkExists(FullLabelName) _
        Then Call RaiseError(OZ80_ERROR_DUPLICATE_LABEL): GoTo ErrOut
    
    'Create the Chunk
    If Interrupt = INDEFINITE _
        Then Call BeginChunk(FullLabelName) _
        Else Call BeginChunk(FullLabelName, Interrupt)
    
'    Dim LogMsg As String
'    Let LogMsg = "PROC " & LabelName
'    If Interrupt <> INDEFINITE _
'        Then Let LogMsg = LogMsg & " INTERRUPT $" & oz80.HexStr8(CLng(Interrupt))
'    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    
    'Log a header for the disassembly
    Call Chunk.LogASM.Add("PROC {$}", LabelName)
    Call Chunk.LogASM.Add("=========================================================================")
    Let ChunkMode = MODE_UNSET
    
    'Enter the Procedure Chunk to process its innards
    Call ScopeChunk
    
    'Finalise the disassembly log
    If ChunkMode = MODE_DATA Then
        Call Chunk.LogASM.Add( _
            LOG_DATA, _
            ChunkDataAddr, ChunkData, _
            "·", "·", "·", "·", "·", "·", "·", "·", _
            "·", "·", "·", "·", "·", "·", "·", "·" _
        )
    End If
    Call Chunk.LogASM.Add( _
        "=========================================================================" _
    )
    Call RaiseMessage(OZ80_LOG_DEBUG, Chunk.LogASM.Concatenate)
    
ErrOut:
    'Leave this scope
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'DefineRAMBank : Define a RAM Block
'======================================================================================
'DEF RAM    $#{name}    <expr>      <expr>
'======================================================================================
Private Sub DefineRAMBank()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'TODO: Implementation
    Call File.NextToken
    Call File.NextToken
    Call File.NextToken
End Sub

'DefineRAMHash : Process a RAM block
'======================================================================================
'RAM       [<expr>]     [ ... ]
'======================================================================================
Private Sub DefineRAMHash()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'TODO: Implementation
    Call File.NextToken
    If File.TokenKind = TOKEN_SQUARE_OPEN Then
        Let File.TokenIndex = File.TokenValue
        Call File.NextToken
    End If
End Sub

'DefineSection : Define a Section
'======================================================================================
'DEF SECTION    ::<section>    [SLOT   <list>]
'======================================================================================
Private Sub DefineSection()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'ERROR: Not a section name following the section keyword
    If File.TokenKind <> TOKEN_SECTION Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_SECTION_NAME _
    ): Exit Sub
    
    'Keep ahold of the Section identifier
    Dim SectionName As String
    Let SectionName = File.TokenText
    
    'Has this Section already been defined?
    If Labels.Exists(File.TokenText) Then Call RaiseError( _
        OZ80_ERROR_DUPLICATE_SECTION _
    ): Exit Sub
    
    Call Labels.Add(SectionName, INDEFINITE)
    Call SectionSlots.Add(SectionName, 0)
    
    '----------------------------------------------------------------------------------
    
    'If the file ends here it's valid, as the params are optional
    If Not File.NextToken() Then GoTo Finish
    
    'The only allowed Slot lists are: 0 / 1 / 2 / 0, 1 / 0, 1, 2 / 1, 2. _
     These are stored as a bit-pattern allowing packing into one value, _
     which when zero infers the default "let the assembler choose for me"
    Dim Slot As OZ80_SLOT
    
    'The only parameter is the Slot List
    If File.TokenKind = TOKEN_KEYWORD_SLOT Then
        'ERROR: File ending before we have the value
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'ERROR: There needs to be at least one item in the List
        If Not File.TokenIsExpression Then Call RaiseError( _
            OZ80_ERROR_EXPECTED _
        ): Exit Sub
        
        'Get the list of Slots assigned to the Section
        Dim List As Scripting.Dictionary
        Set List = GetListAsDictionary()
        
        'Can't be more than three -- "0, 1, 2"
        If List.Count > 3 Then Call RaiseError(OZ80_ERROR_INVALID_SLOT): Exit Sub
        
        'An indefinite value cannot be used for the Slot
        If List.Item(0) = INDEFINITE Then GoTo ErrIndefinite
        
        'Check for one item: 0, 1 or 2
        If List.Count = 1 Then
            If List.Item(0) = 0 Then Let Slot = SLOT0: GoTo Finish
            If List.Item(0) = 1 Then Let Slot = SLOT1: GoTo Finish
            If List.Item(0) = 2 Then Let Slot = SLOT2: GoTo Finish
            'An invalid value
            GoTo ErrInvalid
        End If
        
        'Check for two items: 0, 1 or 1, 2
        If List.Count = 2 Then
            'An indefinite value cannot be used for the Slot
            If List.Item(1) = INDEFINITE Then GoTo ErrIndefinite
            If List.Item(0) = 0 And List.Item(1) = 1 Then Let Slot = SLOT0 Or SLOT1: GoTo Finish
            If List.Item(0) = 1 And List.Item(1) = 2 Then Let Slot = SLOT1 Or SLOT2: GoTo Finish
            'An invalid value
            GoTo ErrInvalid
        End If
        
        'Check for 0, 1, 2
        If List.Item(0) <> 0 Then GoTo ErrInvalid
        If List.Item(1) <> 1 Then GoTo ErrInvalid
        If List.Item(2) <> 2 Then GoTo ErrInvalid
        Let Slot = SLOT0 Or SLOT1 Or SLOT2
        
    End If

Finish:
    'Record the result:
    Let SectionSlots(SectionName) = Slot
    Call RaiseMessage( _
        OZ80_LOG_STATUS, LOG_SECTION, _
        SectionName _
    )
'    RaiseEvent Message(OZ80_LOG_STATUS, _
'        "SECTION " & File.Strings(SectionKey) & _
'        " SLOT " & IIf( _
'            Slot = 0, "?", _
'                IIf(Slot And OZ80_SLOT.SLOT0, "0 ", "") & _
'                IIf(Slot And OZ80_SLOT.SLOT1, "1 ", "") & _
'                IIf(Slot And OZ80_SLOT.SLOT2, "2 ", "") _
'            ) _
'    )
    Exit Sub

    '----------------------------------------------------------------------------------
ErrInvalid:
    'Only a specific Slot pattern can be used
    Call RaiseError(OZ80_ERROR_INVALID_SLOT)
    Exit Sub

ErrIndefinite:
    'An indefinite Value cannot be used for a Slot number, we will only be _
     able to define the indefinite Values (i.e. Label locations in ROM) once _
     the whole code is parsed and we use the Sections to layout the ROM
    Call RaiseError(OZ80_ERROR_INDEFINITE)
End Sub

'DefineStart : Define the System's boot Procedure
'======================================================================================
'START      { ... }
'======================================================================================
Private Sub DefineStart()
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'The only allowed token is the opening of the Chunk
    If File.TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): Exit Sub
    
    'The chunk cannot be left open
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'You can't define the START procedure twice
    If ChunkExists(STR_START) Then _
        Call RaiseError(OZ80_ERROR_DUPLICATE_START): Exit Sub
    
    'Create the Chunk
    'TODO: Check that the System has been defined
    Call BeginChunk(STR_START, 0)
    
    Let CurrentSection = vbNullString
    Let CurrentLabel = STR_START
    
    'Enter the Procedure Chunk to process its innards
    Call ScopeChunk
    
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'DefineSystem : Define the target system (e.g. SEGA Master System)
'======================================================================================
'SYSTEM     <text>
'======================================================================================
Private Sub DefineSystem()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Check that the parameter is a list (or at least the start of one)
    If Not File.TokenIsList Then Call RaiseError(OZ80_ERROR_EXPECTED_SYSTEM_NAME): Exit Sub
    
    'Fetch the System name string from the proceeding list _
     (if there was any error during that, fold upwards)
    Dim SystemName As bluString
    If GetText(SystemName) Then Exit Sub
    
    'Check if it's any of the allowed System names
    If SystemName.Equals(STR_SMS) Then
        'Configure the assembler for the SEGA Master System
        Let My_System = SYSTEM_SMS
    Else
        'Not a valid system name
        Call RaiseError(OZ80_ERROR_EXPECTED_SYSTEM_NAME)
    End If
End Sub

'DefineTable : Process a data table Chunk
'======================================================================================
'TABLE  :<label>
'       [HELP       <text>]
'       [SECTION    ::<section>]
'{ [...] }
'======================================================================================
Private Sub DefineTable()
    'ERROR: The file ended before the first parameter
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'ERROR: No Label name for Table
    If File.TokenKind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_TABLE_NAME _
    ): GoTo ErrOut
    
    'Remember the String of the Label name, _
     we can't create the Label until the parameters are read
    Dim LabelName As String
    Let LabelName = File.TokenText
    
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
ReadParams:
    'Keep track of which params have been included or not
    Dim IsSectionDefined As Boolean
    
    '----------------------------------------------------------------------------------
    '`SECTION` Parameter : Set which Section the Table belongs to _
     ----------------------------------------------------------------------------------
    If File.TokenKind = TOKEN_KEYWORD_SECTION Then
        'ERROR: The Section parameter has already been specified
        If IsSectionDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_TABLE_SECTION _
        ): GoTo ErrOut
        
        'ERROR: File ended before the Section name?
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'ERROR: Not a Section name following the Section keyword
        If File.TokenKind <> TOKEN_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_SECTION_NAME _
        ): GoTo ErrOut
        
        'ERROR: Section has not been defined
        If Not Labels.Exists(File.TokenText) Then Call RaiseError( _
            OZ80_ERROR_INVALID_SECTION _
        ): GoTo ErrOut
        
        'Record the Section name
        Dim SectionName As String
        Let SectionName = File.TokenText
        Let IsSectionDefined = True
        
        'Pass over the Section name
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        GoTo ReadParams
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the Chunk
    If File.TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): GoTo ErrOut
    
    'The chunk cannot be left open
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'When we assemble the Table, we need to be able to normalise _
     the Property names to their Section+Label
    Let CurrentSection = SectionName
    Let CurrentLabel = LabelName
    
    'Now define the Label name; if a Section was given, it needs to be prepended
    Dim FullLabelName As String
    Let FullLabelName = CurrentSection & CurrentLabel
    
    'You can't name two Chunks the same thing
    If ChunkExists(FullLabelName) Then _
        Call RaiseError(OZ80_ERROR_DUPLICATE_LABEL): GoTo ErrOut
    
    'Create the Chunk
    Call BeginChunk(FullLabelName)
    
'    Dim LogMsg As String
'    Let LogMsg = "PROC " & LabelName
'    If Interrupt <> INDEFINITE _
'        Then Let LogMsg = LogMsg & " INTERRUPT $" & oz80.HexStr8(CLng(Interrupt))
'    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    
    'Log a header for the disassembly
    Call Chunk.LogASM.Add("TABLE {$}", LabelName)
    Call Chunk.LogASM.Add("=========================================================================")
    Let ChunkMode = MODE_UNSET
    
    'Enter the Table Chunk to process its innards
    Call ScopeChunk
    
    'Finalise the disassembly log
    If ChunkMode = MODE_DATA Then
        Call Chunk.LogASM.Add( _
            LOG_DATA, _
            ChunkDataAddr, ChunkData, _
            "·", "·", "·", "·", "·", "·", "·", "·", _
            "·", "·", "·", "·", "·", "·", "·", "·" _
        )
    End If
    Call Chunk.LogASM.Add( _
        "=========================================================================" _
    )
    Call RaiseMessage(OZ80_LOG_DEBUG, Chunk.LogASM.Concatenate)
    
ErrOut:
    'Leave this scope
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'ScopeChunk: Within a code/data Chunk, i.e. `PROC`, `TABLE`
'======================================================================================
Private Sub ScopeChunk()
    Dim Value As Double
    'Within a code/data chunk we can have local labels, data statements and Z80 code
    
Continue:
    'Z80 code?
    If File.TokenIsInstruction Then
        '*** TIMER BEGIN ***
        Call oz80.Profiler.EventBegin(EVENT_PROCESSZ80)
        Call ProcessZ80
        Call oz80.Profiler.EventEnd(EVENT_PROCESSZ80)
        '*** TIMER END ***
        
        If My_Error Then Exit Sub
        GoTo Continue
    
    '----------------------------------------------------------------------------------
    'Byte Data
    '----------------------------------------------------------------------------------
'    ElseIf File.TokenKind = TOKEN_KEYWORD_BYTE Then
'        'Move past the `BYTE` keyword
'        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
'
'        'There MUST be at least one list item after the keyword
'        'TODO: Add error here
'        If Not File.TokenIsList Then Stop
'
'        'Start walking the list
'        Do
'            If File.TokenIsExpression Then
'                If ParseExpression(Value) Then Stop
'
'                If ChunkMode <> MODE_DATA Then
'                    Dim i As Long
'                    For i = 0 To 15
'                        Let ChunkData(i) = oz80.HexStr8(-1)
'                    Next
'                    Let ChunkDataIndex = Chunk.Size Mod 16
'                    Let ChunkDataAddr = oz80.HexStr16(Chunk.Size)
'
'                    If ChunkMode <> MODE_UNSET Then Call Chunk.LogASM.Add( _
'                        "-----+-------------------+-----------------------------+-----------------" _
'                    )
'                    Call Chunk.LogASM.Add( _
'                        "ADDR | 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F | 0123456789ABCDEF" _
'                    )
'                    Call Chunk.LogASM.Add( _
'                        "-----|-------------------------------------------------|-----------------" _
'                    )
'                    Let ChunkMode = MODE_DATA
'                End If
'
'                If Value = INDEFINITE Then
'                    Call Chunk.AddByte(0)
'                    Let ChunkData(ChunkDataIndex) = oz80.HexStr8(-1)
'                Else
'                    Call Chunk.AddByte(CLng(Value))
'                    Let ChunkData(ChunkDataIndex) = oz80.HexStr8(CLng(Value))
'                End If
'
'                Let ChunkDataIndex = ChunkDataIndex + 1
'                If ChunkDataIndex = 16 Then
'                    Call Chunk.LogASM.Add( _
'                        LOG_DATA, _
'                        ChunkDataAddr, ChunkData, _
'                        "·", "·", "·", "·", "·", "·", "·", "·", _
'                        "·", "·", "·", "·", "·", "·", "·", "·" _
'                    )
'                    Erase ChunkData
'                    Let ChunkDataIndex = 0
'                End If
'            ElseIf Not File.TokenIsList Then
'                Exit Do
'            End If
'
'            Call File.NextToken
'        Loop
    End If
    
    'The chunk might be empty `{}`, in which case we don't want to walk it
    If File.TokenKind = TOKEN_BRACES_CLOSE Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    
    'Move to the next token. If the file ends unexpectedly, the chunk has not been _
     closed properly (closing brace is expected)
    If Not File.NextToken() Then
        Call RaiseError(OZ80_ERROR_FILE_END)
    
    'When the closing brace is encountered, we'll drop out of the scope
    ElseIf File.TokenKind <> TOKEN_BRACES_CLOSE Then
        GoTo Continue
    End If
    
'    'Finalise the disassembly log
'    If ChunkMode = MODE_DATA Then
'        Call RaiseMessage( _
'            OZ80_LOG_DEBUG, LOG_DATA, _
'            ChunkDataAddr, ChunkData, _
'            "·", "·", "·", "·", "·", "·", "·", "·", _
'            "·", "·", "·", "·", "·", "·", "·", "·" _
'        )
'    End If
'    Call RaiseMessage(OZ80_LOG_DEBUG, _
'        "=========================================================================" _
'    )
'    Let ChunkMode = MODE_UNSET
    
Finish:
    'Move onto the next token ready for the root scope to process. _
     The file could end just after the chunk, so no error is necessary
    Call File.NextToken
End Sub

'GetZ80Parameter : Fetch the type / value of a Z80 instruction parameter
'======================================================================================
'ReturnParam    | An `oz80_Param` structure to populate
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function GetZ80Parameter( _
    ByRef ReturnParam As oz80Param _
) As OZ80_ERROR
    Dim Value As Double
    
    'If the parameter doesn't contain a Value, default to -1. This is so that when _
     building the opcode we can immediately tell when to include the offset byte
    Let ReturnParam.Value = -1
    
    'A memory reference, such as `[ix+$8]` _
     ----------------------------------------------------------------------------------
    If File.TokenKind = TOKEN_SQUARE_OPEN Then
        'ERROR: File ends just as the memory expression is opened
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'The first item in the memory expression can be a register _
         ..............................................................................
        If File.TokenIsRegister Then
            'Copy the register type to the parameter being built
            Let ReturnParam.Token = File.TokenKind
            
            'The `IN` & `OUT` instructions accept only the C register, _
             though in reality this is register pair BC
            If File.TokenKind = TOKEN_Z80_C Then
                Let ReturnParam.Mask = MASK_MEM_BC
            
            '`LD` accepts registers BC & DE
            ElseIf File.TokenKind = TOKEN_Z80_BC Then
                Let ReturnParam.Mask = MASK_MEM_BC
            ElseIf File.TokenKind = TOKEN_Z80_DE Then
                Let ReturnParam.Mask = MASK_MEM_DE
            
            'HL / IX & IY
            ElseIf File.TokenKind = TOKEN_Z80_HL Then
                Let ReturnParam.Mask = MASK_MEM_HL
            'IX/IY can have an additional expression, e.g. `ld a, [ix+$8]`
            ElseIf File.TokenKind = TOKEN_Z80_IX Then
                Let ReturnParam.Mask = MASK_MEM_IX
                GoTo AddExpr
            ElseIf File.TokenKind = TOKEN_Z80_IY Then
                Let ReturnParam.Mask = MASK_MEM_IY
                GoTo AddExpr
                
            ElseIf File.TokenKind = TOKEN_Z80_SP Then
                Let ReturnParam.Mask = MASK_MEM_SP
            Else
                'ERROR: register not allowed here
                Stop
            End If
            
            'ERROR: File ends before the memory expression is closed
            If Not File.NextToken() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
            
            GoTo EndScope
            
AddExpr:    'Look for the `+` following the register name
            If Not File.NextToken() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
            
            'If present, skip over it and check for an expression
            If File.TokenKind = TOKEN_OPERATOR_ADD Then
                'ERROR: Something must follow the "+"!
                If Not File.NextToken() _
                    Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Try and calculate the Expression
                '(If there was an error in the Expression, the error event _
                  will already have been raised, we only need to fold upwards)
                If ParseExpression(Value) Then GoTo ErrOut
                
                'TODO: Handle indefinite value
                'TODO: Error if number is over 16-bit
                'TODO: Issue warning for rounding
                
                If Value = INDEFINITE _
                    Then Let ReturnParam.Value = 0 _
                    Else Let ReturnParam.Value = Value
                
                If File.TokenKind = TOKEN_SQUARE_CLOSE Then
                    Call File.NextToken
                    GoTo ErrOut
                End If
                'ERROR: Unexpected token in memory expression
                Stop
            End If
        
        '..............................................................................
        Else
            'Is it an expression at all?
            'TODO: ERROR: Not a valid expression
            If Not File.TokenIsExpression Then Stop
        
            'Besides registers, memory addresses are allowed
            Let ReturnParam.Token = TOKEN_NUMBER
            
            'Try and calculate the expression
            '(If there was an error in the Expression, the error event _
              will already have been raised, we only need to fold upwards)
            If ParseExpression(Value) Then GoTo ErrOut
            
            'TODO: Handle indefinite value
            'TODO: Error if number is over 16-bit
            'TODO: Issue warning for rounding
            
            If Value = INDEFINITE _
                Then Let ReturnParam.Value = 0 _
                Else Let ReturnParam.Value = Value
            
            'Any number will class as 16-bit since it's the maximum allowed
            Let ReturnParam.Mask = ReturnParam.Mask Or MASK_MEM_VAL
        End If
    
EndScope:
        '..............................................................................
        If File.TokenKind = TOKEN_SQUARE_CLOSE Then
            Call File.NextToken
            GoTo ErrOut
        End If
        'ERROR: Unexpected token in memory expression
        Stop
    
    'A Register _
     ----------------------------------------------------------------------------------
    ElseIf File.TokenIsRegister Or File.TokenIsFlag Then
        'Copy the register type to the parameter being built
        Let ReturnParam.Token = File.TokenKind
        
        'Map the register token to the mask bit; token numbers are sequential _
         whilst parameter masks are bit combinations allowing rapid comparisons _
         of parameters with allowed options
        Let ReturnParam.Mask = TokenRegister(File.TokenKind)
        
        'Now move beyond the register token
        Call File.NextToken
    
    'Anything else classes as a calculable expression _
     ----------------------------------------------------------------------------------
    ElseIf File.TokenIsExpression Then
        'We manually build a number token, since we don't want to use the first _
         token of the expression -- it might be a label for instance and we can _
         only send registers or numbers to be assembled into opcodes
        Let ReturnParam.Token = TOKEN_NUMBER
        Let ReturnParam.Mask = MASK_VAL

GetValue:
        'Try and calculate the expression
        '(If there was an error in the Expression, the error event _
          will already have been raised, we only need to fold upwards)
        If ParseExpression(Value) Then GoTo ErrOut
        
        'TODO: Handle indefinite value
        'TODO: Error if number is over 16-bit
        'TODO: Issue warning for rounding
        
        If Value = INDEFINITE _
            Then Let ReturnParam.Value = 0 _
            Else Let ReturnParam.Value = Value
    
    '----------------------------------------------------------------------------------
    Else
        'ERROR: Not a Z80 parameter
        'TODO: parameter specific error number
        Call RaiseError(OZ80_ERROR_EXPRESSION_Z80)
    End If

ErrOut:
    'Return the error number
    Let GetZ80Parameter = My_Error
End Function

'ProcessZ80 : Process a Z80 instruction (and parameters)
'======================================================================================
'{asm} [<param>] [<param>] [<param>]
'======================================================================================
Private Sub ProcessZ80()
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN
    Let Instruction = File.TokenKind
    'Skip over the Z80 instruction token now that we know it. _
     If the file ends it's an error because the chunk has not been closed
    If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Up to three parameters are supported to allow for the undocumented instructions, _
     e.g. `SET 0 [IX+$8] A` which sets bit 0 of the address and copies the result to A
    LSet Z80Param1 = Z80ParamEmpty
    LSet Z80Param2 = Z80ParamEmpty
    LSet Z80Param3 = Z80ParamEmpty
    
    '----------------------------------------------------------------------------------
    
    'A built in macro allows `PUSH` and `POP` Z80 instructions to handle multiple _
     parameters; each parameter is broken into it's own `PUSH`/`POP` instruction
    If (Instruction = TOKEN_Z80_PUSH) _
    Or (Instruction = TOKEN_Z80_POP) _
    Then
        'Is there any parameters to begin with?
        If Not File.TokenIsParameter Then
            Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
            Exit Sub
        End If
        
        Do
            'Fetch a parameter
            '(If there was an error, the error number will already have been set, _
              we only need to fold upwards to end the assembling
            If GetZ80Parameter(Z80Param1) Then Exit Sub
            
            'If it's not a register, then it's invalid
            If (Z80Param1.Mask And MASK_REGS_AF_BC_DE_HL_IXY) = 0 Then
                Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
                Exit Sub
            End If
            
            'Assemble this as a single `PUSH`/`POP` instruction
            Call AssembleZ80(Instruction)
            
        'Look for the next register in the list
        Loop While File.TokenIsRegister = True
    
    '----------------------------------------------------------------------------------
    ElseIf File.TokenIsParameter Then
        'Parameter Expressions might return an unknown Value _
         (such as when a Label, which has not yet been addressed, is used)
        
        'Get the first parameter
        '(If there was an error, the error number will already have been set, _
          we only need to fold upwards to end the assembling
        If GetZ80Parameter(Z80Param1) Then Exit Sub
        
        'Check if a second parameter exists
        If File.TokenIsParameter Then
            'Fetch the second parameter
            If GetZ80Parameter(Z80Param2) Then Exit Sub
            
            'Check if a third parameter exists
            If File.TokenIsParameter Then
                'Fetch the third parameter (very uncommon)
                'NOTE: Parameter three is only used for some undocumented instructions _
                 and will not be IX/IY (though IX/IY can be in previous parameters)
                If GetZ80Parameter(Z80Param3) Then Exit Sub
            End If
        End If
        
        Call AssembleZ80(Instruction)
    End If
End Sub

'AssembleZ80 : Assemble bytes from a Z80 instruction and parameters
'======================================================================================
'Instruction    | Token number of the Z80 instruction
'[Prefix]       | Pre-calculated opcode prefix ($DD/$FD) for IX/IY params
'[Offset]       | Pre-calculated offset-byte for opcode, e.g. "[IX+$FF]"
'[Value]        | 8/16-bit parameter Value
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function AssembleZ80( _
    ByRef Instruction As OZ80_TOKEN _
) As OZ80_ERROR
    'Initialse variables to hold each piece of the assembly
    Dim Opcode As Long
    Dim Prefix As Long, Offset As Long, Value As Long
    Let Prefix = -1: Let Offset = -1: Let Value = -1
    
    'Use of IX / IY requires a prefix byte ($DD/$FD) before the opcode. _
     We need to record which parameter is the IX/IY reference, along with its value
    Dim ParamIXY As oz80Param
    
    'Is Param 1 an IX/IY parameter?
    If (Z80Param1.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Z80Param1
    'Is Param 1 an 8/16-bit parameter Value?
    ElseIf (Z80Param1.Mask And MASK_ANY_VAL) <> 0 Then
        Let Value = Z80Param1.Value
    'Is Param 2 an IX/IY parameter?
    ElseIf (Z80Param2.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Z80Param2
    'Is Param 2 an 8/16-bit parameter Value?
    ElseIf (Z80Param1.Mask And MASK_ANY_VAL) <> 0 Then
        Let Value = Z80Param2.Value
    End If
        
    'If there is any IX/Y parameter, setup the opcode prefix
    If ParamIXY.Mask <> 0 Then
        'Is it IX or IY?
        If (ParamIXY.Mask And MASK_ANY_IX) <> 0 _
            Then Let Prefix = &HDD& _
            Else Let Prefix = &HFD&
        'Does it include an offset byte? i.e. "[IX+$8]"
        If (ParamIXY.Mask And MASK_MEM_IXY) <> 0 Then Let Offset = ParamIXY.Value
    End If
    
    Dim Param1Str As String, Param2Str As String, Param3Str As String
    If Z80Param1.Mask <> 0 Then
        Let LogMsg.Text = ParamToString(Z80Param1)
        Let Param1Str = LogMsg.Left(10).Text
    End If
    If Z80Param2.Mask <> 0 Then
        Let LogMsg.Text = ParamToString(Z80Param2)
        Let Param2Str = LogMsg.Left(10).Text
    End If
    If Z80Param3.Mask <> 0 Then
        Let LogMsg.Text = ParamToString(Z80Param3)
        Let Param3Str = LogMsg.Left(10).Text
    End If
    
    If ChunkMode <> MODE_Z80 Then
        If ChunkMode = MODE_DATA Then
            Call Chunk.LogASM.Add( _
                LOG_DATA, _
                ChunkDataAddr, ChunkData, _
                "·", "·", "·", "·", "·", "·", "·", "·", _
                "·", "·", "·", "·", "·", "·", "·", "·" _
            )
        End If
        
        If ChunkMode <> MODE_UNSET Then Call Chunk.LogASM.Add( _
            "-----+-------------------+-----------------------------+-----------------" _
        )
        Call Chunk.LogASM.Add( _
            "ADDR | PX OP OF CB LO HI | MNEMONIC   PARAM1    PARAM2    PARAM3" _
        )
        Call Chunk.LogASM.Add( _
            "-----|-------------------|-----------------------------------------------" _
        )
        Let ChunkMode = MODE_Z80
    End If
    
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
    
    '   ixyh    ixh|iyh
    '   ixyl    ixl|iyl
    '   ixyhl   ixh|iyh|ixl|iyl
        
    'Z80    Param 1                 Param 2                 Opcode         Cycles _
     ----------------------------------------------------------------------------------
    'ADC    a                       a|b|c|d|e|h|ixyh|l|ixyl rr 88+r             4
    '                               [hl|ix+$8|iy+$8]        rr 8E ??            7
    '                               $8                      CE ??               7
    '       hl                      bc|de|hl|sp             ED 4A+rr            15
                                
    'ADD    a                       a|b|c|d|e|h|ixyh|l|ixyl rr 80+r             4
    '                               [hl|ix+$8|iy+$8]        rr 86 ??            7
    '                               $8                         C6 ??            7
    '       hl                      bc|de|hl|sp                09+rr            11
    '       ix                      bc|de|ix|sp             rr 09+rr            11
    '       iy                      bc|de|iy|sp             rr 09+rr            11
    
    'AND    a|b|c|d|e|h|ixyh|l|ixyl                         rr A0+r
    '       [hl|ix+$8|iy+$8]                                rr A6 ??
    '       $8                                                 E6 ??
    
    'BIT    0-7                     a|b|c|d|e|h|l              CB 40+r+b
    '                               [hl|ix+$8|iy+$8]        rr CB ?? 46+b
    '       0-7     [ix+$8|iy+$8]   a|b|c|d|e|h|l           rr CB ?? 40+r+b
    
    'CALL   $16                                                CD ?? ??
    '       c|nc|m|p|z|nz|pe|po     $16                        C4+f
    
    'CCF                                                       3F
    
    'CP     a|b|c|d|e|h|ixyh|l|ixyl                         rr B8+r
    '       [hl|ix+$8|iy+$8]                                rr BE ??
    '       $8                                                 FE ??
    
    'CPD                                                    ED A9
    'CPDR                                                   ED B9
    'CPI                                                    ED A1
    'CPIR                                                   ED B1
    'CPL                                                       2F
    'DAA                                                       27
    
    'DEC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 05+r*8
    '       [hl|ix+$8|iy+$8]                                rr 35 ??
    '       bc|de|hl|sp                                        0B+rr
    '       ix|iy                                           rr 2B
    
    'DI                                                        F3
    
    'DJNZ   $8                                                 10 ??
    
    'EI                                                        FB
    
    'EX     [sp]                    hl|ix|iy                rr E3
    '       af                      af                         08
    '       de                      hl                         EB
    
    'EXX                                                       D9
    'HALT                                                      76
    
    'IM     0-2                                                46+im
    
    'IN     a                       [$8]                       DB ??
    '       [c]                                             ED 70
    '       a|b|c|d|e|h|l           [c]                     ED 40+r
    
    'INC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 04+r*8
    '       bc|de|hl|sp|ix|iy                               rr 03+rr
    '       [hl|ix+$8|iy+$8]                                rr 34 ??
    
    'IND                                                    ED AA
    'INDR                                                   ED BA
    'INI                                                    ED A2
    'INIR                                                   ED B2
    
    'JP     $16                                                C3 ?? ??
    '       [hl|ix|iy]                                      rr E9
    '       c|nc|m|p|z|nz|pe|po     $16                        C2+f ?? ??
    
    'JR     $8                                                 18 ??
    '       c|nc|z|nz               $8                         20+f
    
    'LD     i                       a                       ED 47
    '       r                       a                       ED 4F
    '       a                       i                       ED 57
    '                               r                       ED 5F
    '                               a|b|c|d|e|h|ixyh|l|ixyl rr 78+r
    '                               [hl|ix+$8|iy+$8]        rr 7E ??
    '                               $8                         3E ??
    '                               [bc|de]                    0A+rr
    '                               [$16]                      3A ?? ??
    '       b                       a|b|c|d|e|h|ixyh|l|ixyl rr 40+r
    '                               [hl|ix+$8|iy+$8]        rr 46 ??
    '                               $8                         06 ??
    '       c                       a|b|c|d|e|h|ixyh|l|ixyl rr 48+r
    '                               [hl|ix+$8|iy+$8]        rr 4E ??
    '                               $8                         0E ??
    '       d                       a|b|c|d|e|h|ixyh|l|ixyl rr 50+r
    '                               [hl|ix+$8|iy+$8]        rr 56 ??
    '                               $8                         16 ??
    '       e                       a|b|c|d|e|h|ixyh|l|ixyl    58+r
    '                               [hl|ix+$8|iy+$8]        rr 5E ??
    '                               $8                         1E ??
    '       h                       a|b|c|d|e|h|l              60+r
    '                               [hl|ix+$8|iy+$8]        rr 66 ??
    '                               $8                         26 ??
    '       ixh                     a|b|c|d|e|ixh|ixl       rr 60+r
    '                               $8                      rr 26 ??
    '       iyh                     a|b|c|d|e|iyh|iyl       rr 60+r
    '                               $8                      rr 26 ??
    '       l                       a|b|c|d|e|h|l              68+r
    '                               [hl|ix+$8|iy+$8]        rr 6E ??
    '                               $8                         2E ??
    '       ixl                     a|b|c|d|e|ixh|ixl       rr 68+r
    '                               $8                      rr 2E ??
    '       iyl                     a|b|c|d|e|iyh|iyl       rr 68+r
    '                               $8                      rr 2E ??
    '       bc                      [$16]                   ED 4B ?? ??
    '                               $16                        01+rr ?? ??
    '       de                      [$16]                   ED 5B ?? ??
    '                               $16                        01+rr ?? ??
    '       hl                      [$16]                      2A ?? ??
    '                               $16                        01+rr ?? ??
    '       sp                      [$16]                   ED 7B ?? ??
    '                               hl|ix|iy                rr F9
    '                               $16                        01+rr ?? ??
    '       ix|iy                   [$16]                   rr 2A ?? ??
    '                               $16                     rr 21 ?? ??
    '       [hl|ix+$8|iy+$8]        a|b|c|d|e|h|l           rr 70+r
    '                               $8                      rr 36 ??
    '       [bc]                    a                          02+rr
    '       [de]                    a                          02+rr
    '       [$16]                   a                          32 ?? ??
    '                               bc                      ED 43 ?? ??
    '                               de                      ED 53 ?? ??
    '                               hl|ix|iy                rr 22 ?? ??
    '                               sp                      ED 73 ?? ??
    
    'LDD                                                    ED A8
    'LDDR                                                   ED B8
    'LDI                                                    ED A0
    'LDIR                                                   ED B0
    'NEG                                                    ED 44
    'NOP                                                       00
    
    'OR     a|b|c|d|e|h|ixyh|l|ixyl                         rr B0+r
    '       [hl|ix+$8|iy+$8]                                rr B6 ??
    '       $8                                                 F6 ??
    
    'OUT    [$8]                    a                          D3 ??
    '       [C]                     0                       ED 71
    '                               a|b|c|d|e|h|l           ED 41+r*8
    
    'OUTD                                                   ED AB
    'OTDR                                                   ED BB
    'OUTI                                                   ED A3
    'OTIR                                                   ED B3
    
    'POP    af|bc|de|hl|ix|iy                               rr C1+rr
    
    'PUSH   af|bc|de|hl|ix|iy                               rr C5+rr
    
    'RES    0-7                     a|b|c|d|e|h|l              CB 80+r+b
    '                               [hl|ix+$8|iy+$8]        rr CB ?? 86+b
    '       0-7     [ix+$8|iy+$8]   a|b|c|d|e|h|l           rr CB ?? 80+r+b
    
    'RET                                                       C9
    '       c|nc|m|p|z|nz|pe|po                                C0+f
    
    'RETI                                                   ED 4D
    'RETN                                                   ED 45
    'RLA                                                       17
    
    'RL     a|b|c|d|e|h|l                                      CB 10+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 16
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 10+r
    
    'RLCA                                                      07
    
    'RLC    a|b|c|d|e|h|l                                      CB 00+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 06
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 00+r
    
    'RLD                                                    ED 6F
    'RRA                                                       1F
    
    'RR     a|b|c|d|e|h|l                                      CB 18+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 1E
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 18+r
    
    'RRCA                                                      0F
    
    'RRC    a|b|c|d|e|h|l                                      CB 08+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 0E
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 08+r
    
    'RRD                                                    ED 67
    
    'RST    0|08|10|18|20|28|30|38                             C7+num
    
    'SBC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 98+r
    '       [hl|ix+$8|iy+S8]                                rr 9E ??
    '       a                       $8                         DE ??
    '       hl                      bc|de|hl|sp             ED+rr
    
    'SCF                                                       37
    
    'SET    0-7                     a|b|c|d|e|h|l              CB C0+b+r
    '                               [hl|ix+$8|iy+$8]        rr CB ?? C6+b+r
    '       0-7     [ix+$8|iy+$8]   a|b|c|d|e|h|l           rr CB ?? C0+r+b
    
    'SLA    a|b|c|d|e|h|l                                      CB 20+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 26
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 20+r
    
    'SLL    a|b|c|d|e|h|l                                      CB 30+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 36
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 30+r
    
    'SRA    a|b|c|d|e|h|l                                      CB 28+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 2E
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB 28+r
    
    'SRL    a|b|c|d|e|h|l                                      CB 38+r
    '       [hl|ix+$8|iy+$8]                                rr CB ?? 3E
    '       [ix+$8|iy+$8]           a|b|c|d|e|h|l           rr CB ?? 38+r
    
    'SUB    a|b|c|d|e|h|ixyh|l|ixyl                         rr 90+r
    '       [hl|ix+$8|iy+$8]                                rr 96 ??
    '       $8                                                 D6 ??
    
    'XOR    a|b|c|d|e|h|ixyh|l|ixyl                         rr A8+r
    '       [hl|ix+$8|iy+$8]                                rr AE ??
    '       $8                                                 EE ??
    
    '----------------------------------------------------------------------------------
    
    'The disassembly begins with the current offset address _
     (each instruction can vary in the number of bytes used)
    Dim LogAdr As String
    Let LogAdr = oz80.HexStr16(Chunk.Size)
    
    'See, BASIC does have a good use!
    On (Instruction - [_TOKEN_Z80_BEGIN] + 1) GoTo _
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
    '`ADC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y1 Or OpcodeZ_Register(Z80Param2.Mask)
    '`ADC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y1 Or Z6
    '`ADC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Q1 Or Z2 Or OpcodeP_RegisterPair(Z80Param2.Token)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iADD:
    'Add: _
     ..................................................................................
    '`ADD HL, BC|DE|HL|SP`
    If TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD IX, BC|DE|IX|SP`
    ElseIf TestParams(MASK_REG_IX, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IX Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IX"), _
         we still need to validate the params to ensure HL & IX are not mixed
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD IY, BC|DE|IY|SP`
    ElseIf TestParams(MASK_REG_IY, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IY Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IY"), _
         we still need to validate the params to ensure HL & IY are not mixed
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD A, A|B|C|D|E|H|IXH|IYH|L|IYH|IYL|[HL|IX+$8|IY+$8]`
    ElseIf TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeZ_Register(Z80Param2.Mask)
    '`ADD A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iAND:
    'Bitwise AND: _
     ..................................................................................
    '`AND A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y4 Or OpcodeZ_Register(Z80Param1.Mask)
    '`AND $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y4 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iBIT:
    'Test a bit: _
     ..................................................................................
    '`BIT 0-7, A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X1 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`BIT 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - Undocumented set bit & copy to register
    ElseIf TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X1 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iCALL:
    'Call Procedure: _
     ..................................................................................
    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
    If TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token) Or Z4: GoTo Output16
    '`CALL $16` - absolute call
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Q1 Or Z5: GoTo Output16
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iCCF:
    '..................................................................................
    If TestParams() Then Let Opcode = Y7 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCP:
    '..................................................................................
    '`CP A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y7 Or OpcodeZ_Register(Z80Param1.Mask)
    '`CP $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y7 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iCPD:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y5 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPDR:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y7 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iCPI:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y4 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPIR:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y6 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPL:
    'Compliment: Flip the bits of the Accumulator _
     ..................................................................................
    If TestParams() Then Let Opcode = Y5 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDAA:
    'Decimal Adjust Accumulator: _
     ..................................................................................
    If TestParams() Then Let Opcode = Y4 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iDEC:
    'Decrement: _
     ..................................................................................
    '`DEC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Q1 Or Z3
    '`DEC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z5
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iDI:
    'Disable Interrupts: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Y6 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDJNZ:
    'Decrement, Jump if Not Zero: _
     ..................................................................................
    If TestParams(MASK_VAL) Then: Let Opcode = Y2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
     
iEI:
    'Enable Interrupts: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Y7 Or Z3: GoTo Output
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
    '`EX DE, HL`
    ElseIf TestParams(MASK_REG_DE, MASK_REG_HL) Then
        Let Opcode = X3 Or Y5 Or Z3
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iEXX:
    'Exchange All: _
     ..................................................................................
    If TestParams() Then Let Opcode = X3 Or Q1 Or P1 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iHALT:
    'Halt: Stop the processor and wait for interrupts _
     ..................................................................................
    If TestParams() Then Let Opcode = X1 Or Y6 Or Z6: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iIM:
    'Interrupt Mode: _
     ..................................................................................
    '`IM 0|1|2`
    If TestParams(MASK_VAL) Then
        If Z80Param1.Value = 0 Then
            Let Opcode = X1 Or Z6
        ElseIf Z80Param1.Value = 1 Then
            Let Opcode = X1 Or Y2 Or Z6
        ElseIf Z80Param1.Value = 2 Then
            Let Opcode = X1 Or Y3 Or Z6
        Else
            'ERROR: Wrong parameters for Z80 instruction
            GoTo InvalidParams
        End If
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iIN:
    'Read from port: _
     ..................................................................................
    '`IN A, ($8)`
    If TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = X3 Or Y3 Or Z3
    '`IN A|B|C|D|E|H|L, [C]`
    ElseIf TestParams(MASK_REGS_ABCDEHL, MASK_MEM_BC) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask)
    '`IN (C)`
    ElseIf TestParams(MASK_MEM_BC) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Y6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iINC:
    'Increment: _
     ..................................................................................
    '`INC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Z3
    '`INC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z4
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iIND:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y5 Or Z2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINDR:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y7 Or Z2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iINI:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y4 Or Z2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINIR:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y6 Or Z2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iJP:
    'Jump: _
     ..................................................................................
    '`JP [HL|IX|IY]` or `JP HL|IX|IY` - indirect absolute jump
    If TestParams(MASK_MEM_HLIXY Or MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P2 Or Q1 Or Z1
    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
    ElseIf TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token) Or Z2: GoTo Output16
    '`JP $16` - absolute jump
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Z3: GoTo Output16
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iJR:
    'Jump Relative: _
     ..................................................................................
    '`JR $8` - relative immediate jump
    If TestParams(MASK_VAL) Then
        Let Opcode = Y3
    '`JR NZ|Z|NC|C, $8` - conditional relative immediate jump
    ElseIf TestParams(MASK_FLAGS_CZ, MASK_VAL) Then
        Let Opcode = OpcodeY_Condition(Z80Param1.Token) + Y4
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iLD:
    'The Mother-Load: _
     ..................................................................................
    '`LD BC|DE|HL|SP|IX|IY, $16`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY, MASK_VAL) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Z1: GoTo Output16
    '`LD [BC], A`
    ElseIf TestParams(MASK_MEM_BC, MASK_REG_A) Then
        Let Opcode = Z2
    '`LD [DE], A`
    ElseIf TestParams(MASK_MEM_DE, MASK_REG_A) Then
        Let Opcode = P1 Or Z2
    '`LD A, [BC]`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_BC) Then
        Let Opcode = Q1 Or Z2
    '`LD A, [DE]`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_DE) Then
        Let Opcode = P1 Or Q1 Or Z2
    '`LD [$16], HL|IX|IY`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_HL_IXY) Then
        Let Opcode = P2 Or Z2: GoTo Output16
    '`LD [$16], A`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = P3 Or Z2: GoTo Output16
    '`LD HL|IX|IY, [$16]`
    ElseIf TestParams(MASK_REGS_HL_IXY, MASK_MEM_VAL) Then
        Let Opcode = P2 Or Q1 Or Z2: GoTo Output16
    '`LD A, [$16]`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = P3 Or Q1 Or Z2: GoTo Output16
    '`LD A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8], $8`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY, MASK_VAL) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z6
    'You can't combine IX and IY use in the same opcode, there's only one prefix, _
     so these two following checks ensure that IX & IY are not used together
    '`LD A|B|C|D|E|H|IXH|L|IXL[HL|IX+$8], A|B|C|D|E|H|IXH|L|IXL`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD A|B|C|D|E|H|IYH|L|IYL[HL|IY+$8], A|B|C|D|E|H|IYH|L|IYL`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    'You can't have `LD [IX+$8], [IX+$8]!
    '`LD A|B|C|D|E|H|IXH|L|IXL, A|B|C|D|E|H|IXH|L|IXL|[HL|IX+$8]`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD A|B|C|D|E|H|IYH|L|IYL, A|B|C|D|E|H|IYH|L|IYL|[HL|IY+$8]`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD SP, HL|IX|IY`
    ElseIf TestParams(MASK_REG_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P3 Or Q1 Or Z1
    '`LD [$16], BC|DE|SP`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_BC_DE_SP) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param2.Token) Or Z3: GoTo Output16
    '`LD BC|DE|SP, ($16)`
    ElseIf TestParams(MASK_REGS_BC_DE_SP, MASK_MEM_VAL) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Q1 Or Z3
        GoTo Output16
    '`LD I, A`
    ElseIf TestParams(MASK_REG_I, MASK_REG_A) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Z7
    '`LD R, A`
    ElseIf TestParams(MASK_REG_R, MASK_REG_A) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Y1 Or Z7
    '`LD A, I`
    ElseIf TestParams(MASK_REG_A, MASK_REG_I) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Y2 Or Z7
    '`LD A, R`
    ElseIf TestParams(MASK_REG_A, MASK_REG_R) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Y3 Or Z7
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iLDD:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y5: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iLDDR:
    '..................................................................................
    If TestParams() Then: Let Prefix = &HED&: Let Opcode = X2 Or Y7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iLDI:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y4: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iLDIR:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y6: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iNEG:
    'Negate the Accumulator: _
     ..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X1 Or Z4: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iNOP:
    'No Operation: _
     ..................................................................................
    '`NOP` - Opcode is 0
    If TestParams() Then: Let Opcode = 0: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOR:
    'Bitwise OR: _
     ..................................................................................
    '`OR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y6 Or OpcodeZ_Register(Z80Param1.Mask)
    '`OR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y6 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iOUT:
    'Write to port: _
     ..................................................................................
    '`OUT [$8], A`
    If TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = X3 Or Y2 Or Z3
    '`OUT [C], A|B|C|D|E|H|L`
    ElseIf TestParams(MASK_MEM_BC, MASK_REGS_ABCDEHL) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or OpcodeY_Register(Z80Param2.Mask) Or Z1
    '`OUT [C], 0`
    ElseIf TestParams(MASK_MEM_BC, MASK_VAL) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or Y6 Or Z1
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iOUTD:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y5 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTDR:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y7 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iOUTI:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y4 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTIR:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X2 Or Y6 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iPOP:
    'Fetch from stack: _
     ..................................................................................
    '`POP AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Z1
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iPUSH:
    'Push a register on to the stack: _
     ..................................................................................
    '`PUSH AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Z5
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iRES:
    'Reset a bit: _
     ..................................................................................
    '`RES 0-7 A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`RES 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RES memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X2 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRET:
    'Return: _
     ..................................................................................
    '`RET NZ|Z|NC|C|PO|PE|P|M` - conditional return
    If TestParams(MASK_FLAGS) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token)
    '`RET`
    ElseIf TestParams() Then
        Let Opcode = X3 Or Q1 Or Z1
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iRETI:
    'Return from Interrupt: _
     ..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X1 Or Y1 Or Z5: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRETN:
    'Return from NMI: _
     ..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X1 Or Z5: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y2 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRL:
    '..................................................................................
    '`RL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y2 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y2 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iRLCA:
    '..................................................................................
    If TestParams() Then Let Opcode = Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLC:
    '..................................................................................
    '`RLC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = OpcodeZ_Register(Z80Param1.Mask)
    '`RLC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RLC memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRLD:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X1 Or Y5 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRRA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y3 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRR:
    '..................................................................................
    '`RR A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y3 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RR [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RR memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y3 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iRRCA:
    '..................................................................................
    If TestParams() Then Let Opcode = Y1 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRRC:
    '..................................................................................
    '`RRC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y1 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RRC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RRC memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y1 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRRD:
    '..................................................................................
    If TestParams() Then Let Prefix = &HED&: Let Opcode = X1 Or Y4 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRST:
    '"Reset" - Fire specific interrupt vector: _
     ..................................................................................
    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
    If TestParams(MASK_VAL) Then
        'TODO: Need to validate the number is in range at some point in the chain
        Let Opcode = X3 Or OpcodeY_RST(Z80Param1.Value) Or Z7
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iSBC:
    '..................................................................................
    '`SBC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y3 Or OpcodeZ_Register(Z80Param2.Mask)
    '`SBC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y3 Or Z6
    '`SBC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Prefix = &HED&
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param2.Token) Or Z2
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iSCF:
    '..................................................................................
    If TestParams() Then Let Opcode = Y6 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iSET:
    'Set a bit: _
     ..................................................................................
    '`SET 0-7 A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X3 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`SET 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SET memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X3 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSLA:
    '..................................................................................
    '`SLA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y4 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SLA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLA memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y4 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iSRA:
    '..................................................................................
    '`SRA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y5 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SRA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRA memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y5 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSLL:
    '..................................................................................
    '`SLL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y6 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SLL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y6 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSRL:
    '..................................................................................
    '`SRL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y7 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SRL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y7 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iSUB:
    'Subtract: _
     ..................................................................................
    '`SUB A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y2 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SUB $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y2 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iXOR:
    'Bitwise XOR: _
     ..................................................................................
    '`XOR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y5 Or OpcodeZ_Register(Z80Param1.Mask)
    '`XOR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y5 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If

    '----------------------------------------------------------------------------------
    'Write the assembled bytes to the Chunk
    '----------------------------------------------------------------------------------
Output:
    'TODO: Validate 8-bit Value
    Call Chunk.AddInstruction( _
        Prefix, Opcode, Offset, Value _
    )
    
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_PROCESSZ80_LOG)
    Call Chunk.LogASM.Add(LOG_Z80_VAL8, _
        LogAdr, _
        oz80.HexStr8(Prefix), oz80.HexStr8(Opcode), oz80.HexStr8(Offset), _
        oz80.HexStr8(Value), _
        oz80.TokenName(Instruction), _
        Param1Str, Param2Str, Param3Str _
    )
    Call oz80.Profiler.EventEnd(EVENT_PROCESSZ80_LOG)
    '*** TIMER END ***
    
    GoTo Finish
    
Output16:
    '..................................................................................
    'TODO: Validate 16-bit Value
    'TODO: 16-bit Value opcodes never have an offset
    Call Chunk.AddInstruction( _
        Prefix, Opcode, Offset, _
        (Value And &HFF&), ((Value \ &H100&) And &HFF&) _
    )
    
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_PROCESSZ80_LOG)
    Call Chunk.LogASM.Add(LOG_Z80_VAL16, _
        LogAdr, _
        oz80.HexStr8(Prefix), oz80.HexStr8(Opcode), _
        oz80.HexStr8(Value And &HFF&), _
        oz80.HexStr8((Value And &HFF00&) \ &H100&), _
        oz80.TokenName(Instruction), _
        Param1Str, Param2Str _
    )
    Call oz80.Profiler.EventEnd(EVENT_PROCESSZ80_LOG)
    '*** TIMER END ***
    
    GoTo Finish
    
OutputCB:
    '..................................................................................
    'The double-opcodes use $CB as the primary opcode and then a second opcode after _
     the offset byte. It's fastest to handle and log this special case specifically
    Call Chunk.AddInstructionCB(Prefix, Opcode, Offset)
    
    '*** TIMER BEGIN ***
    Call oz80.Profiler.EventBegin(EVENT_PROCESSZ80_LOG)
    Call Chunk.LogASM.Add(LOG_Z80_CB, _
        LogAdr, _
        oz80.HexStr8(Prefix), oz80.HexStr8(Offset), oz80.HexStr8(Opcode), _
        oz80.TokenName(Instruction), _
        Param1Str, Param2Str, Param3Str _
    )
    Call oz80.Profiler.EventEnd(EVENT_PROCESSZ80_LOG)
    '*** TIMER END ***
    GoTo Finish
    
InvalidParams:
    '----------------------------------------------------------------------------------
    'ERROR: Wrong parameters for Z80 instruction
    Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
Finish:
    'Return the error number
    Let AssembleZ80 = My_Error
    Exit Function
End Function

'TestParams : Check if parameters are of the right type desired
'======================================================================================
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
        If (Z80Param1.Mask <> 0) Then Exit Function
        'If there is no first parameter, pass True
        If (Z80Param1.Mask = 0) Then GoTo Finish
    End If
    
    '1st Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param1.Mask And Test1) = 0 Then Exit Function
    
    'Are we testing a second parameter?
    If (Test2 = 0) Then
        'If there shouldn't be a second parameter, return False
        If (Z80Param2.Mask <> 0) Then Exit Function
        'If there is no second parameter, the test has passed
        If (Z80Param2.Mask = 0) Then GoTo Finish
    End If
    
    '2nd Parameter Test: _
     ----------------------------------------------------------------------------------
        
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param2.Mask And Test2) = 0 Then Exit Function
    
    'If there shouldn't be a third parameter, return False
    If (Test3 = 0) And (Z80Param3.Mask <> 0) Then Exit Function
    'If there is no third parameter, the test has passed
    If (Test3 = 0) And (Z80Param3.Mask = 0) Then GoTo Finish
    
    '3rd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param3.Mask And Test3) = 0 Then Exit Function
    
Finish:
    Let TestParams = True
End Function

'GetListAsDictionary : Reads a whole list into a Dictionary
'======================================================================================
Private Function GetListAsDictionary() As Scripting.Dictionary
    'Instantiate the return object
    Set GetListAsDictionary = New Scripting.Dictionary
    
    'As long as the list continues, keep fetching Values
    Do While File.TokenIsExpression
        'Parse the Expression; if it's indefinite, _
         leave that up to the caller to handle
        Dim Value As Double
        'If there was an error in the expression, the error event _
         will already have been raised, we only need to fold upwards
        If ParseExpression(Value) Then Exit Function
        
        Call GetListAsDictionary.Add( _
            GetListAsDictionary.Count, Value _
        )
    Loop
End Function

'GetText : Reads a list, interpretting Values as Unicode points
'======================================================================================
'ReturnText     | A bluString object that will be populated with the text
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function GetText( _
    ByRef ReturnText As bluString _
) As OZ80_ERROR
    'If the bluString is not instantiated, do so now
    If ReturnText Is Nothing Then Set ReturnText = New bluString
    
    'If the current token is not a list item, then the list has come to an end
    Do While File.TokenIsList
    
        'If the current token is text, append it to what we have
        If File.TokenKind = TOKEN_TEXT Then
            Call ReturnText.Append(File.TokenText)
        
        'Otherwise it'll be a calculable Expression, _
         hopefully providing an ASCII value
        ElseIf File.TokenIsExpression Then
            'Get the Expression Value
            Dim Value As Double
            'If there were an error in parsing the Expression, fold up
            If ParseExpression(Value) Then GoTo ErrOut
            
            'TODO: Indefnite values? Allowed or not in Text? _
                   If we allow for text mapping with different sized strings then _
                   the string size cannot be determinate
            
            'Any UCS-2 value is allowed, it's up to the caller _
             to do text-mapping and validation
            If (Value < 0) Or (Value > 65535) Then
                Call RaiseError(OZ80_ERROR_TEXT_CHAR)
                GoTo ErrOut
            End If
        End If
    
        'Continue reading the list. Note that when the file ends, so does the list _
         and this is not an error as far as this function is concerned. If the parent _
         is within a scope and expecting the closing bracket, then it's up to them _
         to raise the error
        Call File.NextToken
    Loop
    
    'The resultant string won't be changing any more
    Call ReturnText.ReduceBuffer
ErrOut:
    'Return the error number
    Let GetText = My_Error
End Function

'GetListItem : Read in an Expression from a list
'======================================================================================
Private Function GetListItem( _
             ByRef Value As Double, _
    Optional ByRef Length As Long = 0 _
) As Boolean
    'If the current token is not an expression, the list has ended; return False
    If Not File.TokenIsExpression Then Exit Function
    
    'Parse the Expression; if it's indefinite, _
     leave that up to the caller to handle
    '(If there was an error in the Expression, the error event _
      will already have been raised, we only need to fold upwards)
    If ParseExpression(Value) Then Exit Function
    
    'Increases the number of items read from the list thus far
    Let Length = Length + 1
    
    'Return True as the list has not ended yet
    Let GetListItem = True
End Function

'ParseExpression : Evaluate a calculation into a Value
'======================================================================================
'ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
'               | Note that this is a 64-bit floating-pointer number, so it's up to
'               | the caller to validate the number according to their needs
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function ParseExpression( _
    ByRef ReturnValue As Double _
) As OZ80_ERROR
    Dim Operand1 As Double              'Left-hand value in a calculation
    Dim Operand2 As Double              'Right-hand value in a calculation
    Dim Operator As OZ80_TOKEN          'Operator to use
    
    'Begin with fetching the first operand
    If ParseValue(Operand1) Then GoTo ErrOut
    'For the moment the Expression's calculated Value will be this unless we find _
     an Operator to extend the Expression
    Let ReturnValue = Operand1
    
    'The file could legitimately end with an Expression, so don't error out if the _
     token stream comes to an end since the Expression is valid. The parent scopes _
     will handle an unexpected end error
    If File.EOF Then Exit Function
    
    'If next token is not an Operator assume the Expression has finished, _
      the parent will handle whatever the next token is used for
    If Not File.TokenIsOperator Then Exit Function
    
Continue:
    'Remember the Operator for after we have operand 2 to work with
    Let Operator = File.TokenKind
    
    'Get the second operand
    'TODO: Error - unexpected end
    If Not File.NextToken() Then Stop
    'Cannot have an operator hanging there without an operand following
    If ParseValue(Operand2) Then Stop
        
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If (Operand1 = INDEFINITE) _
    Or (Operand2 = INDEFINITE) Then
        'Mark expression result as indefinite and skip applying the operator
        Let ReturnValue = INDEFINITE
    
    ElseIf Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 + Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 - Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 * Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        Let ReturnValue = Operand1 / Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let ReturnValue = Operand1 ^ Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        Let ReturnValue = Operand1 Mod Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        Let ReturnValue = Operand1 And Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        Let ReturnValue = Operand1 Or Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_XOR Then
        Let ReturnValue = Operand1 Xor Operand2
        
    Else
        Stop
    End If
    
    'Does the expression continue?
    If File.TokenIsOperator Then
        'Set the current result as the first operand and go around again
        Let Operand1 = ReturnValue
        GoTo Continue
    End If
    
ErrOut:
    'Return the error number
    Let ParseExpression = My_Error
End Function

'ParseValue : Get a single Value as part of an Expression
'======================================================================================
'ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
'               | Note that this is a 64-bit floating-pointer number, so it's up to
'               | the caller to validate the number according to their needs
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function ParseValue( _
    ByRef ReturnValue As Double _
) As OZ80_ERROR
    '----------------------------------------------------------------------------------
    'A nested Expression?
    '----------------------------------------------------------------------------------
    If File.TokenKind = TOKEN_PARENS_OPEN Then
        'Move over the open parentheses. If the file ends here, it's an error
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Parse the nested Expression and return its Value as our own
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        
        'The nested Expression has to be closed
        If File.TokenKind <> TOKEN_PARENS_CLOSE Then
            Call RaiseError(OZ80_ERROR_EXPECTED_BRACKET)
            GoTo ErrOut
        End If
        
        'Move over the closing brace. The Expression can end here, so it's not _
         necessarily an error if the file suddenly ends. The parent can check for _
         any scopes remaining open
        Call File.NextToken
        
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'The `HI` function -- returns the high-byte of a 16-bit Value
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_HI Then
        'Move over the `HI` keyword. If the file ends here, it's an error
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        'Fetch the parameter to the `HI` function
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        'Check that this is no more than a 16-bit number
        If ReturnValue > 65535 Then Call RaiseError(OZ80_ERROR_OVERFLOW_HILO): GoTo ErrOut
        'Extract the high-byte
        Let ReturnValue = (ReturnValue And &HFF00&) \ &H100&
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'The `LO` function -- returns the low-byte of a 16-bit Value
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_LO Then
        'Move over the `LO` keyword. If the file ends here, it's an error
        If Not File.NextToken Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        'Fetch the parameter to the `LO` function
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        'Check that this is no more than a 16-bit number _
         (despite this not being an issue to the low 8-bits, we prevent numbers _
          higher than 16-bit to protect against unintended behaviour)
        If ReturnValue > 65535 Then Call RaiseError(OZ80_ERROR_OVERFLOW_HILO): GoTo ErrOut
        'Extract the low-byte
        Let ReturnValue = ReturnValue And &HFF&
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'Negative Numbers
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_OPERATOR_SUB Then
        'If the negative sign preceeds, remember this until the full number has _
         been calculated. ONLY decimal numbers can be negative, you cannot negate _
         hexadecimal/binary numbers or named-items (Constants, RAM names, Labels etc.)
        Dim IsNegative As Boolean: Let IsNegative = True
        'Move forward past the negative sign. If the file ends here, it's an error
        If Not File.NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        'Continue parsing the number...
        GoTo GetNumber
        
    '----------------------------------------------------------------------------------
    'A decimal number:
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_NUMBER Then
GetNumber:
        'Begin with getting the given value
        Let ReturnValue = File.TokenValue
        'Was there a negative sign preceeding this number?
        If IsNegative Then Let ReturnValue = -ReturnValue
        
        'Move over the number. If the file ends here, the Value is still valid, _
         so no error. The parent can check if any scope has been left open. We skip _
         checking for a number prefix, but we do need to apply any negative sign
        If File.NextToken Then
            'Check for a number prefix, e.g. `KB`
            '..........................................................................
            If File.TokenKind = TOKEN_PREFIX_K Then
                'Check for min/max number you can multiply by 1'000 without overflowing
                If (ReturnValue < NUMBER_MIN_K) Or (ReturnValue > NUMBER_MAX_K) Then
                    Call RaiseError(OZ80_ERROR_OVERFLOW)
                    GoTo ErrOut
                End If
                'Multiply by 1'000
                Let ReturnValue = ReturnValue * 1000
                'Move over the prefix
                Call File.NextToken
                
            '..........................................................................
            ElseIf File.TokenKind = TOKEN_PREFIX_KB Then
                'Check for min/max number you can multiply by 1'024 without overflowing
                If (ReturnValue < NUMBER_MIN_KB) Or (ReturnValue > NUMBER_MAX_KB) Then
                    Call RaiseError(OZ80_ERROR_OVERFLOW)
                    GoTo ErrOut
                End If
                'Multiply by 1'024
                Let ReturnValue = ReturnValue * 1024
            Else
                'If the token is not a prefix, exit this function without moving the _
                 token pointer forward (as we've already done that)
                GoTo ErrOut
            End If
        End If
    
    '----------------------------------------------------------------------------------
    'Constants
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_CONST Then
        'Has the constant been defined?
        If Not Values.Exists(File.TokenText) Then
            'ERROR: Constant not yet defined. Constants cannot be indefinite
            Call RaiseError(OZ80_ERROR_UNDEFINED_CONST)
            GoTo ErrOut
        End If
        'Retrieve the Constant value
        Let ReturnValue = Values.Item(File.TokenText)
        
    '----------------------------------------------------------------------------------
    'RAM Names
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_RAM _
        Or File.TokenKind = TOKEN_HASH _
    Then
        'Has the RAM Name / Hash value been defined?
        If Not Values.Exists(File.TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the RAM address / Has Value
            Let ReturnValue = Values.Item(File.TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    'The `START` keyword refers to the System's boot vector
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_KEYWORD_START Then
        'TODO: This is determined by the System, so we can forgo having to reference it
        If Not Labels.Exists(STR_START) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the START address
            Let ReturnValue = Labels.Item(STR_START)
        End If
    
    '----------------------------------------------------------------------------------
    'Labels
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_LABEL Then
        'Label names have to be expanded to their fully-qualified form
        'TODO: If a Label is not found in this Section, check the global Section
        If Not Labels.Exists(CurrentSection & File.TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the RAM address
            Let ReturnValue = Labels.Item(CurrentSection & File.TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    'Properties
    '----------------------------------------------------------------------------------
    ElseIf File.TokenKind = TOKEN_PROPERTY_USE Then
        'Property names have to be expanded to their fully-qualified form
        If Not Labels.Exists(CurrentSection & CurrentLabel & File.TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the RAM address
            Let ReturnValue = Labels.Item(CurrentSection & CurrentLabel & File.TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    Else
        Debug.Print oz80.TokenName(File.TokenKind)
        Stop
    End If
    
    'Move over the processed token, the parent expects the token pointer _
     to be on the next unread token
    Call File.NextToken
ErrOut:
    'Return the error number
    Let ParseValue = My_Error
End Function

'BeginChunk : Begin a new Chunk to hold assembled Z80 code / data
'======================================================================================
Private Sub BeginChunk( _
             ByRef LabelName As String, _
    Optional ByRef Address As Double = -1 _
)
    'Register the Label name in the database
    Call AddUpdateLabel(LabelName, Address)
    'Does the Chunk already exist? On the first pass we want to ensure that the user _
     doesn't try and define two Chunks with the same name, and this is handled by the _
     `PROC` handler. On the second pass the Chunks will already have been defined
    If ChunkExists(LabelName) Then
        'Attach the existing Chunk
        Set Chunk = GetChunk(LabelName)
    Else
        'Create and attach the new chunk
        Set Chunk = New oz80_Chunk
        Let Chunk.LabelName = LabelName
        Let Chunk.Address = CLng(Address)
        
        Let ChunkCount = ChunkCount + 1
        ReDim Preserve Chunks(1 To ChunkCount) As oz80_Chunk
        Set Chunks(ChunkCount) = Chunk
    End If
End Sub

'AddUpdateLabel: Blindly adds to or updates the Label database
'======================================================================================
'LabelName      | The fully-qualified Label name, i.e. including the Section name
'[Value]        | The new address to set for the Label
'======================================================================================
Private Sub AddUpdateLabel( _
             ByRef LabelName As String, _
    Optional ByRef Value As Double = 0 _
)
    'Add a new label or update an existing label value?
    If Not Labels.Exists(LabelName) _
        Then Call Labels.Add(LabelName, Value) _
        Else Let Labels.Item(LabelName) = Value
End Sub

'Layout
'======================================================================================
Private Function Layout( _
) As OZ80_ERROR
    '
End Function

'/// GENERIC //////////////////////////////////////////////////////////////////////////
'/// Functions used anywhere throughout the clas

'ChunkExists : Test if a Chunk is in the array by the Label name
'======================================================================================
Private Property Get ChunkExists( _
    ByRef LabelName As String _
) As Boolean
    'If there are no Chunks, just return False
    If ChunkCount = 0 Then Exit Property
    'Go over each Chunk and check the name
    Dim i As Long
    For i = 1 To ChunkCount
        If Chunks(i).LabelName = LabelName Then
            Let ChunkExists = True
            Exit Property
        End If
    Next i
End Property

'GetChunk : Retrieves a Chunk in the array by its Label Name
'======================================================================================
Private Function GetChunk( _
    ByRef LabelName As String _
) As oz80_Chunk
    'Go over each Chunk and check the name
    Dim i As Long
    For i = 1 To ChunkCount
        If Chunks(i).LabelName = LabelName Then
            Set GetChunk = Chunks(i)
            Exit Function
        End If
    Next i
End Function

'ParamToString : Get a textual representation of a Z80 instruction parameter
'======================================================================================
'Param          | A Parameter Obect that will contain the details
'---------------+----------------------------------------------------------------------
'Returns        | A string with the Parameter's name
'======================================================================================
Private Function ParamToString( _
    ByRef Param As oz80Param _
) As String
    If Param.Mask = MASK_REG_A Then
        Let ParamToString = "A"
    ElseIf Param.Mask = MASK_REG_B Then Let ParamToString = "B"
    ElseIf Param.Mask = MASK_REG_C Then Let ParamToString = "C"
    ElseIf Param.Mask = MASK_REG_D Then Let ParamToString = "D"
    ElseIf Param.Mask = MASK_REG_E Then Let ParamToString = "E"
    ElseIf Param.Mask = MASK_REG_H Then Let ParamToString = "H"
    ElseIf Param.Mask = MASK_REG_L Then Let ParamToString = "L"
    ElseIf Param.Mask = MASK_REG_I Then Let ParamToString = "I"
    ElseIf Param.Mask = MASK_REG_R Then Let ParamToString = "R"
    ElseIf Param.Mask = MASK_REG_AF Then Let ParamToString = "AF"
    ElseIf Param.Mask = MASK_REG_BC Then Let ParamToString = "BC"
    ElseIf Param.Mask = MASK_REG_DE Then Let ParamToString = "DE"
    ElseIf Param.Mask = MASK_REG_HL Then Let ParamToString = "HL"
    ElseIf Param.Mask = MASK_REG_SP Then Let ParamToString = "SP"
    ElseIf Param.Mask = MASK_REG_IX Then Let ParamToString = "IX"
    ElseIf Param.Mask = MASK_REG_IXL Then Let ParamToString = "IXL"
    ElseIf Param.Mask = MASK_REG_IXH Then Let ParamToString = "IXH"
    ElseIf Param.Mask = MASK_REG_IY Then Let ParamToString = "IY"
    ElseIf Param.Mask = MASK_REG_IYL Then Let ParamToString = "IYL"
    ElseIf Param.Mask = MASK_REG_IYH Then Let ParamToString = "IYH"
    
    ElseIf Param.Mask = MASK_VAL Then
        If Param.Value > 255 _
            Then Let ParamToString = "$" & oz80.HexStr16(Param.Value) _
            Else Let ParamToString = "$" & oz80.HexStr8(Param.Value)
        
    'The mask bits do not specify every flag, _
     we refer to the token kind for that
    ElseIf (Param.Mask And MASK_FLAGS) <> 0 Then
        If Param.Token = TOKEN_Z80_C Then
            Let ParamToString = "C"
        ElseIf Param.Token = TOKEN_Z80_NC Then Let ParamToString = "NC"
        ElseIf Param.Token = TOKEN_Z80_Z Then Let ParamToString = "Z"
        ElseIf Param.Token = TOKEN_Z80_NZ Then Let ParamToString = "NZ"
        ElseIf Param.Token = TOKEN_Z80_P Then Let ParamToString = "P"
        ElseIf Param.Token = TOKEN_Z80_PE Then Let ParamToString = "PE"
        ElseIf Param.Token = TOKEN_Z80_PO Then Let ParamToString = "PO"
        ElseIf Param.Token = TOKEN_Z80_M Then Let ParamToString = "M"
        End If
    
    'Memory references
    ElseIf Param.Mask = MASK_MEM_HL Then
        Let ParamToString = "(HL)"
    ElseIf Param.Mask = MASK_MEM_IX Then
        If Param.Value >= 0 _
            Then Let ParamToString = "[IX+$" & oz80.HexStr8(Param.Value) & "]" _
            Else Let ParamToString = "[IX]"
    ElseIf Param.Mask = MASK_MEM_IY Then
        If Param.Value >= 0 _
            Then Let ParamToString = "[IY+$" & oz80.HexStr8(Param.Value) & "]" _
            Else Let ParamToString = "[IY]"
    ElseIf Param.Mask = MASK_MEM_BC Then Let ParamToString = "[BC]"
    ElseIf Param.Mask = MASK_MEM_DE Then Let ParamToString = "[DE]"
    ElseIf Param.Mask = MASK_MEM_SP Then Let ParamToString = "[SP]"
    ElseIf Param.Mask = MASK_MEM_VAL Then
        If Param.Value > 255 _
            Then Let ParamToString = "[$" & oz80.HexStr16(Param.Value) & "]" _
            Else Let ParamToString = "[$" & oz80.HexStr8(Param.Value) & "]"
    Else
        Let ParamToString = oz80.TokenName(Param.Token)
    End If
End Function

'RaiseMessage : Provide a progress message
'======================================================================================
'LogLevel       | The type of log message, for filtering out lower-level messages
'FormatText     | The message to log, which can include "{$}" markers to be replaced
'               | with the values provided below
'FormatValues...| A list of Variant values (Strings, Longs &c.) to fill in
'======================================================================================
Private Sub RaiseMessage( _
    ByRef LogLevel As OZ80_LOG, _
    ByRef FormatText As String, _
    ParamArray FormatValues() As Variant _
)
    Let LogMsg.Text = FormatText
    
    'Are there any values to format?
    If UBound(FormatValues) < 0 Then
        'No, raise just the text given
        RaiseEvent Message(LogLevel, LogMsg)
    Else
        'Format the given log message before sending it up
        RaiseEvent Message(LogLevel, LogMsg.Format(FormatValues))
    End If
End Sub

'RiaseError : Provide error details through an event
'======================================================================================
'Number         | Error number to set on the Class and raise via Event
'[Line]         | Line number of the source file the error occurred on
'[Col]          | Column number of the source file the error occurred on
'---------------+----------------------------------------------------------------------
'Returns        | The same error number, for simple tail recursion
'======================================================================================
Private Function RaiseError( _
             ByRef Number As OZ80_ERROR, _
    Optional ByRef Line As Long = -1, Optional ByRef Col As Long = -1 _
) As OZ80_ERROR
    'Set the error number on this class
    Let My_Error = Number
    
    If Line < 0 Then Let Line = File.TokenLine
    If Col < 0 Then Let Col = File.TokenCol
    
    'Fetch the detailed error text
    Dim Title As String, Description As String
    Call oz80.GetOZ80Error(Number, Title, Description)
    
    'Pass the error up to the controller
    RaiseEvent Error(File.Path, Number, Title, Description, Line, Col)
    
    'We return the error number, purely for nicer looking tail-recursion
    Let RaiseError = Number
End Function
