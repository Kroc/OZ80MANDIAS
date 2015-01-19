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
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: oz80_Chunk

'Dependencies       oz80.bas _
                    bluString.bas [bluArrayStrings.bas, <bluArray.bas>]

'--------------------------------------------------------------------------------------

'Code and/or data is assembled in separate Chunks (this class) which are arranged _
 according to Section layout before being written out to the ROM. _
 i.e. each `PROC` / `TABLE` will be its own Chunk

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535(v=vs.85).aspx>
Private Declare Sub api_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'/// DATA /////////////////////////////////////////////////////////////////////////////

'For ease of logging, we'll embed the printable name of the Chunk, _
 (the actual name is stored in the File, well out of reach of many routines)
Public LabelName As String

'Which section the chunk belongs to. This will be the CRC32 of the Section name, _
 as stored in oz80_Assembler. Value will be 0 for no particular Section
Public Section As Long

'The absolute position of the chunk in the ROM. _
 This will be assigned after the first assembly phase when all Chunk sizes are known _
 and they can be placed in their given Sections
Public Address As Long

'The assembled code/data goes here
Private Data() As Byte

'Length of the chunk in bytes
Private My_Size As Long

'Each Chunk starts at this size (many Chunks will be small Z80 routines). _
 If this buffer size is filled, it will be doubled. Once the Chunk is "finalised", _
 the buffer will be cut down to the size of the actual contents
Private Const BUFFER_SIZE As Long = 1024
Private BufferSize As Long

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    'Initialise the space where the assembled bytes go. _
     This chunk of bytes will be copied into the ROM once the location has been chosen
    Let BufferSize = BUFFER_SIZE
    ReDim Data(0 To BufferSize - 1) As Byte
    
    'The ROM address of the Chunk can of course be 0, so default to indefinite
    Let Me.Address = -1
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Erase Data
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Size
'======================================================================================
Public Property Get Size() As Long: Let Size = My_Size: End Property

'AddByte : Add a byte to the assembled output and move along
'======================================================================================
Public Sub AddByte( _
    ByRef Value As Long _
)
    'Add the byte to the ROM and move to the next byte
    Let Data(My_Size) = Value And &HFF
    Let My_Size = My_Size + 1
    
    'If the chunk's buffer is full then double the buffer size to compensate. _
     This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, 16 KB
    If My_Size >= BufferSize Then
        Let BufferSize = BufferSize + BufferSize
        ReDim Preserve Data(0 To BufferSize - 1) As Byte
    End If
End Sub

'AddWord : Add a 2-byte value, in little-endian order
'======================================================================================
Public Sub AddWord( _
    ByRef Value As Long _
)
    'Add the low-byte first _
     (the `AddByte` routine clips to 8-bit automatically)
    Call AddByte(Value)
    'And the high-byte
    Call AddByte((Value And &HFF00&) \ &H100)
End Sub

'AddInstruction : Adds a Z80 instruction and parameters to the Chunk
'======================================================================================
'Prefix         | $DD for IX Register, $FD for IY Register, $ED for some opcodes
'Opcode         | Z80 Instruction opcode
'Offset         | 8-Bit Index for IX & IY memory references, e.g. "(IX+$FF)"
'ValueLo        | 8-Bit value or the Low-byte of a 16-bit value
'ValueHi        | High-byte of a 16-bit value
'======================================================================================
Public Sub AddInstruction( _
    Optional ByRef Prefix As Long = -1, _
    Optional ByRef Opcode As Long = -1, _
    Optional ByRef Offset As Long = -1, _
    Optional ByRef ValueLo As Long = -1, _
    Optional ByRef ValueHi As Long = -1 _
)
    If Prefix <> -1 Then Call AddByte(Prefix)
    If Opcode <> -1 Then Call AddByte(Opcode)
    If Offset <> -1 Then Call AddByte(Offset)
    If ValueLo <> -1 Then Call AddByte(ValueLo)
    If ValueHi <> -1 Then Call AddByte(ValueHi)
End Sub

'AddInstructionCB : Adds a CB-opcode Z80 instruction and parameters to the Chunk
'======================================================================================
'Prefix         | $DD for IX Register, $FD for IY Register, $ED for some opcodes
'Opcode         | Z80 Instruction opcode
'Offset         | 8-Bit Index for IX & IY memory references, e.g. "(IX+$FF)"
'======================================================================================
Public Sub AddInstructionCB( _
    Optional ByRef Prefix As Long = -1, _
    Optional ByRef Opcode As Long = -1, _
    Optional ByRef Offset As Long = -1 _
)
    If Prefix <> -1 Then Call AddByte(Prefix)
    AddByte (&HCB&)
    If Offset <> -1 Then Call AddByte(Offset)
    If Opcode <> -1 Then Call AddByte(Opcode)
End Sub

'Trim
'======================================================================================
Public Sub Trim()
    'Chunks can be empty!
    Dim NewSize As Long
    If My_Size = 0 Then Let NewSize = 0 Else Let NewSize = My_Size - 1
    'Trim the buffer size down to the actual size
    ReDim Preserve Data(0 To NewSize) As Byte
    'Update the buffer size to match
    Let BufferSize = NewSize
End Sub

'CopyInto
'======================================================================================
Public Sub CopyInto( _
    ByRef OutputArray() As Byte _
)
    'Check that it'll fit! We'll need to increase the ROM size if necessary
    If Address + My_Size > UBound(OutputArray) Then
        'Double the size of the ROM!
        ReDim Preserve OutputArray(((UBound(OutputArray) + 1) * 2) - 1) As Byte
        'TODO: There's a maximum size of the ROM
    End If
    
    Call api_RtlMoveMemory( _
        OutputArray(Address), Data(0), My_Size _
    )
End Sub
