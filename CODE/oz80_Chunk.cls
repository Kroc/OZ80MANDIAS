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
 i.e. each `PROC` / `TABLE` will be its own Chunk

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub api_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef ptrDestination As Any, _
    ByRef ptrSource As Any, _
    ByVal Length As Long _
)

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'For ease of logging, we'll embed the printable name of the Chunk, _
 (the actual name is stored in the TokenStream, well out of reach of many routines)
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

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise the space where the assembled bytes go. _
     This chunk of bytes will be copied into the ROM once the location has been chosen
    Let BufferSize = BUFFER_SIZE
    ReDim Data(0 To BufferSize - 1) As Byte
    
    'The ROM address of the Chunk can of course be 0, so default to indefinite
    Let Me.Address = -1
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase Data
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Size _
 ======================================================================================
Public Property Get Size() As Long: Let Size = My_Size: End Property

'AddByte : Add a byte to the assembled output and move along _
 ======================================================================================
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

'AddWord : Add a 2-byte value, in little-endian order _
 ======================================================================================
Public Sub AddWord( _
    ByRef Value As Long _
)
    'Add the low-byte first _
     (the `AddByte` routine clips to 8-bit automatically)
    Call AddByte(Value)
    'And the high-byte
    Call AddByte((Value And &HFF00&) \ &H100)
End Sub

'AddOpcode : Inserts a single opcode into the byte stream _
 ======================================================================================
Public Function AddOpcode( _
    ByRef Opcode As Long _
) As String
    'Insert the provided opcode byte
    Call AddByte(Opcode)
    'Return a disassembly
    Let AddOpcode = "·· " & oz80.HexStr8(Opcode) & " ·· ·· ·· ··"
End Function

'AddOpcodePrefix : Inserts a single opcode into the byte stream with IX/IY prefix _
 ======================================================================================
Public Function AddOpcodePrefix( _
    ByRef Prefix As Long, _
    ByRef Opcode As Long _
) As String
    'Is an IX/IY prefix byte is provided?
    If Prefix <> 0 Then
        'Add the prefix to the byte stream
        Call AddByte(Prefix)
        'Return a disassembly
        Let AddOpcodePrefix = oz80.HexStr8(Prefix) & " " & oz80.HexStr8(Opcode) & " ·· ·· ·· ··"
    Else
        'Return a disassembly
        Let AddOpcodePrefix = "··" & oz80.HexStr8(Opcode) & " ·· ·· ·· ··"
    End If
    'Insert the provided opcode byte
    Call AddByte(Opcode)
End Function

'AddOpcode_Offset : Inserts an opcode with an optional offset value, i.e. "(ix+$FF)" _
 ======================================================================================
Public Function AddOpcode_Offset( _
    ByRef Prefix As Long, _
    ByRef Opcode As Long, _
    ByRef Offset As Long _
) As String
    'Is an IX/IY prefix byte is provided?
    If Prefix <> 0 Then
        'Add the prefix to the byte stream
        Call AddByte(Prefix)
        Let AddOpcode_Offset = oz80.HexStr8(Prefix) & " "
    Else
        Let AddOpcode_Offset = "·· "
    End If
    'Insert the provided opcode byte
    Call AddByte(Opcode)
    'If an offset exists, output it
    If Offset >= 0 Then
        'TODO: Validate 8-bit limit
        Call AddByte(Offset)
        'Return the disassembly (including the offset)
        Let AddOpcode_Offset = AddOpcode_Offset & _
            oz80.HexStr8(Opcode) & " " & _
            oz80.HexStr8(Offset) & " ·· ·· ··"
    Else
        'Return the disassembly (without the offset)
        Let AddOpcode_Offset = AddOpcode_Offset & _
            oz80.HexStr8(Opcode) & " ·· ·· ·· ··"
    End If
End Function

'AddOpcode_Value8Bit : Inserts opcode with 8-bit value, e.g. "ld a, $FF" _
 ======================================================================================
Public Function AddOpcode_Value8Bit( _
    ByRef Opcode As Long, _
    ByRef Value As Long _
) As String
    'Insert the provided opcode byte
    Call AddByte(Opcode)
    'TODO: Validate 8-bit limit
    Call AddByte(Value)
    'Return a disassembly
    Let AddOpcode_Value8Bit = _
        "·· " & oz80.HexStr8(Opcode) & " ·· ·· " & oz80.HexStr8(Value) & " ··"
End Function

'AddOpcode_Value8Bit_Offset : Inserts opcode with 8-bit value and optional offset _
 ======================================================================================
Public Function AddOpcode_Value8Bit_Offset( _
    ByRef Prefix As Long, _
    ByRef Opcode As Long, _
    ByRef Value As Long, _
    ByRef Offset As Long _
) As String
    'Is an IX/IY prefix byte is provided?
    If Prefix <> 0 Then
        'Add the prefix to the byte stream
        Call AddByte(Prefix)
        Let AddOpcode_Value8Bit_Offset = oz80.HexStr8(Prefix) & " "
    Else
        Let AddOpcode_Value8Bit_Offset = "·· "
    End If
    'Insert the provided opcode byte
    Call AddByte(Opcode)
    'If an offset value, e.g. "(ix+$FF)", is provided, include that
    If Offset >= 0 Then
        'TODO: Validate 8-bit limit
        Call AddByte(Offset)
        'Return a disassembly
        Let AddOpcode_Value8Bit_Offset = AddOpcode_Value8Bit_Offset & _
            oz80.HexStr8(Opcode) & " " & _
            oz80.HexStr8(Offset) & " ·· " & _
            oz80.HexStr8(Value) & " ··"
    Else
        Let AddOpcode_Value8Bit_Offset = AddOpcode_Value8Bit_Offset & _
            oz80.HexStr8(Opcode) & " ·· ·· " & _
            oz80.HexStr8(Value) & " ··"
    End If
    'Lastly the 8-bit value
    'TODO: Validate 8-bit limit
    Call AddByte(Value)
End Function

'AddOpcode_Value16Bit : Inserts opcode with 16-bit value, e.g. "jp $FFFF" _
 ======================================================================================
Public Function AddOpcode_Value16Bit( _
    ByRef Prefix As Long, _
    ByRef Opcode As Long, _
    ByRef Value As Long _
) As String
    'Is an IX/IY prefix byte is provided?
    If Prefix <> 0 Then
        'Add the prefix to the byte stream
        Call AddByte(Prefix)
        Let AddOpcode_Value16Bit = oz80.HexStr8(Prefix) & " "
    Else
        Let AddOpcode_Value16Bit = "·· "
    End If
    'Insert the provided opcode byte
    Call AddByte(Opcode)
    'TODO: Validate 16-bit limit
    Call AddWord(Value)
    'Return a disassembly
    Let AddOpcode_Value16Bit = AddOpcode_Value16Bit & _
        oz80.HexStr8(Opcode) & " ·· ·· " & _
        oz80.HexStr8(Value And &HFF) & " " & _
        oz80.HexStr8((Value And &HFF00&) \ &H100)
End Function

'AddOpcodeED : Inserts an ED-prefixed double opcode into the byte stream _
 ======================================================================================
Public Function AddOpcodeED( _
    ByRef Opcode As Long _
) As String
    'Add the ED prefix
    Call AddByte(&HED)
    'and the provided opcode
    Call AddByte(Opcode)
    'Return a disassembly
    Let AddOpcodeED = "ED " & oz80.HexStr8(Opcode) & " ·· ·· ·· ··"
End Function

'AddOpcode_Value16Bit : Inserts an ED-prefixed double opcode with 16-bit value _
 ======================================================================================
Public Function AddOpcodeED_Value16Bit( _
    ByRef Opcode As Long, _
    ByRef Value As Long _
) As String
    'Add the ED prefix
    Call AddByte(&HED)
    'and the provided opcode
    Call AddByte(Opcode)
    'TODO: Validate 16-bit limit
    Call AddWord(Value)
    'Return a disassembly
    Let AddOpcodeED_Value16Bit = _
        "ED " & _
        oz80.HexStr8(Opcode) & " ·· ·· " & _
        oz80.HexStr8(Value And &HFF) & " " & _
        oz80.HexStr8((Value And &HFF00&) \ &H100)
End Function

'AddOpcodeCB : Inserts a CB-prefixed double opcode with optional offset value _
 ======================================================================================
Public Function AddOpcodeCB( _
    ByRef Prefix As Long, _
    ByRef Opcode As Long, _
    ByRef Offset As Long _
) As String
    'Is an IX/IY prefix byte is provided?
    If Prefix <> 0 Then
        'Add the prefix to the byte stream
        Call AddByte(Prefix)
        Let AddOpcodeCB = oz80.HexStr8(Prefix) & " "
    Else
        Let AddOpcodeCB = "·· "
    End If
    Call AddByte(&HCB)
    'If an offset value, e.g. "(ix+$FF)", is provided, include that
    If Offset >= 0 Then
        'TODO: Validate 8-bit limit
        Call AddByte(Offset)
        'Return a disassembly
        Let AddOpcodeCB = AddOpcodeCB & _
            "CB " & _
            oz80.HexStr8(Offset) & " " & _
            oz80.HexStr8(Opcode) & " ·· ··"
    Else
        'Return a disassembly
        Let AddOpcodeCB = AddOpcodeCB & _
            "CB ·· " & _
            oz80.HexStr8(Opcode) & " ·· ··"
    End If
    'Finally, insert the opcode
    Call AddByte(Opcode)
End Function

'Trim _
 ======================================================================================
Public Sub Trim()
    'Chunks can be empty!
    Dim NewSize As Long
    If My_Size = 0 Then Let NewSize = 0 Else Let NewSize = My_Size - 1
    'Trim the buffer size down to the actual size
    ReDim Preserve Data(0 To NewSize) As Byte
    'Update the buffer size to match
    Let BufferSize = NewSize
End Sub

'CopyInto _
 ======================================================================================
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
