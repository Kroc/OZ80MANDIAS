VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OZ80Data"
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
'CLASS :: OZ80Data

'This is OZ80's user data structure. A means of storing a variety of byte-packed data _
 suitable for use with Z80 code or in macro processing

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'This will be the byte-array where the data is packed. This class will allow you to _
 pluck Bytes, Integers, Longs &c. out from any point in the array
Private Data() As Byte

'Size of our structure in bytes, cannot be more than 2 GB
Private My_Size As Long

'/// API //////////////////////////////////////////////////////////////////////////////

'Undocumented VB6 APIs to write a number of bytes to memory

Private Declare Sub api_PutMem1 Lib "msvbvm60" Alias "PutMem1" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Byte _
)
Private Declare Sub api_PutMem2 Lib "msvbvm60" Alias "PutMem2" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Integer _
)
Private Declare Sub api_PutMem4 Lib "msvbvm60" Alias "PutMem4" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Long _
)

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Erase Data
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Public Sub AddByte( _
    ByRef Value As Byte _
)
    Let My_Size = My_Size + 1
    ReDim Preserve Data(1 To My_Size) As Byte
    
    Let Data(My_Size) = Value
End Sub

Public Property Get Pointer( _
) As Long
    Let Pointer = VarPtr(Data(0))
End Property

'Size
'======================================================================================
Public Property Get Size( _
) As Long
    Let Size = My_Size
End Property
