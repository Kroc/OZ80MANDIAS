VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_Binary"
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
'CLASS :: oz80_Binary

'Increasing an array byte by byte is *extremely* slow (if the array has no room to _
 grow in the RAM, the whole thing has to be moved up or other garbage collection ran). _
 To minimise the number of times the array has to be expanded, we start with a buffer _
 size and double it each time it's filled. Once you've completed adding your data, _
 call `trim` to cut the array down to its final size
 
'This is the default buffer size to use, and the 'current' size. _
 Note that this is 1-based, so that it doubles correctly
Private Const BUFFER_DEFAULT As Long = 1024
Private My_Buffer As Long

'This is the actual binary array
Private My_Data() As Byte

'Current position within the array
Private My_Index As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Let Me.Buffer = BUFFER_DEFAULT
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase My_Data
End Sub

'/// PROPERTIES ///////////////////////////////////////////////////////////////////////

Public Property Get Buffer() As Long
    Let Buffer = My_Buffer
End Property

Public Property Let Buffer(ByVal Length As Long)
    'Validate
    If Length < 0 Then Err.Raise (9)
    
    Let My_Buffer = Length
    ReDim Preserve My_Data(0 To My_Buffer - 1) As Byte
    If My_Index > My_Buffer Then
        '
    End If
End Property

Public Property Get Index() As Long
    Let Index = My_Index
End Property

Public Property Let Index(ByRef Position As Long)
    'Validate
    If Position < 0 Then Err.Raise (9)
    If Position >= My_Buffer Then Err.Raise (9)
    
    Let My_Index = Position
End Property

'/// PROCEDURES ///////////////////////////////////////////////////////////////////////

'AddByte : Adds one byte to the array and moves the index forward _
 ======================================================================================
Public Sub AddByte( _
    ByRef Datum As Long _
)   'Long is used for speed only -- the value is truncated to 255
    Let My_Data(My_Index) = Datum And &HFF
    Let My_Index = My_Index + 1
    
    'Has the buffer overflowed?
    If My_Index >= My_Buffer Then
        'Double the buffer size
        Let My_Buffer = My_Buffer + My_Buffer
        'Increase the byte array
        ReDim Preserve My_Data(0 To My_Buffer - 1) As Byte
    End If
End Sub

'AddWord : Adds two bytes to the array, in little-endian (reversed) order _
 ======================================================================================
Public Sub AddWord( _
    ByRef Datum As Long _
)
    'Add the low-byte first
    Call AddByte(Datum And &HFF)
    'And the high-byte
    Call AddByte((Datum And &HFF00&) \ &H100)
End Sub

'Trim : Reduces the binary from the bufer size to the current index position _
 ======================================================================================
Private Sub Trim()
    '
End Sub
