VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OZ80Blob"
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
'CLASS :: OZ80Blob

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private Data() As Byte

Private My_Buffer As Long
Private My_Length As Long
Private My_Cursor As Long

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Erase Data
End Sub

Public Property Get Length() As Long
    Let Length = My_Length
End Property

Public Property Let Length( _
    ByVal NewLength As Long _
)
    Let My_Length = NewLength
    'Does the buffer need increasing?
    If My_Length > My_Buffer Then
        'If the buffer is not allocated, initialise its size
        If My_Buffer = 0 Then Let My_Buffer = 256
        'Reallocate the buffer
        Let My_Buffer = My_Buffer * 2
        ReDim Preserve Data(1 To My_Buffer) As Byte
    End If
End Property

Public Function AppendBYTE( _
    ByRef Value As Byte _
)
    'Increase the length by a byte
    Let Me.Length = My_Length + 1
    'Write the new byte
    Let bluW32.MemByte(VarPtr(Data(My_Length))) = Value
End Function

Public Function AppendWORD( _
    ByRef Value As Long _
)
    'Increase the length by two bytes
    Let Me.Length = My_Length + 2
    'Write the new word (unsigned Integer)
    Let bluW32.MemInteger(VarPtr(Data(My_Length))) = bluW32.MemInteger(VarPtr(Value))
End Function

Public Function AppendTRIP( _
    ByRef Value As Double _
)
    '
End Function

Public Function AppendLONG( _
    ByRef Value As Double _
)
    '
End Function

Public Function AppendREAL( _
    ByRef Value As Double _
)
    '
End Function

Public Function Clear()
    Let My_Length = 0
End Function
