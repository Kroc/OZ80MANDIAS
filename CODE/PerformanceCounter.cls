VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerformanceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This class was adapted from one included in "Fastest CRC32 Out There", _
 <Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74450&lngWId=1>

Private seqFreq As Currency
Private seqStart As Currency
Private fAvailable As Boolean

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    If QueryPerformanceFrequency(seqFreq) = 0& Then
        MsgBox "PerfCounter not available", vbExclamation
    Else
        fAvailable = True
    End If
End Sub

Public Sub Reset()
    Call QueryPerformanceCounter(seqStart)
End Sub

Public Property Get Elapsed() As Double
    Dim seqNow As Currency
    
    Call QueryPerformanceCounter(seqNow)
    If fAvailable Then
        Let Elapsed = 1000 * (seqNow - seqStart) / seqFreq
    End If
End Property


