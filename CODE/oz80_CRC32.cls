VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_CRC32"
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
'CLASS :: oz80_CRC32

'The CRC function of this class is from "Calculating CRC32 With VB" by Steve McMahon _
 <www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp> which was _
 derived from from the VB.NET CRC32 algorithm by Paul Caton

'/// API //////////////////////////////////////////////////////////////////////////////

'<www.vbforums.com/showthread.php?729385-VB6-clsStrToIntArray-cls-Cast-String-To-Integer-Array>
Private Type SAFEARRAY1D    'Represents a safe array. (One Dimensional)
    cDims      As Integer   'The count of dimensions.
    fFeatures  As Integer   'Flags used by the SafeArray.
    cbElements As Long      'The size of an array element.
    cLocks     As Long      'The number of times the array has been locked without a corresponding unlock.
    pvData     As Long      'Pointer to the data.
    cElements  As Long      'The number of elements in the dimension.
    lLbound    As Long      'The lower bound of the dimension.
End Type                    'http://msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CompStrPtr  As Long
Private CompStrSA1D As SAFEARRAY1D

Dim CompStrArr() As Integer

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'This is the official polynomial used by CRC32 in PKZip. _
 Often the polynomial is shown reversed (04C11DB7)
Private Const CRC_POLYNOMIAL As Long = &HEDB88320
'Store a look up table for the CRC calculations
Private CRCTable(0 To 255) As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Dim i As Long, ii As Long
    
    'Create the CRC look-up table _
     (I will not pretend to understand how this works other than it goes through _
      shifting and flipping bits, make of that what you will)
    Dim Value As Long
    For i = 0 To 255
        Value = i
        For ii = 8 To 1 Step -1
            If (Value And 1) Then
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                Let Value = Value Xor CRC_POLYNOMIAL
            Else
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next
        Let CRCTable(i) = Value
    Next
    
    With CompStrSA1D
        .cDims = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .cbElements = 2&
        .cLocks = 1&
        .lLbound = 0&
    End With
    
    Let CompStrPtr = msvbvm60_VarPtr(CompStrArr())
    Call msvbvm60_PutMem4(CompStrPtr, VarPtr(CompStrSA1D))
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase CRCTable
    
    Call msvbvm60_PutMem4(CompStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'HashASCIIString : Perform a Cyclic-Redundancy-Check on a given string _
 ======================================================================================
'Based upon <vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp>
Public Function HashASCIIString(ByRef Data As String) As Long
    Dim Length As Long
    
    With CompStrSA1D
       Let .pvData = StrPtr(Data)
       Let .cElements = Len(Data)
       Let Length = .cElements
    End With

    Dim i As Long
    Dim iLookup As Integer

    Let HashASCIIString = &HFFFFFFFF
    Let Length = Length - 1
    For i = 0 To Length
        Let iLookup = (HashASCIIString And &HFF) Xor (CompStrArr(i) And &HFF)
        Let HashASCIIString = ((HashASCIIString And &HFFFFFF00) \ &H100) And 16777215
        Let HashASCIIString = HashASCIIString Xor CRCTable(iLookup)
    Next

   Let HashASCIIString = Not (HashASCIIString)
End Function
