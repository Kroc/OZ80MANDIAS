VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CRC32 : Copyright (C) Kroc Camen, 2014
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit

'The CRC function of this class is from "Calculating CRC32 With VB" by Steve McMahon _
 <www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp> which was _
 derived from from the VB.NET CRC32 algorithm by Paul Caton

'/// API //////////////////////////////////////////////////////////////////////////////

'Cast String to Array: _
 --------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB string directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for VB arrays _
 <msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx>
Private Type SAFEARRAY1D
    cDims      As Integer               'The count of dimensions
    fFeatures  As Integer               'Flags used by the SafeArray
    cbElements As Long                  'The size of an array element
    cLocks     As Long                  'Lock count
    pvData     As Long                  'Pointer to the data
    cElements  As Long                  'The number of elements in the dimension
    lLbound    As Long                  'The lower bound of the dimension
End Type

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

'Undocumented VB6 API to get at the array pointer
Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Byte            'Empty array that will be hacked

'--------------------------------------------------------------------------------------

'Store a look up table for the CRC calculations
Private CRCTable(0 To 255) As Long
'This is the official polynomial used by CRC32 in PKZip. _
 Often the polynomial is shown reversed (04C11DB7)
Private Const CRC_POLYNOMIAL As Long = &HEDB88320

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    With CastStrSA1D
        Let .cDims = 1
        Let .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        Let .cbElements = 2&
        Let .cLocks = 1&
        Let .lLbound = 0&
    End With
    
    Let CastStrPtr = msvbvm60_VarPtr(CastStrArr())
    Call msvbvm60_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
    
    'Create the CRC look-up table _
     (I will not pretend to understand how this works other than it goes through _
      shifting and flipping bits, make of that what you will)
    Dim i As Long, Value As Long
    For i = 0 To 255
        Value = i
        Dim ii As Long
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
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase CRCTable
    
    'Clean up our array hacking (for casting a string to an array) _
     otherwise VB will crash when it frees up the memory
    Call msvbvm60_PutMem4(CastStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'CastStringToArray _
 ======================================================================================
Public Function CastStringToArray(ByRef Source As String) As Byte()
    'Take the string, and attach it to our dummy array header
    With CastStrSA1D
        Let .pvData = StrPtr(Source)
        Let .cElements = LenB(Source)
    End With
    
    'Return a reference to the array
    Let CastStringToArray = CastStrArr
End Function

'Calculate _
 ======================================================================================
Public Function Calculate(ByRef Data As String) As Long
    Dim Length As Long
    
    'Cast the string to an array by attaching its pointer to our dummy VB array header
    With CastStrSA1D
        Let .pvData = StrPtr(Data)
        Let .cElements = LenB(Data)
        Let Length = .cElements
    End With

    'Calculate the CRC by twiddling the bits
    
    Dim i As Long
    Dim iLookup As Byte

    Let Calculate = &HFFFFFFFF
    Let Length = Length - 1
    For i = 0 To Length
        Let iLookup = (Calculate And &HFF) Xor CastStrArr(i)
        Let Calculate = ((Calculate And &HFFFFFF00) \ &H100) And 16777215
        Let Calculate = Calculate Xor CRCTable(iLookup)
    Next

    Let Calculate = Not (Calculate)
End Function
