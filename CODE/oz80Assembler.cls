VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80Assembler"
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
'CLASS :: oz80Assembler

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE = 0                 'Assembly completed successfully
    OZ80_ERROR_UNEXPECTED = 1           'I made a programming mistake
    OZ80_ERROR_FILENOTFOUND = 2         'Requested file does not exist
    OZ80_ERROR_FILEREAD = 3             'Some kind of problem with file handle open
    OZ80_ERROR_INVALIDNAME = 4          'Invalid label/property/variable name
    OZ80_ERROR_BADWORD = 5              'Couldn't parse a word
End Enum

Public InIDE As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    On Error GoTo Err_True
    
    'Do something that only faults in the IDE
    Debug.Print 1 \ 0
    Let Me.InIDE = False
    Exit Sub

Err_True:
    Let Me.InIDE = True
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Assemble : Take a source file and produce a binary _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    Dim Tokeniser As oz80Tokeniser
    
    Debug.Print
    Debug.Print "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision
    
    Dim StartTime As Single
    Let StartTime = Timer
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
    'Create a tokeniser object to hold the machine representation of the text files; _
     the assembler doesn't work with the original text directly
    Set Tokeniser = New oz80Tokeniser
    'Explode the source code file into tokens
    If Tokeniser.Tokenise(FilePath) <> OZ80_ERROR_NONE Then GoTo Finish
    
    'Stage 2: Validate & Optimise Token Stream _
     ----------------------------------------------------------------------------------
    'The token stream is only a word-for-word machine representation of the source _
     code, we need to check the grammer for validity and whilst there we will do some _
     optimisations on the stream (e.g. collapsing lists into byte streams)
    
     
Finish:
    'For any error that occured, it will be assumed that the relevant function has _
     already printed the error since it will have the right information to hand
    Set Tokeniser = Nothing
End Function
