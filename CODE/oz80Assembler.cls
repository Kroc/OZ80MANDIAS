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

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Assemble : Take a source file and produce a binary _
 ======================================================================================
Public Sub Assemble(ByVal FilePath As String)
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
    Call Tokeniser.Tokenise(FilePath)
    
    'Stage 2: Validate & Optimise Token Stream _
     ----------------------------------------------------------------------------------
    'The token stream is only a word-for-word machine representation of the source _
     code, we need to check the grammer for validity and whilst there we will do some _
     optimisations on the stream (e.g. collapsing lists into byte streams)
     
     
    'Clean up
    Set Tokeniser = Nothing
End Sub

