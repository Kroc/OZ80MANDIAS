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

Private TokenStream As oz80TokenStream

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Assemble : Take a source file and produce a binary _
 ======================================================================================
Public Sub Assemble(ByVal FilePath As String)
    Debug.Print
    Debug.Print "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
    'Create a tokenStream object to hold the machine representation of the text files; _
     the assembler doesn't work with the original text directly
    Set TokenStream = New oz80TokenStream
    
    'Create a textParser that will explode the source file into separate words and _
     convert them into tokens which will be fed into the tokenStream we created
    Dim TextParser As oz80TextParser
    Set TextParser = New oz80TextParser
    Call TextParser.Parse(FilePath, TokenStream)
    Set TextParser = Nothing
    
    'Stage 2: Variable Sweep _
     ----------------------------------------------------------------------------------
    'Now we have the tokenStream, walk through it and populate information on the _
     variables defined. This is so that we will be able to refer to variables forward _
     of references that appear earlier in the source
End Sub

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up
    Set TokenStream = Nothing
End Sub
