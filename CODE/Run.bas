Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'MAIN : "Look on my works ye Mighty, and despair!" _
 ======================================================================================
Public Sub Main()
    Dim Assembler As oz80Assembler
    Set Assembler = New oz80Assembler
    
    Dim StartTime As Single
    Let StartTime = Timer
    
'    Call Assembler.Assemble(App.Path & "\S1.sms.asm")
    Call Assembler.Assemble(App.Path & "\sonic1-sms.oz80")
    
    If InIDE = False Then MsgBox Timer - StartTime
    
    Set Assembler = Nothing
End Sub

'PROPERTY InIDE : Are we running the code from the Visual Basic IDE? _
 ======================================================================================
Public Property Get InIDE() As Boolean
    On Error GoTo Err_True
    
    'Do something that only faults in the IDE
    Debug.Print 1 \ 0
    InIDE = False
    Exit Property

Err_True:
    InIDE = True
End Property

