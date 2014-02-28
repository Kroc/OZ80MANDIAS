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
    
    'TODO: This will obviously be converted to use the command arguments
    Call Assembler.Assemble(App.Path & "\sonic1-sms.oz80")
    
    If Assembler.InIDE = False Then MsgBox Format$(Timer - StartTime, "0.000")
    
    Set Assembler = Nothing
End Sub

