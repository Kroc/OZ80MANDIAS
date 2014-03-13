Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'This is not part of the assembler itself, this is just a bootstrap to test it

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'MAIN : "Look on my works ye Mighty, and despair!" _
 ======================================================================================
Public Sub Main()
    Dim StartTime As Single
    Let StartTime = Timer
    
    'TODO: This will obviously be converted to use the command arguments
    Call oz80.Assemble(App.Path & "\sonic1-sms.oz80")
    
    'Do something that only faults in the IDE
    On Error GoTo Err_True
    Debug.Print 1 \ 0
    MsgBox Format$(Timer - StartTime, "0.000")
Err_True:
End Sub

