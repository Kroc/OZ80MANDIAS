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
    Log "OZ80MANDIAS"
    Call OZ80_Parser.Parse(App.Path & "\test.OZ8.asm")
    Debug.Print
End Sub

'Log _
 ======================================================================================
Public Sub Log(ByVal Msg As String)
    Debug.Print Msg
End Sub

