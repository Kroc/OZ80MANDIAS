Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'MAIN : "Look on my works ye Mighty, and despair!" _
 ======================================================================================
Public Sub Main()
    Call OZ8.Assemble(App.Path & "\test.OZ8.asm")
End Sub
