Attribute VB_Name = "Run"
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: Run

Private Sub Main()
    Dim CommandArgs() As String
    Let CommandArgs = blu.CommandParams()
    
    If LBound(CommandArgs) = 0 Then
        MsgBox "OZ80MANDIAS usage:" & vbNewLine & vbNewLine _
             & "oz80mandias.exe <filePath>"
        Exit Sub
    End If
    
    Load UI
    Call UI.Show
    Call UI.Assemble(CommandArgs(1))
End Sub
