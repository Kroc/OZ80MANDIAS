VERSION 5.00
Begin VB.Form UI 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "OZ80MANDIAS"
   ClientHeight    =   5076
   ClientLeft      =   108
   ClientTop       =   372
   ClientWidth     =   6876
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5076
   ScaleWidth      =   6876
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   0
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3372
   End
End
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Assembler As oz80_Assembler
Attribute Assembler.VB_VarHelpID = -1

'FORM Load _
 ======================================================================================
Private Sub Form_Load()
    Call Me.Show
    
    Dim StartTime As Single
    Let StartTime = Timer
    
    Set Assembler = New oz80_Assembler
        
    'TODO: This will obviously be converted to use the command arguments
    Call Assembler.Assemble(App.Path & "\Sonic1-sms-oz80\Sonic1-sms.oz80")
    
    'Do something that only faults in the IDE
    On Error GoTo Err_True
    Debug.Print 1 \ 0
    MsgBox Format$(Timer - StartTime, "0.000")
Err_True:

    Set Assembler = Nothing
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    Call Me.txtLog.Move( _
        0, 0, Me.ScaleWidth, Me.ScaleHeight _
    )
End Sub

'EVENT <Assembler> Error _
 ======================================================================================
Private Sub Assembler_Error( _
    ByVal Number As OZ80_ERROR, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)
    Debug.Print
    Debug.Print "! ERROR: #" & Number
    If Line > 0 And Col > 0 Then
        Debug.Print "- Line: " & Format$(Line, "#,#") & " Col: " & Col
    End If
    Debug.Print "- " & Description
    Debug.Print
End Sub

'EVENT <Assembler> Message _
 ======================================================================================
Private Sub Assembler_Message(Text As String)
    Debug.Print Text
End Sub
