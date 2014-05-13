VERSION 5.00
Begin VB.Form UI 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "OZ80MANDIAS"
   ClientHeight    =   5070
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BorderStyle     =   0  'None
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
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
    Call Log
    Call Log("! ERROR: #" & Number)
    If Line > 0 And Col > 0 Then
        Call Log("- Line: " & Format$(Line, "#,#") & " Col: " & Col)
    End If
    Call Log("- " & Description)
    Call Log
End Sub

'EVENT <Assembler> Message _
 ======================================================================================
Private Sub Assembler_Message(Text As String)
    Call Log(Text)
End Sub

Private Sub Log(Optional ByRef Text As String = vbNullString)
    'Thanks to Jdo300 for this execllent tip to precent flicker _
     <xtremevbtalk.com/showpost.php?s=8d2b9228eb01630447300b4488275f11&p=1330080&postcount=2>
    Let Me.txtLog.SelStart = Len(Me.txtLog.Text)
    Let Me.txtLog.SelText = Text & vbCrLf
    
    DoEvents
End Sub
