VERSION 5.00
Begin VB.Form UI 
   Caption         =   "OZ80MANDIAS"
   ClientHeight    =   7005
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_GETLINECOUNT As Long = 186
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_USER As Long = &H400
Private Const EM_GETEVENTMASK As Long = (WM_USER + 59)
Private Const EM_SETEVENTMASK As Long = (WM_USER + 69)

Private Const WM_VSCROLL As Long = &H115

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Private WithEvents Assembler As oz80_Assembler
Attribute Assembler.VB_VarHelpID = -1

Private LogText As bluArrayStrings

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Form_Initialize()
    Set LogText = New bluArrayStrings
    Let LogText.AllowDuplicates = True
End Sub

Private Sub Form_Terminate()
    Set LogText = Nothing
End Sub

'FORM Load
'======================================================================================
Private Sub Form_Load()
    Call Me.Show
    
    Dim StartTime As Single
    Let StartTime = Timer

'    Dim i As Long
    Dim Test As New bluString
    Let Test.Text = "ABCD�F"
    Debug.Print Test.Left(10, ASTERISK).Text
    
'    Debug.Print Test.Join(Test.Clone.LCase).Text
'    Debug.Print Test.Append("!").Text
'    Debug.Print Test.Insert("_1_2_3_", 2).Text
'    Debug.Print Test.Prepend("�").Text
'    Debug.Print Test.Insert("G", -1).Text
'    Call Test.CharPush(65535)
'    Debug.Print Test.CharPull()
'    Debug.Print Test.CharRemove(2), Test.Text
'    Debug.Print Test.CharInsert(2, 66).Text
'    Debug.Print Test.Remove(14, 6).Text
'    Debug.Print Test.Replace("_1_2_3_", "{$}").Text
'    Debug.Print Test.Wrap("""").Text
'    Debug.Print Test.Format(3.141).Text
'    Debug.Print Test.Overwrite(2, "***").Text
'    Debug.Print Hex(Test.CRC())
    
'    For i = 0 To 999999
'        Call Test.Replace("_1_2_3_", "_!_").Replace("_!_", "_1_2_3_")
'        Call Test.Equals("�AB_1_2_3_CD�FG!!")
'    Next i
'    MsgBox Format$(Timer - StartTime, "0.000")
'    End
    
    Set Assembler = New oz80_Assembler

    'TODO: This will obviously be converted to use the command arguments
    Call Assembler.Assemble(App.Path & "\Sonic1-sms-oz80\sonic1-sms.oz80")
'    Call Assembler.Assemble(App.Path & "\TEST\test.oz80")
    
    'Do something that only faults in the IDE
    On Error GoTo Err_True
    Debug.Print 1 \ 0
    MsgBox Format$(Timer - StartTime, "0.000")
Err_True:

    Set Assembler = Nothing
    
    Call SendMessage( _
        Me.txtLog.hWnd, EM_SETSEL, _
        Len(Me.txtLog.Text), Len(Me.txtLog.Text) _
    )
    
    Call SendMessageString( _
        Me.txtLog.hWnd, EM_REPLACESEL, _
        ByVal 0, LogText.Concatenate() _
    )
    
    Call SendMessage( _
        Me.txtLog.hWnd, WM_VSCROLL, 7, ByVal 0 _
    )
End Sub

'FORM Resize _
 ======================================================================================
Private Sub Form_Resize()
    Call Me.txtLog.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

'EVENT <Assembler> Error _
 ======================================================================================
Private Sub Assembler_Error( _
    ByRef FilePath As String, _
    ByVal Number As OZ80_ERROR, _
    ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)
    Call Log
    Call Log("! ERROR: #" & Number & " " & Title, OZ80_LOG_ACTION)
    Call Log("- File: """ & FilePath & """", OZ80_LOG_INFO)
    If Line > 0 And Col > 0 Then
        Call Log("- Line: " & Format$(Line, "#,#") & " Col: " & Col, OZ80_LOG_INFO)
    End If
    Call Log("- " & Description, OZ80_LOG_INFO)
End Sub

'EVENT <Assembler>_Message
'======================================================================================
Private Sub Assembler_Message( _
    ByRef LogLevel As OZ80_LOG, ByRef LogText As bluString _
)
    Static PrevLog As OZ80_LOG
    
    Dim Prefix As String
    If LogLevel = OZ80_LOG_ACTION Then Let Prefix = Prefix & "*"
    If LogLevel = OZ80_LOG_INFO Then Let Prefix = Prefix & "-"
    If LogLevel = OZ80_LOG_STATUS Then Let Prefix = Prefix & "="
    If LogLevel = OZ80_LOG_DEBUG Then Let Prefix = Prefix & "."
    
    If (LogLevel = OZ80_LOG_ACTION) And (PrevLog <> OZ80_LOG_ACTION) Then
        Let Prefix = vbCrLf & Prefix
    End If
    
    Call Log(Prefix & " " & Replace(LogText.Text, vbCrLf, vbCrLf & "  "), LogLevel)
    Let PrevLog = LogLevel
End Sub

'Log : Add a message to the log
'======================================================================================
Private Sub Log( _
    Optional ByRef Text As String = vbNullString, _
    Optional ByRef LogLevel As OZ80_LOG = OZ80_LOG_ACTION _
)
'    Debug.Print Text
    
'    If LogLevel >= OZ80_LOG_DEBUG Then Exit Sub

    Call LogText.Add( _
        Text & vbCrLf _
    )
    
'    'http://weblogs.asp.net/jdanforth/88458
'    Call SendMessage( _
'        Me.txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0 _
'    )
'    Dim EventMask As Long
'    Let EventMask = SendMessage( _
'        Me.txtLog.hWnd, EM_GETEVENTMASK, 0, ByVal 0 _
'    )
    
    'Thanks to Jdo300 for this execllent tip to prevent flicker _
     <xtremevbtalk.com/showpost.php?p=1330080&postcount=2>
    'Overcome the 64K text limit in VB6: _
     <www.tek-tips.com/viewthread.cfm?qid=1469439>
    
'    Call LockWindowUpdate(Me.txtLog.hWnd)
    
'    Call SendMessage( _
'        Me.txtLog.hWnd, EM_SETSEL, _
'        Len(Me.txtLog.Text), Len(Me.txtLog.Text) _
'    )
'
'    Call SendMessageString( _
'        Me.txtLog.hWnd, EM_REPLACESEL, _
'        ByVal 0, Text _
'    )
'
'    Call SendMessage( _
'        Me.txtLog.hWnd, WM_VSCROLL, 7, ByVal 0 _
'    )
    
'    Call LockWindowUpdate(0)
    
'    Call SendMessage( _
'        Me.txtLog.hWnd, EM_SETEVENTMASK, 0, ByVal EventMask _
'    )
'    Call SendMessage( _
'        Me.txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0 _
'    )
'    DoEvents
End Sub
