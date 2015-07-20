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
      ScrollBars      =   3  'Both
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

Private WithEvents Assembler As OZ80
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
    '
End Sub

'FORM Resize
'======================================================================================
Private Sub Form_Resize()
    Call Me.txtLog.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

'EVENT <Assembler> Error
'======================================================================================
Private Sub Assembler_Error( _
    ByRef FilePath As String, _
    ByVal Number As OZ80_ERROR, _
    ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)
    Dim Message As String
    Let Message = "ERROR #" & Number & " " & Title & vbNewLine
    Let Message = Message & "File: """ & FilePath & """" & vbNewLine
    If Line > 0 And Col > 0 Then
        Let Message = Message & "Line: " & Format$(Line, "#,#") & " Col: " & Col & vbNewLine
    End If
    Let Message = Message & vbNewLine & Description & vbNewLine
    
    'Display the particular source code line
    '------------------------------------------------------------------------------
    If FilePath <> vbNullString Then
        Dim Lines() As String
        
        Dim FileNumber As Integer
        Let FileNumber = FreeFile
        Open FilePath For Input Access Read Lock Write As #FileNumber
    
        Let Lines = Split( _
            StrConv(InputB(LOF(FileNumber), FileNumber), vbUnicode), _
            vbNewLine _
        )
        
        Let Message = Message _
            & vbNewLine _
            & "¶     ¶" & vbNewLine
        If Line > 2 Then Let Message = Message _
            & "|" & Right$("     " & Str$(Line - 2), 5) & "| " & Lines(Line - 3) & vbNewLine
        If Line > 1 Then Let Message = Message _
            & "|" & Right$("     " & Str$(Line - 1), 5) & "| " & Lines(Line - 2) & vbNewLine
        Let Message = Message _
            & "|" & Right$("     " & Str$(Line), 5) & "| " & Lines(Line - 1) & vbNewLine _
            & "`-----'-" & String$(Col - 1, "-") & "^" & vbNewLine
        
        Close #FileNumber
    End If
    
    Debug.Print Message
    Call Log(Message, OZ80_LOG_ERROR)
End Sub

'EVENT <Assembler>_Message
'======================================================================================
Private Sub Assembler_Message( _
    ByRef LogLevel As OZ80_LOG, ByRef LogText As bluString _
)
    Call Log(LogText.Text, LogLevel)
End Sub

Public Function Assemble( _
    ByRef FilePath As String _
) As OZ80_ERROR
    Dim StartTime As Single
    Let StartTime = Timer

    Dim i As Long
    Dim Test As New bluString
'    Let Test.Text = "ABCD…F"
'    Debug.Print Test.Insert("The Quick Brown Fox", 0).Truncate(12).Text
'    End
    
'    Debug.Print Test.Left(10, ASTERISK).Text
    
'    Debug.Print Test.Join(Test.Clone.LCase).Text
'    Debug.Print Test.Append("!").Text
'    Debug.Print Test.Insert("_1_2_3_", 2).Text
'    Debug.Print Test.Prepend("È").Text
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
'        Call Test.Equals("ÈAB_1_2_3_CD…FG!!")
'    Next i
'    MsgBox Format$(Timer - StartTime, "0.000")
'    End
    
    Set Assembler = New OZ80

    'TODO: This will obviously be converted to use the command arguments
    Call Assembler.Assemble(FilePath)
    
    Call Log("External Time: " & Format$(Timer - StartTime, "0.000"), OZ80_LOG_INFO)
    Call Log(Assembler.Profiler.Report, OZ80_LOG_INFO)
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
End Function

'Log : Add a message to the log
'======================================================================================
Private Sub Log( _
    Optional ByRef Text As String = vbNullString, _
    Optional ByRef LogLevel As OZ80_LOG = OZ80_LOG_ACTION _
)
    Static PrevLog As OZ80_LOG
   
    Dim Prefix As String
    If LogLevel = OZ80_LOG_ERROR Then
        Let Prefix = "!"
    ElseIf LogLevel = OZ80_LOG_ACTION Then Let Prefix = "*"
    ElseIf LogLevel = OZ80_LOG_INFO Then Let Prefix = "-"
    ElseIf LogLevel = OZ80_LOG_STATUS Then Let Prefix = "="
    ElseIf LogLevel = OZ80_LOG_DEBUG Then Let Prefix = "."
    End If
    
    If (LogLevel <= OZ80_LOG_ACTION) And (PrevLog > OZ80_LOG_ACTION) Then
        Let Prefix = vbNewLine & Prefix
    End If
    
    Dim Msg As String
    Let Msg = Prefix & " " & Replace(Text, vbNewLine, vbNewLine & "  ") & vbNewLine
    
'    Debug.Print Msg
    Call LogText.Add(Msg)
    
    Let PrevLog = LogLevel
 
'    '----------------------------------------------------------------------------------
'
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
'        ByVal 0, Text & vbCrLf _
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
    DoEvents
End Sub
