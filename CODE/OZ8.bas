Attribute VB_Name = "OZ8"
Option Explicit
Option Compare Text
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: OZ8

'The assembler core code

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

Private Const OZ8_SYNTAX_COMMENT = ";"
Private Const OZ8_SYNTAX_LABEL = ":"
Private Const OZ8_SYNTAX_VARIABLE = "!"
Private Const OZ8_SYNTAX_MACRO = "@"
Private Const OZ8_SYNTAX_OBJECT = "#"

Private Const OZ8_SYNTAX_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVQXYZ_"
Private Const OZ8_SYNTAX_NUMERIC = "0123456789-"

Dim Keywords As New Dictionary

Private Enum OZ80_CONTEXT
    UNKNOWN = 0
    'Inside a comment
    COMMENT = 1
    
    KEYWORD = 2
    LABEL = 3
    MACRO = 4
    'When inside a string, i.e. `... "some text" ...`
    QUOTED = 100
End Enum
#If False Then
    Private UNKNOWN, COMMENT, KEYWORD, LABEL, MACRO, QUOTED
#End If

'Contexts can be nested (e.g. a label within calculation within a data statement), _
 so a stack is managed to handle the recursive nature
Private ContextStack(0 To 255) As OZ80_CONTEXT
Private ContextPointer As Long

Public Enum OZ8_ERROR
    NOERROR = 0
    'An invalid label name
    INVALID_LABEL = 1
End Enum

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Assemble : Parse and assemble a file into a binary _
 ======================================================================================
Public Function Assemble( _
    ByVal FilePath As String, _
    Optional ByVal OutputPath As String = vbNullString _
) As OZ8_ERROR
    Log "OZ80MANDIAS"
    
    'Populate the table of keywords that begin new statements
    Call Keywords.RemoveAll
    Dim Key As Variant
    For Each Key In Array( _
        "SET", "BANK", "AT", "DATA", "FILL", "INCLUDE", "IF", "BEGIN", "EXIT" _
    )
        Call Keywords.Add(CStr(Key), CStr(Key))
    Next Key
    
    Erase ContextStack
    Let ContextPointer = 0
    
    Log ProcessFile(FilePath)
    
    Call Keywords.RemoveAll
    Debug.Print
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ProcessFile : Since files can be nested (with `INCLUDE`) make this process recursive _
 ======================================================================================
Private Function ProcessFile(ByVal FilePath As String) As OZ8_ERROR
    Dim FileNumber As Integer: Let FileNumber = FreeFile()
    Open FilePath For Binary Access Read Lock Write As #FileNumber
    
    Log "Reading file: " & FilePath
    Log "Length: " & LOF(FileNumber) & " bytes"
    
    'As we process characters we wait for a whole word or statement to build up
    Dim Word As String
    
    Do
        'Check the current context to see what we should be doing with this information
        Select Case Context
            Case OZ80_CONTEXT.UNKNOWN: '-----------------------------------------------
                'Lines can begin with whitespace, labels & keywords; _
                 parenthesis, variables and numbers are not allowed
                GoSub ReadWord
            
            Case OZ80_CONTEXT.COMMENT: '-----------------------------------------------
                'Comments can contain any characters, but finish at the end of the line
                'We don't need to process comments any further once read, just leave _
                 the comment context and continue
                Call PopContext
            
            Case OZ80_CONTEXT.LABEL: '-------------------------------------------------
                'When a label is defined, code can follow on the same line
                Call PopContext
                
        End Select
        
        If Context <> OZ80_CONTEXT.UNKNOWN Then Log Word
    Loop While EOF(FileNumber) = False
    GoTo Finish
    
    '==================================================================================
ReadWord:
    Let Word = vbNullString
ReadChar:
    Dim Char As String * 1
    'If the file ends, treat it as a remaining end of line
    If EOF(FileNumber) = True Then Let Char = vbCr Else Get #FileNumber, , Char
    'If this is just a blank line, skip
    ''If Word = vbNullString And IsEndOfLine(Char) = True Then Return
    If IsEndOfLine(Char) = True Then Return
    
ProcessWord:
    Select Case Context
        Case OZ80_CONTEXT.UNKNOWN: '---------------------------------------------------
            'Lines can begin with whitespace, labels & keywords; _
             parenthesis, variables and numbers are not allowed
            
            'Ignore whitespace at the start of lines
            If IsWhitespace(Char) = True Then GoTo ReadChar
            
            Select Case Char
                Case OZ8_SYNTAX_COMMENT
                    Call PushContext(COMMENT)
                Case OZ8_SYNTAX_LABEL
                    Call PushContext(LABEL)
                Case OZ8_SYNTAX_MACRO
                    Call PushContext(MACRO)
                Case Else
                    'If a line does not begin with a comment, label or macro name _
                     then it may only begin with a keyword
                    Call PushContext(KEYWORD)
            End Select
        
        Case OZ80_CONTEXT.COMMENT: '---------------------------------------------------
            'Comments can contain any characters, but finish at the end of the line
            'this subroutine already handles automatic end of line termination, _
             so we don't need to do anything explicit here
        
        Case OZ80_CONTEXT.LABEL: '-----------------------------------------------------
            'Label names must begin with ":", contain A-Z, 0-9, underscore and dash _
             only with the restriction that the first letter must be A-Z or an _
             underscore and not a numeral
            
            'Is this the first letter of the label name?
            If Len(Word) = 1 Then
                'For the first letter of the label, accept only alpha
                If InStr(OZ8_SYNTAX_ALPHA, Char) = 0 Then
                    Let ProcessFile = INVALID_LABEL
                    GoTo Finish
                End If
            Else
                'After at least one letter, any whitespace ends the label
                If IsWhitespace(Char) = True Then Return
                'For the second letter onwards allow alphanumerics
                If InStr(OZ8_SYNTAX_ALPHA & OZ8_SYNTAX_NUMERIC, Char) = 0 Then
                    Let ProcessFile = INVALID_LABEL
                    GoTo Finish
                End If
            End If
            
    End Select
    
    If IsEndOfLine(Char) = True Then Return
    Let Word = Word & Char
    GoTo ReadChar
    
Finish:
    Close #FileNumber
End Function

'PushContext : Add a new context to the stack _
 ======================================================================================
Private Sub PushContext(ByVal Context As OZ80_CONTEXT)
    Let ContextPointer = ContextPointer + 1
    Let ContextStack(ContextPointer) = Context
End Sub

'PopContext : End a context and return to the previous one _
 ======================================================================================
Private Function PopContext() As OZ80_CONTEXT
    Let ContextPointer = ContextPointer - 1
    Let PopContext = ContextStack(ContextPointer)
End Function

'IsWhitespace : check for meaningless whitespace (space, tab) _
 ======================================================================================
Private Function IsWhitespace(ByVal Char As String) As Boolean
    'For speed we won't use an OR statement as both comparisons are executed
    If Char = " " Then Let IsWhitespace = True: Exit Function
    If Char = vbTab Then Let IsWhitespace = True: Exit Function
End Function

'IsEndOfLine _
 ======================================================================================
Private Function IsEndOfLine(ByVal Char As String) As Boolean
    'For speed we won't use an OR statement as both comparisons are executed
    If Char = vbCr Then Let IsEndOfLine = True: Exit Function
    If Char = vbLf Then Let IsEndOfLine = True: Exit Function
    If Asc(Char) = 0 Then Let IsEndOfLine = True: Exit Function
End Function

Private Function Log(ByVal Msg As String, Optional ByVal Depth As Long = -1)
        If Depth = -1 Then Let Depth = ContextPointer
        If Depth > 0 Then Debug.Print String(Depth - 1, vbTab);
        Debug.Print Msg
End Function

'/// PRIVATE PROPERTIES ///////////////////////////////////////////////////////////////

'PROPERTY Context : Current syntax context (i.e. if we are in a macro, data &.c) _
 ======================================================================================
Private Property Get Context() As OZ80_CONTEXT
    Let Context = ContextStack(ContextPointer)
End Property
