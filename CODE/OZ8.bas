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

Private Const OZ80_SYNTAX_COMMENT = ";"
Private Const OZ80_SYNTAX_LABEL = ":"
Private Const OZ80_SYNTAX_VARIABLE = "!"
Private Const OZ80_SYNTAX_MACRO = "@"
Private Const OZ80_SYNTAX_OBJECT = "#"

Private Const OZ80_SYNTAX_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVQXYZ_"
Private Const OZ80_SYNTAX_NUMERIC = "0123456789-"

Private Const OZ80_KEYWORD_SET = "SET"
Private Const OZ80_KEYWORDS = "|" & OZ80_KEYWORD_SET & "|"

'--------------------------------------------------------------------------------------

Private Enum OZ80_CONTEXT
    UNKNOWN = 0
    'Inside a comment
    COMMENT = 1
    
    KEYWORD = 2
    LABEL = 3
    VARIABLE = 4
    MACRO = 5
    'When inside a string, i.e. `... "some text" ...`
    
    QUOTED = 100
    
    KEYWORD_SET = 1000
End Enum
#If False Then
    Private UNKNOWN, COMMENT, KEYWORD, LABEL, MACRO, QUOTED, _
            KEYWORD_SET
#End If

'Contexts can be nested (e.g. a label within calculation within a data statement), _
 so a stack is managed to handle the recursive nature
Private ContextStack(0 To 255) As OZ80_CONTEXT
Private ContextPointer As Long

'--------------------------------------------------------------------------------------

Public Enum OZ80_ERROR
    NOERROR = 0
    INVALID_LABEL = 1
    INVALID_VARIABLE = 2
    UNKNOWN_KEYWORD = 3
End Enum

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Assemble : Parse and assemble a file into a binary _
 ======================================================================================
Public Function Assemble( _
    ByVal FilePath As String, _
    Optional ByVal OutputPath As String = vbNullString _
) As OZ80_ERROR
    Log "OZ80MANDIAS"
    
    Erase ContextStack
    Let ContextPointer = 0
    
    Log ProcessFile(FilePath)
    
    Debug.Print
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ProcessFile : Since files can be nested (with `INCLUDE`) make this process recursive _
 ======================================================================================
Private Function ProcessFile(ByVal FilePath As String) As OZ80_ERROR
    Dim FileNumber As Integer: Let FileNumber = FreeFile()
    Open FilePath For Binary Access Read Lock Write As #FileNumber
    
    Log "Reading file: " & FilePath
    Log "Length: " & LOF(FileNumber) & " bytes"
    
    'As we process characters we wait for a whole word or statement to build up
    Dim Word As String
    Dim EndOfLine As Boolean
    
    Do
        'Check the current context to see what we should be doing with this information
        Select Case Context
            Case OZ80_CONTEXT.UNKNOWN: '-----------------------------------------------
                'Lines can begin with whitespace, labels & keywords; _
                 parenthesis, variables and numbers are not allowed
                GoSub ReadWord
                
                Select Case True
                    Case Word = vbNullString
                        '
                    Case Left$(Word, 1) = OZ80_SYNTAX_COMMENT
                        '
                    Case Left$(Word, 1) = OZ80_SYNTAX_LABEL
                        Call PushContext(LABEL)
                    Case Else
                        'If a line does not begin with a comment, label or macro name _
                        then it may only begin with a keyword
                        Call PushContext(KEYWORD)
                End Select
            
            Case OZ80_CONTEXT.COMMENT: '-----------------------------------------------
                'Comments can contain any characters, but finish at the end of the line
                'We don't need to process comments any further once read, just leave _
                 the comment context and continue
                Call PopContext
            
            Case OZ80_CONTEXT.LABEL: '-------------------------------------------------
                'Label names must begin with ":", contain A-Z, 0-9, underscore and _
                 dash with the restriction that the first letter must be A-Z or an _
                 underscore and not a numeral
                If IsValidName(Word) = False Then
                    Let ProcessFile = INVALID_LABEL: GoTo Finish
                End If
                Call PopContext
            
            Case OZ80_CONTEXT.KEYWORD: '-----------------------------------------------
                'Which keyword has been specified?
                Select Case Word
                    Case OZ80_KEYWORD_SET:      PushContext (KEYWORD_SET)
                    Case Else
                        'Unknown keyword!
                        Let ProcessFile = UNKNOWN_KEYWORD: GoTo Finish
                End Select
            
            Case OZ80_CONTEXT.KEYWORD_SET: '-------------------------------------------
                'Format: _
                        SET !<variableName> <expr>
                
                'Get the variable name
                Call PushContext(VARIABLE)
                GoSub ReadWord
                
                If IsValidName(Word) = False Then
                    Let ProcessFile = INVALID_VARIABLE: GoTo Finish
                End If
                
                Call PopContext
                Call PopContext
                
        End Select
        
    Loop While EOF(FileNumber) = False
    GoTo Finish
    
    '==================================================================================
ReadWord:
    Let Word = vbNullString
    Let EndOfLine = False
    
    Dim IsComment As Boolean
    Let IsComment = False
    
ReadChar:
    Dim Char As String * 1
    'Read a charcter. If the file ends, treat it as a remaining end of line
    If EOF(FileNumber) = True Then Let Char = vbCr Else Get #FileNumber, , Char
    
    'If the line ends, so does the word
    If IsEndOfLine(Char) = True Then
        Let EndOfLine = True
        'If the 'word' was a comment, discard it and read the next word until we _
         get something meaningful
        If IsComment = True Then GoTo ReadWord
        'Otherwise return to the context processor with the word we've extracted
        GoTo EndWord
    End If
    
    'Is this a comment? (in which case don't end the word on spaces)
    If Word = vbNullString Then
        If Char = OZ80_SYNTAX_COMMENT Then Let IsComment = True
    End If
    
    'If not a comment, end the word on a space instead of at the end of the line
    If IsComment = False Then
        If IsWhitespace(Char) = True Then GoTo EndWord
    End If
    
    Let Word = Word & Char
    GoTo ReadChar

EndWord:
    If Word <> vbNullString Then Log Word
    Return
    
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

'IsValidName : If a string is a valid label/variable/macro/object &c. name _
 ======================================================================================
Private Function IsValidName(ByVal ItemName As String) As Boolean
    Let IsValidName = True
    If ItemName Like "?[!_A-Z]*" Then IsValidName = False: Exit Function
    If ItemName Like "??*[!A-Z_0-9-]*" Then IsValidName = False: Exit Function
End Function

'Log _
 ======================================================================================
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
