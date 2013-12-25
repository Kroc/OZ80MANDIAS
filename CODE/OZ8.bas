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
Private Const OZ80_SYNTAX_NUMBER_HEX = "$"
Private Const OZ80_SYNTAX_NUMBER_BIN = "%"

Private Const OZ80_SYNTAX_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVQXYZ_"
Private Const OZ80_SYNTAX_NUMERIC = "0123456789-"

Private Const OZ80_KEYWORD_SET = "SET"
Private Const OZ80_KEYWORDS = "|" & OZ80_KEYWORD_SET & "|"

Private Const OZ80_OPERAND_ADD = "+"
Private Const OZ80_OPERAND_SUB = "-"
Private Const OZ80_OPERAND_MUL = "*"
Private Const OZ80_OPERAND_DIV = "/"
Private Const OZ80_OPERAND_POW = "^"
Private Const OZ80_OPERAND_MOD = "\"

Private Const OZ80_OPERATORS = _
    "|" & OZ80_OPERAND_ADD & "|" & OZ80_OPERAND_SUB & "|" & OZ80_OPERAND_MUL & _
    "|" & OZ80_OPERAND_DIV & "|" & OZ80_OPERAND_POW & "|" & OZ80_OPERAND_MOD & "|"
    

'--------------------------------------------------------------------------------------

Private Enum OZ80_CONTEXT
    UNKNOWN = 0
    'When inside a string, i.e. `... "some text" ...`
    QUOTED = 1
    
    NUMBER = 10
    LABEL = 11
    VARIABLE = 12
    MACRO = 13
    FUNCT = 14
    
    EXPRESSION = 100
    
    KEYWORD = 1000
    KEYWORD_SET = 1001
End Enum
#If False Then
    Private UNKNOWN, QUOTED, _
            NUMBER, LABEL, VARIABLE, MACRO, FUNCT, _
            EXPRESSION, _
            KEYWORD, KEYWORD_SET
#End If

'--------------------------------------------------------------------------------------

Public Enum OZ80_ERROR
    NOERROR = 0
    INVALID_LABEL = 1
    INVALID_VARIABLE = 2
    UNKNOWN_KEYWORD = 3
End Enum

Private FileNumber As Integer
Private Word As String
Private Context As OZ80_CONTEXT

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Assemble : Parse and assemble a file into a binary _
 ======================================================================================
Public Function Assemble( _
    ByVal FilePath As String, _
    Optional ByVal OutputPath As String = vbNullString _
) As OZ80_ERROR
    Log "OZ80MANDIAS"
    
    Log ProcessFile(FilePath)
    
    Debug.Print
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ProcessFile : Since files can be nested (with `INCLUDE`) make this process recursive _
 ======================================================================================
Private Function ProcessFile(ByVal FilePath As String) As OZ80_ERROR
    Let FileNumber = FreeFile()
    Open FilePath For Binary Access Read Lock Write As #FileNumber
    
    Log "Reading file: " & FilePath
    Log "Length: " & LOF(FileNumber) & " bytes"
    
    Do
        Call GetWord
        If Word = vbNullString Then Exit Do
        
        Call ContextRoot
    Loop
    
    Close #FileNumber
End Function

'ContextRoot : The context at the start of a line (and not within a block) _
 ======================================================================================
Private Function ContextRoot() As OZ80_CONTEXT
    'Lines can begin with labels & keywords; _
     parenthesis, variables and numbers are not allowed
     
    Call GetWord
    Select Case Context
        Case OZ80_CONTEXT.LABEL
            Call ContextLabel
            
        Case OZ80_CONTEXT.KEYWORD
            Call ContextKeyword
            
        Case Else
            
    End Select
End Function

'ContextLabel _
 ======================================================================================
Private Function ContextLabel() As OZ80_CONTEXT
    'Label names must begin with ":", contain A-Z, 0-9, underscore and _
     dash with the restriction that the first letter must be A-Z or an _
     underscore and not a numeral
End Function

'ContextKeyword _
 ======================================================================================
Private Function ContextKeyword() As OZ80_CONTEXT
    Select Case Word
        Case OZ80_KEYWORD_SET
            'Format: _
                SET !<variableName> <expr>
    End Select
End Function

'ContextVariable _
 ======================================================================================
Private Function ContextVariable() As OZ80_CONTEXT
    '
End Function

'ContextExpression _
 ======================================================================================
Private Function ContextExpression() As OZ80_CONTEXT
    'An expression is anything that results in a value, i.e. a number, _
     a label/property, a calculation, a function call &c.
End Function

'GetWord : Read the next unit of text from the file _
 ======================================================================================
Private Sub GetWord()
ReadWord:
    Let Word = vbNullString
    Dim IsComment As Boolean
    Let IsComment = False
    
ReadChar:
    Dim Char As String * 1
    'Read a charcter. If the file ends, treat it as a remaining end of line
    If EOF(FileNumber) = True Then GoTo EndWord Else Get #FileNumber, , Char
    
    'Ignore whitespace and line-breaks prior to beginning a word
    If Word = vbNullString Then
        If IsWhitespace(Char) = True Then GoTo ReadChar
        If IsEndOfLine(Char) = True Then GoTo ReadChar
        'Is this a comment? (in which case don't end the word on spaces)
        If Char = OZ80_SYNTAX_COMMENT Then Let IsComment = True
    End If
    
    'If the line ends, so does the word
    If IsEndOfLine(Char) = True Then
        'If the 'word' was a comment, discard it and read the next word until we _
         get something meaningful
        If IsComment = True Then GoTo ReadWord
        'Otherwise return to the context with the word we've extracted
        GoTo EndWord
    End If
    
    'If not a comment, end the word on a space instead of at the end of the line
    If IsComment = False Then
        If IsWhitespace(Char) = True Then GoTo EndWord
    End If
    
    Let Word = Word & Char
    GoTo ReadChar

EndWord:
    Select Case True
        Case Left$(Word, 1) = OZ80_SYNTAX_LABEL: Let Context = LABEL
        Case Left$(Word, 1) = OZ80_SYNTAX_VARIABLE: Let Context = VARIABLE
        Case Left$(Word, 1) = OZ80_SYNTAX_MACRO: Let Context = MACRO
        Case IsKeyword(Word): Let Context = KEYWORD
        Case IsNumber(Word): Let Context = NUMBER
        Case Else
            Let Context = UNKNOWN
    End Select
    If Word <> vbNullString Then Log Word
End Sub

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

'IsKeyword _
 ======================================================================================
Private Function IsKeyword(ByVal Word As String)
    Let IsKeyword = (InStr(OZ80_KEYWORDS, "|" & Word & "|") > 0)
End Function

'IsOperator _
 ======================================================================================
Private Function IsOperator(ByVal Word As String) As Boolean
    Let IsOperator = (InStr(OZ80_OPERATORS, "|" & Word & "|") > 0)
End Function

'IsLabel _
 ======================================================================================
Private Function IsLabel(ByVal Word As String) As Boolean
    If Left$(Word, 1) = OZ80_SYNTAX_LABEL Then Let IsLabel = True
End Function

'IsNumber _
 ======================================================================================
Private Function IsNumber(ByVal Word As String) As Boolean
    If Left$(Word, 1) = OZ80_SYNTAX_NUMBER_HEX Then Let IsNumber = True: Exit Function
    If Left$(Word, 1) = OZ80_SYNTAX_NUMBER_BIN Then Let IsNumber = True: Exit Function
    Let IsNumber = Not (Word Like "*[!0-9]*")
End Function

'Log _
 ======================================================================================
Private Function Log(ByVal Msg As String, Optional ByVal Depth As Long = -1)
'    If Depth = -1 Then Let Depth = ContextPointer
'    If Depth > 0 Then Debug.Print String(Depth - 1, vbTab);
    Debug.Print Msg
End Function
