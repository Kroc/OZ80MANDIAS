VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80TextParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: OZ8_Parser

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

Private REGISTERS As Scripting.Dictionary

Private Const SYNTAX_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVQXYZ_"
Private Const SYNTAX_NUMERIC = "0123456789"

Private Const SYNTAX_COMMENT = ";"
Private Const SYNTAX_COMMENT_OPEN = "/*"
Private Const SYNTAX_COMMENT_CLOSE = "*/"
Private Const SYNTAX_LABEL = ":"
Private Const SYNTAX_VARIABLE = "#"
Private Const SYNTAX_MACRO = "@"
Private Const SYNTAX_NUMBER_HEX = "$"
Private Const SYNTAX_NUMBER_BIN = "%"
Private Const SYNTAX_LIST = ","
Private Const SYNTAX_PAREN_OPEN = "("
Private Const SYNTAX_PAREN_CLOSE = ")"

Private Const OPERATOR_ADD = "+"
Private Const OPERATOR_SUB = "-"
Private Const OPERATOR_MUL = "*"
Private Const OPERATOR_DIV = "/"
Private Const OPERATOR_POW = "^"
Private Const OPERATOR_MOD = "\"

Private Const OPERATORS = _
    OPERATOR_ADD & OPERATOR_SUB & OPERATOR_MUL & OPERATOR_DIV & OPERATOR_POW & _
    OPERATOR_MOD

'These characters can break a word without spaces to separate, _
 they also form a word on their own
Private Const SYNTAX_WORDBREAKS = _
    SYNTAX_COMMENT & SYNTAX_LIST & SYNTAX_PAREN_OPEN & SYNTAX_PAREN_CLOSE & _
    OPERATORS

Private Const KEYWORD_DEF = "DEF"
Private Const KEYWORDS = "|" & KEYWORD_DEF & "|"

'--------------------------------------------------------------------------------------

'Private Enum OZ80_CONTEXT
'    'When parsing Z80 code
'    ASM = 0
'    'When inside a string, i.e. `... "some text" ...`
'    QUOTED = 1
'
'    NUMBER = 10
'    LABEL = 11
'    VARIABLE = 12
'    MACRO = 13
'    FUNCT = 14
'
'    EXPRESSION = 100
'
'    KEYWORD = 1000
'    KEYWORD_SET = 1001
'End Enum
'#If False Then
'    Private ASM, QUOTED, _
'            NUMBER, LABEL, VARIABLE, MACRO, FUNCT, _
'            EXPRESSION, _
'            KEYWORD, KEYWORD_SET
'#End If

'--------------------------------------------------------------------------------------

Public Enum OZ80_ERROR
    NOERROR = 0
    INVALID_LABEL = 1
    INVALID_VARIABLE = 2
    UNKNOWN_KEYWORD = 3
End Enum

'/// VARIABLES ////////////////////////////////////////////////////////////////////////

Private Type OZ80_WORD
    Word    As String
    Line    As Long
    Col     As Long
    Kind    As Byte                     '=OZ80_WORD_KIND
End Type

Private Enum OZ80_WORD_KIND
    ASM = 1
    Number = 2
    KEYWORD = 3
    Label = 4
    VARIABLE = 5
    MACRO = 6
    FUNCT = 7
End Enum

Private Words() As OZ80_WORD

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Parse : Parse source code into a token stream usable by the assembler _
 ======================================================================================
Public Function Parse( _
    ByVal FilePath As String, _
    ByRef TokenStream As oz80TokenStream _
) As OZ80_ERROR
    'Initialise the array of register mappings from text to data values
    Set REGISTERS = Nothing
    Set REGISTERS = New Scripting.Dictionary
    Call REGISTERS.Add("A", TOKEN_REGISTER_A)
    Call REGISTERS.Add("B", TOKEN_REGISTER_B)
    Call REGISTERS.Add("C", TOKEN_REGISTER_C)
    Call REGISTERS.Add("D", TOKEN_REGISTER_D)
    Call REGISTERS.Add("E", TOKEN_REGISTER_E)
    Call REGISTERS.Add("H", TOKEN_REGISTER_H)
    Call REGISTERS.Add("I", TOKEN_REGISTER_I)
    Call REGISTERS.Add("L", TOKEN_REGISTER_L)
    Call REGISTERS.Add("R", TOKEN_REGISTER_R)
    Call REGISTERS.Add("AF", TOKEN_REGISTER_AF)
    Call REGISTERS.Add("BC", TOKEN_REGISTER_BC)
    Call REGISTERS.Add("DE", TOKEN_REGISTER_DE)
    Call REGISTERS.Add("HL", TOKEN_REGISTER_HL)
    Call REGISTERS.Add("IXL", TOKEN_REGISTER_IXL)
    Call REGISTERS.Add("IXH", TOKEN_REGISTER_IXH)
    Call REGISTERS.Add("IX", TOKEN_REGISTER_IX)
    Call REGISTERS.Add("IYL", TOKEN_REGISTER_IYL)
    Call REGISTERS.Add("IYH", TOKEN_REGISTER_IYH)
    Call REGISTERS.Add("IY", TOKEN_REGISTER_IY)
    Call REGISTERS.Add("SP", TOKEN_REGISTER_SP)
    Call REGISTERS.Add("PC", TOKEN_REGISTER_PC)
    '(and the shadow registers)
    Call REGISTERS.Add("A'", TOKEN_REGISTER_A)
    Call REGISTERS.Add("B'", TOKEN_REGISTER_B)
    Call REGISTERS.Add("C'", TOKEN_REGISTER_C)
    Call REGISTERS.Add("D'", TOKEN_REGISTER_D)
    Call REGISTERS.Add("E'", TOKEN_REGISTER_E)
    Call REGISTERS.Add("H'", TOKEN_REGISTER_H)
    Call REGISTERS.Add("AF'", TOKEN_REGISTER_AF)
    Call REGISTERS.Add("BC'", TOKEN_REGISTER_BC)
    Call REGISTERS.Add("DE'", TOKEN_REGISTER_DE)
    Call REGISTERS.Add("HL'", TOKEN_REGISTER_HL)
    
    Debug.Print ExplodeFile(FilePath)
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ExplodeFile : Break a file into all its individual words for tokenising _
 ======================================================================================
Private Function ExplodeFile(ByVal FilePath As String) As OZ80_ERROR
    Dim Text As String
    
    Dim FileNumber As String
    Let FileNumber = FreeFile
    Debug.Print "Exploding file: " & FilePath
    Open FilePath For Input Access Read Lock Write As #FileNumber
    Let Text = StrConv(InputB(VBA.LOF(FileNumber), FileNumber), vbUnicode)
    Debug.Print "Length: " & LOF(FileNumber) & " bytes"
    Close #FileNumber
    
    Dim TextPos As Long
    Erase Words
    
    Dim IsQuoted As Boolean
    Dim IsComment As Boolean
    
    '----------------------------------------------------------------------------------
ReadWord:
    Dim Word As String
    Let Word = vbNullString
    
ReadChar:
    Dim Char As String
    'Read a charcter. If the file ends, treat it as a remaining end of line
    If TextPos >= Len(Text) Then
        GoTo EndWord
    Else
        'Look at the current character _
         (without moving forward to the next)
        Let Char = Mid$(Text, TextPos + 1, 1)
    End If
    
    'Special handling for the beginning of a word:
    If Word = vbNullString Then
        'Ignore whitespace and line-breaks prior to beginning a word
        If IsWhitespace(Char) = True Then GoTo SkipChar
        If IsEndOfLine(Char) = True Then GoTo SkipChar
    End If
    
    'If the line ends, so does the word
    If IsEndOfLine(Char) = True Then
        'If the 'word' was a comment, discard it and read the next word until we _
         get something meaningful
        If IsComment = True Then
            Let IsComment = False
            GoTo SkipWord
        End If
        'Otherwise return to the context with the word we've extracted
        GoTo EndWord
    End If
    
    'If not a comment, ...
    If IsComment = False Then
        'Is this a comment? (in which case don't end the word on spaces)
        If Char = SYNTAX_COMMENT Then Let IsComment = True
        'End the word on a space instead of at the end of the line
        If IsWhitespace(Char) = True Then GoTo EndWord
        'Comments, commas and operators cause the word to end without spaces
        If InStr(SYNTAX_WORDBREAKS, Char) > 0 Then
            'If the word is longer than one character, then save the current character _
             for use in the next word. i.e. when a comma is reached, process the text _
             before the comma as a word, and take the comma as the next word
            If Word <> vbNullString Then GoTo EndWordEarly
            Let Word = Word & Char
            GoTo EndWord
        End If
    End If
    
    Let TextPos = TextPos + 1
    Let Word = Word & Char
    GoTo ReadChar

SkipChar:
    Let TextPos = TextPos + 1
    GoTo ReadChar
    
SkipWord:
    Let TextPos = TextPos + 1
    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
EndWord:
    Let TextPos = TextPos + 1
EndWordEarly:
    If Word = vbNullString Then GoTo EndOfFile
    
    Dim IsDimmed As Boolean
    If IsDimmed = False Then
        ReDim Words(0) As OZ80_WORD
        Let IsDimmed = True
    Else
        ReDim Words(UBound(Words) + 1) As OZ80_WORD
    End If
    With Words(UBound(Words))
        Let .Word = Word
        
        Select Case True
            Case Left$(Word, 1) = SYNTAX_LABEL:     Let .Kind = OZ80_WORD_KIND.Label
            Case Left$(Word, 1) = SYNTAX_VARIABLE:  Let .Kind = OZ80_WORD_KIND.VARIABLE
            Case Left$(Word, 1) = SYNTAX_MACRO:     Let .Kind = OZ80_WORD_KIND.MACRO
            Case IsKeyword(Word):                   Let .Kind = OZ80_WORD_KIND.KEYWORD
            Case IsNumber(Word):                    Let .Kind = OZ80_WORD_KIND.Number
            Case Else:                              Let .Kind = OZ80_WORD_KIND.ASM
        End Select
        
        Debug.Print Word
    End With
    
    GoTo ReadWord
    
EndOfFile:
    Let Text = vbNullString
End Function

'/// CONTEXT PROCEDURES ///////////////////////////////////////////////////////////////

''ContextRoot : The context at the start of a line (and not within a block) _
' ======================================================================================
'Private Function ContextRoot() As OZ80_CONTEXT
'    'Lines can begin with Z80 code, labels & keywords; _
'     but parenthesis, variables and numbers are not allowed
'
'    Select Case Context
'        Case OZ80_CONTEXT.ASM
'            Call ContextAssembly
'
'        Case OZ80_CONTEXT.LABEL
'            Call ContextLabel
'
'        Case OZ80_CONTEXT.KEYWORD
'            Call ContextKeyword
'
'        Case Else
'            'TODO: In the instance of something sitting at root context that's _
'             unrecognised we can't fold any further back in the context tree so _
'             we error here?
'
'    End Select
'End Function
'
''ContextAssembly : Parse Z80 assembly source _
' ======================================================================================
'Private Function ContextAssembly() As OZ80_CONTEXT
'    'Check the mneomic, this is divided into two main lists of those that have _
'     parameters and those that do not
'    Select Case Word
'        Case "adc":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_ADC)
'        Case "add":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_ADD)
'        Case "and":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_AND)
'        Case "bit":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_BIT)
'        Case "call":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CALL)
'        Case "cp":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CP)
'        Case "dec":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DEC)
'        Case "djnz":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DJNZ)
'        Case "ex":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EX)
'        Case "im":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IM)
'        Case "in":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IN)
'        Case "inc":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INC)
'        Case "jp":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_JP)
'        Case "jr":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_JR)
'        Case "ld":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LD)
'        Case "or":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OR)
'        Case "out":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUT)
'        Case "pop":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_POP)
'        Case "push":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_PUSH)
'        Case "res":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RES)
'        Case "ret":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RET)
'        Case "rl":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RL)
'        Case "rlc":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLC)
'        Case "rr":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RR)
'        Case "rrc":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRC)
'        Case "rst":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RST)
'        Case "sbc":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SBC)
'        Case "set":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SET)
'        Case "sla":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SLA)
'        Case "sra":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SRA)
'        Case "sll":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SLL)
'        Case "srl":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SRL)
'        Case "sub":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SUB)
'        Case "xor":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_XOR)
'        Case Else
'            'if none of the above, skip ahead to instructions without parameters
'            GoTo NoParam
'    End Select
'
'    'get the parameter list. the parse is not concerned with the number of parameters _
'     and if they are the right type, the assembler will handle that
'    Select Case ContextList()
'        'TODO: handle tail context
'    End Select
'
'    Exit Function
'
'NoParam:
'    Select Case Word
'        Case "ccf":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CCF)
'        Case "cpd":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPD)
'        Case "cpdr":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPDR)
'        Case "cpi":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPI)
'        Case "cpir":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPIR)
'        Case "cpl":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPL)
'        Case "daa":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DAA)
'        Case "di":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DI)
'        Case "ei":      Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EI)
'        Case "exx":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EXX)
'        Case "halt":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_HALT)
'        Case "ind":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IND)
'        Case "indr":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INDR)
'        Case "ini":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INI)
'        Case "inir":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INIR)
'        Case "ldd":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDD)
'        Case "lddr":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDDR)
'        Case "ldi":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDI)
'        Case "ldir":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDIR)
'        Case "neg":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_NEG)
'        Case "nop":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_NOP)
'        Case "outd":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTD)
'        Case "outdr":   Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTDR)
'        Case "outi":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTI)
'        Case "outir":   Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTIR)
'        Case "reti":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RETI)
'        Case "retn":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RETN)
'        Case "rla":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLA)
'        Case "rlca":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLCA)
'        Case "rld":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLD)
'        Case "rra":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRA)
'        Case "rrca":    Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRCA)
'        Case "rrd":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRD)
'        Case "scf":     Call OZ80_TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SCF)
'    End Select
'
'    'TODO: tail?
'End Function
'
''ContextExpression _
' ======================================================================================
'Private Function ContextExpression() As OZ80_CONTEXT
'    'An expression is anything that results in a value, i.e. a number, _
'     a label/property, a calculation, a function call &c.
'
'    'Read a word and check the suggested context
'    Call GetWord
'    Select Case Context
'
'    End Select
'End Function
'
''ContextKeyword _
' ======================================================================================
'Private Function ContextKeyword() As OZ80_CONTEXT
'    Select Case Word
'        Case KEYWORD_DEF
'            'Format: _
'                DEF #<variableName> <expr>
'
'            Call GetWord
'            If Context <> VARIABLE Then
'                '
'            End If
'
'            Call ContextVariable
'            Call ContextExpression
'
'    End Select
'End Function
'
''ContextLabel _
' ======================================================================================
'Private Function ContextLabel() As OZ80_CONTEXT
'    'Label names must begin with ":", contain A-Z, 0-9, underscore and dash with the _
'     restriction that the first letter must be A-Z or an underscore
'End Function
'
''ContextList _
' ======================================================================================
'Private Function ContextList() As OZ80_CONTEXT
'    'A list is a common construct and consists of one or more expressions separated _
'     by commas. The parameters to any instruction/directive is a list
'
'    'Begin with reading an expression
'    Select Case ContextExpression()
'        'TODO: check if the tail was a comma, if so continue reading
'    End Select
'End Function
'
''ContextVariable _
' ======================================================================================
'Private Function ContextVariable() As OZ80_CONTEXT
'    'TODO
'End Function

'/// VALIDATION PROCEDURES ////////////////////////////////////////////////////////////

'IsEndOfLine _
 ======================================================================================
Private Function IsEndOfLine(ByVal Char As String) As Boolean
    'For speed we won't use an OR statement as both comparisons are executed
    If Char = vbCr Then Let IsEndOfLine = True: Exit Function
    If Char = vbLf Then Let IsEndOfLine = True: Exit Function
    If Asc(Char) = 0 Then Let IsEndOfLine = True: Exit Function
End Function

'IsKeyword _
 ======================================================================================
Private Function IsKeyword(ByVal Word As String)
    Let IsKeyword = (InStr(KEYWORDS, "|" & Word & "|") > 0)
End Function

'IsLabel _
 ======================================================================================
Private Function IsLabel(ByVal Word As String) As Boolean
    If Left$(Word, 1) = SYNTAX_LABEL Then Let IsLabel = True
End Function

'IsNumber _
 ======================================================================================
Private Function IsNumber(ByVal Word As String) As Boolean
    If Left$(Word, 1) = SYNTAX_NUMBER_HEX Then Let IsNumber = True: Exit Function
    If Left$(Word, 1) = SYNTAX_NUMBER_BIN Then Let IsNumber = True: Exit Function
    Let IsNumber = Not (Word Like "*[!0-9]*")
End Function

'IsOperator _
 ======================================================================================
Private Function IsOperator(ByVal Word As String) As Boolean
    Let IsOperator = (InStr(OPERATORS, Word) > 0)
End Function

'IsRegister _
 ======================================================================================
Private Function IsRegister(ByVal Word As String) As Boolean
    'TODO
End Function

'IsValidName : If a string is a valid label/variable/macro/object &c. name _
 ======================================================================================
Private Function IsValidName(ByVal ItemName As String) As Boolean
    Let IsValidName = True
    'Must begin with A-Z or underscore and not a number
    If ItemName Like "?[!_A-Z]*" Then IsValidName = False: Exit Function
    'Must only contain A-Z, underscore and 0-9
    If ItemName Like "??*[!A-Z_0-9]*" Then IsValidName = False: Exit Function
End Function

'IsWhitespace : check for meaningless whitespace (space, tab) _
 ======================================================================================
Private Function IsWhitespace(ByVal Char As String) As Boolean
    'For speed we won't use an OR statement as both comparisons are executed
    If Char = " " Then Let IsWhitespace = True: Exit Function
    If Char = vbTab Then Let IsWhitespace = True: Exit Function
End Function


