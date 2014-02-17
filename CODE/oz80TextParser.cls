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
'CLASS :: oz80TextParser

'The purpose of this class is to read a source code file, parse the text and generate _
 tokens from that, fed into a given oz80TokenStream object. The assembler uses that _
 representation to work from. A certain amount of validation occurs in the parser _
 where it can, but this can't capture every possibility the assembler might run into

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

Private RegisterTokens As Scripting.Dictionary
Private OperatorTokens As Scripting.Dictionary

Private Const SYNTAX_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVQXYZ_"
Private Const SYNTAX_NUMERIC = "0123456789"

Private Const SYNTAX_COMMENT = "`"
Private Const SYNTAX_QUOTE = """"
Private Const SYNTAX_LABEL = ":"
Private Const SYNTAX_VARIABLE = "#"
Private Const SYNTAX_MACRO = "@"
Private Const SYNTAX_FUNCT = "?"
Private Const SYNTAX_NUMBER_HEX = "$"
Private Const SYNTAX_NUMBER_BIN = "%"
Private Const SYNTAX_LIST = ","
Private Const SYNTAX_PAREN_OPEN = "("
Private Const SYNTAX_PAREN_CLOSE = ")"
Private Const SYNTAX_BLOCK_OPEN = "{"
Private Const SYNTAX_BLOCK_CLOSE = "}"

Private Const SYNTAX_OPERATOR_ADD = "+"
Private Const SYNTAX_OPERATOR_SUB = "-"
Private Const SYNTAX_OPERATOR_MUL = "*"
Private Const SYNTAX_OPERATOR_DIV = "/"
Private Const SYNTAX_OPERATOR_POW = "^"
Private Const SYNTAX_OPERATOR_MOD = "\"

Private Const Operators = _
    SYNTAX_OPERATOR_ADD & SYNTAX_OPERATOR_SUB & SYNTAX_OPERATOR_MUL & _
    SYNTAX_OPERATOR_DIV & SYNTAX_OPERATOR_POW & SYNTAX_OPERATOR_MOD

'These characters can break a word without spaces to separate, _
 they also form a word on their own
Private Const SYNTAX_WORDBREAKS = _
    SYNTAX_LIST & SYNTAX_PAREN_OPEN & SYNTAX_PAREN_CLOSE & _
    SYNTAX_BLOCK_OPEN & SYNTAX_BLOCK_CLOSE & Operators

Private Const KEYWORD_DEF = "DEF"
Private Const KEYWORDS = "|" & KEYWORD_DEF & "|"

'--------------------------------------------------------------------------------------

Public Enum OZ80_ERROR
    NOERROR = 0
    INVALID_LABEL = 1
    INVALID_VARIABLE = 2
    UNKNOWN_KEYWORD = 3
End Enum

'/// VARIABLES ////////////////////////////////////////////////////////////////////////

Private Words() As String
Private index As Long                   'index number of current word being tokenised

'A tokenStream object is passed to us by the `Parse` method. _
 This is where we will feed the result of our text parsing
Private TokenStream As oz80TokenStream

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise the array of register mappings from text to data values
    Set RegisterTokens = New Scripting.Dictionary
    With RegisterTokens
        Call .Add("A", TOKEN_REGISTER_A)
        Call .Add("B", TOKEN_REGISTER_B)
        Call .Add("C", TOKEN_REGISTER_C)
        Call .Add("D", TOKEN_REGISTER_D)
        Call .Add("E", TOKEN_REGISTER_E)
        Call .Add("H", TOKEN_REGISTER_H)
        Call .Add("I", TOKEN_REGISTER_I)
        Call .Add("L", TOKEN_REGISTER_L)
        Call .Add("R", TOKEN_REGISTER_R)
        Call .Add("AF", TOKEN_REGISTER_AF)
        Call .Add("BC", TOKEN_REGISTER_BC)
        Call .Add("DE", TOKEN_REGISTER_DE)
        Call .Add("HL", TOKEN_REGISTER_HL)
        Call .Add("IXL", TOKEN_REGISTER_IXL)
        Call .Add("IXH", TOKEN_REGISTER_IXH)
        Call .Add("IX", TOKEN_REGISTER_IX)
        Call .Add("IYL", TOKEN_REGISTER_IYL)
        Call .Add("IYH", TOKEN_REGISTER_IYH)
        Call .Add("IY", TOKEN_REGISTER_IY)
        Call .Add("SP", TOKEN_REGISTER_SP)
        Call .Add("PC", TOKEN_REGISTER_PC)
        'and the shadow registers -- you can't actually manipulate these directly _
         on a z80, this is just a way of labelling for your convenience
        Call .Add("A'", TOKEN_REGISTER_A)
        Call .Add("B'", TOKEN_REGISTER_B)
        Call .Add("C'", TOKEN_REGISTER_C)
        Call .Add("D'", TOKEN_REGISTER_D)
        Call .Add("E'", TOKEN_REGISTER_E)
        Call .Add("H'", TOKEN_REGISTER_H)
        Call .Add("AF'", TOKEN_REGISTER_AF)
        Call .Add("BC'", TOKEN_REGISTER_BC)
        Call .Add("DE'", TOKEN_REGISTER_DE)
        Call .Add("HL'", TOKEN_REGISTER_HL)
    End With
    
    'Initialise the array of operator mappings from text to data values
    Set OperatorTokens = New Scripting.Dictionary
    With OperatorTokens
        Call .Add(SYNTAX_OPERATOR_ADD, OZ80_TOKEN_DATA.TOKEN_OPERATOR_ADD)
        Call .Add(SYNTAX_OPERATOR_SUB, OZ80_TOKEN_DATA.TOKEN_OPERATOR_SUB)
        Call .Add(SYNTAX_OPERATOR_MUL, OZ80_TOKEN_DATA.TOKEN_OPERATOR_MUL)
        Call .Add(SYNTAX_OPERATOR_DIV, OZ80_TOKEN_DATA.TOKEN_OPERATOR_DIV)
        Call .Add(SYNTAX_OPERATOR_POW, OZ80_TOKEN_DATA.TOKEN_OPERATOR_POW)
        Call .Add(SYNTAX_OPERATOR_MOD, OZ80_TOKEN_DATA.TOKEN_OPERATOR_MOD)
    End With
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the arrays
    Erase Words
    Set RegisterTokens = Nothing
    Set OperatorTokens = Nothing
    'Dereference the tokenStream object we were given by the assembler
    Set TokenStream = Nothing
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Parse : Parse source code into a token stream usable by the assembler _
 ======================================================================================
Public Function Parse( _
    ByVal FilePath As String, _
    ByRef DestTokenStream As oz80TokenStream _
) As OZ80_ERROR
    'Remember the token stream we need to feed into
    Set TokenStream = DestTokenStream
    
    Dim Start As Single
    Let Start = Timer
    
    'Split the source file into separate words for our "Look-Ahead Left Right" parser
    ExplodeFile (FilePath)
    
    Debug.Print Format$(Timer - Start, "0.00")
    
    'Now parse the words, validate and tokenise them
'    Let index = 0
'    Do While index < UBound(Words)
'        Call ContextRoot
'    Loop
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ExplodeFile : Break a file into all its individual words for tokenising _
 ======================================================================================
Private Function ExplodeFile(ByVal FilePath As String) As OZ80_ERROR
    Dim IsQuote As Boolean
    Dim IsComment As Boolean
    Dim IsCommentMulti As Boolean
    
    Dim AscQuote As Long
    Let AscQuote = AscW(SYNTAX_QUOTE)
    Dim AscComment As Long
    Let AscComment = AscW(SYNTAX_COMMENT)
    Dim AscList As Long
    Let AscList = AscW(SYNTAX_LIST)
    Dim AscParenOpen As Long, AscParenClose As Long
    Let AscParenOpen = AscW(SYNTAX_PAREN_OPEN)
    Let AscParenClose = AscW(SYNTAX_PAREN_CLOSE)
    Dim AscBlockOpen As Long, AscBlockClose As Long
    Let AscBlockOpen = AscW(SYNTAX_BLOCK_OPEN)
    Let AscBlockClose = AscW(SYNTAX_BLOCK_CLOSE)
    Dim AscOpAdd As Long, AscOpSub As Long, AscOpMul As Long, _
        AscOpDiv As Long, AscOpPow As Long, AscOpMod As Long
    Let AscOpAdd = AscW(SYNTAX_OPERATOR_ADD)
    Let AscOpSub = AscW(SYNTAX_OPERATOR_SUB)
    Let AscOpMul = AscW(SYNTAX_OPERATOR_MUL)
    Let AscOpDiv = AscW(SYNTAX_OPERATOR_DIV)
    Let AscOpPow = AscW(SYNTAX_OPERATOR_POW)
    Let AscOpMod = AscW(SYNTAX_OPERATOR_MOD)
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    Dim Text As String
    Dim FileNumber As String
    Let FileNumber = FreeFile
    
    Debug.Print "Exploding file: " & FilePath
    Open FilePath For Input Access Read Lock Write As #FileNumber
    Let Text = Space$(LOF(FileNumber))
    Let Text = StrConv(InputB(LOF(FileNumber), FileNumber), vbUnicode)
    'Cache this for speed
    Dim TextLen As Long: Let TextLen = Len(Text)
    Debug.Print "- Length: " & Format$(TextLen, "#,#") & " bytes"
    Close #FileNumber
    
    Dim TextPos As Long
    'Provide a reasonable sized buffer of words, doing `ReDim` for every word is _
     *incredibly* slow
    Erase Words
    Dim WordLimit As Long: Let WordLimit = 100000
    ReDim Words(0 To WordLimit) As String
    
    '----------------------------------------------------------------------------------
ReadWord:
    Dim Word As String
    Let Word = vbNullString
    
ReadChar:
    'If the file ends, treat it as a remaining end of line
    If TextPos >= TextLen Then GoTo EndWord
    
    'Look at the current character _
     (without moving forward to the next)
    Dim Char As String
    Let TextPos = TextPos + 1
    Let Char = Mid$(Text, TextPos, 1)
    Dim AscChar As Long
    Let AscChar = AscW(Char)
    
    'skip through comments...
    '----------------------------------------------------------------------------------
    If IsComment = True Then
        'For single-line comments:
        If IsCommentMulti = False Then
            'The end of a line ends a single-line comment:
            If AscChar = 10 Or AscChar = 13 Then
                Let IsComment = False: GoTo ReadWord
            End If
        
        'For multi-line comments, check for backtick:
        ElseIf AscChar = AscComment Then
            'Check for a second back-tick following
            Let TextPos = TextPos + 1
            Let Char = Mid$(Text, TextPos, 1)
            If AscW(Char) = AscComment Then
                Let IsComment = False: Let IsCommentMulti = False
                GoTo ReadWord
            End If
        End If
        'Skip to the next character
        GoTo ReadChar
    
    '----------------------------------------------------------------------------------
    '(false checks are slightly faster)
    ElseIf IsQuote = False Then
        'Special handling for the beginning of a word
        If LenB(Word) = 0 Then
            If AscChar = 32 Then GoTo ReadChar
            If AscChar = 9 Then GoTo ReadChar
            If AscChar = 10 Then GoTo ReadChar
            If AscChar = 13 Then GoTo ReadChar
        Else
            'A word is underway, check for end of line (note the `Else`, we don't want _
             to test line endings twice when the Word is currently empty)
            If AscChar = 10 Or AscChar = 13 Then
                'The end of a line ends a single-line comment:
                If IsCommentMulti = False Then
                    If IsComment = True Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo EndWord2
            End If
        End If
        
        'End the word on a space instead of at the end of the line
        If AscChar = 32 Then GoTo EndWord
        If AscChar = 9 Then GoTo EndWord
        
        'Is this a quote? (in which case, ignore wordbreaks until quote end)
        If AscChar = AscQuote Then Let IsQuote = True: GoTo NextChar
        
        'Check for comment marker
        If AscChar = AscComment Then
            Let IsComment = True
            If LenB(Word) Then
                'Is this a multi-line comment? (two backticks)
                If AscW(Word) = AscComment Then
                    Let IsCommentMulti = True
                    GoTo ReadWord
                End If
                'If a comment begins at the end of a word without any space to _
                 separate, we need to first keep the word before switching to _
                 comment parsing
                GoTo WordBreak
            End If
            GoTo NextChar
        End If
    
        'Symbols that can break a word in the middle without spaces to separate
        If AscChar = AscList Then GoTo WordBreak
        If AscChar = AscParenOpen Then GoTo WordBreak
        If AscChar = AscParenClose Then GoTo WordBreak
        If AscChar = AscBlockOpen Then GoTo WordBreak
        If AscChar = AscBlockClose Then GoTo WordBreak
        
        If AscChar = AscOpAdd Then GoTo WordBreak
        If AscChar = AscOpSub Then GoTo WordBreak
        If AscChar = AscOpMul Then GoTo WordBreak
        If AscChar = AscOpDiv Then GoTo WordBreak
        If AscChar = AscOpPow Then GoTo WordBreak
        If AscChar = AscOpMod Then GoTo WordBreak
    
    'Look for the end of a quote
    '----------------------------------------------------------------------------------
    Else
        'TODO: handle escape chars, skip tabs and line breaks
        If AscChar = AscQuote Then Let IsQuote = False: GoTo EndWord
        
    End If
    
    '----------------------------------------------------------------------------------
    
NextChar:
    Let Word = Word & Char
    GoTo ReadChar

WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    If LenB(Word) Then
        Let TextPos = TextPos - 1
        GoTo EndWord2
    End If
    Let Word = Word & Char
    
    '----------------------------------------------------------------------------------
EndWord:
    If LenB(Word) = 0 Then GoTo EndOfFile
EndWord2:
    Dim WordIndex As Long
    If WordIndex > WordLimit Then
        Let WordLimit = WordLimit + 50000
        ReDim Preserve Words(0 To WordLimit) As String
    End If
    Let Words(WordIndex) = Word
    Let WordIndex = WordIndex + 1
    
    GoTo ReadWord
    
EndOfFile:
    Let Text = vbNullString
    Debug.Print "- Words: " & Format$(WordIndex, "#,#")
End Function

'/// CONTEXT PROCEDURES ///////////////////////////////////////////////////////////////
'After the file is exploded into separate words, we walk through the word list _
 validating the "context", that is, parameters should follow keywords &c. _
 This process is recursive in nature because blocks can be nested, so the code is _
 divided into separate contexts that only concern themselves with the requirements _
 of their context and what should follow next

''ContextRoot : The context at the start of a line (and not within a block) _
' ======================================================================================
'Private Sub ContextRoot()
'    'Lines can begin with Z80 code, labels & keywords
'
'    'Check the current word type
'    Select Case Words(index).Kind
'        'If the word appears to be an assembly instruction:
'        Case OZ80_WORDKIND.ASM
'            Call ContextAssembly
'
'        Case Else
'            Let index = index + 1
'    End Select
'End Sub
'
''ContextAssembly : Parse Z80 assembly source _
' ======================================================================================
'Private Sub ContextAssembly()
'    Dim Word As String
'    Let Word = Words(index).Word
'
'    'Check the mneomic, this is divided into two main lists of those that have _
'     parameters and those that do not
'    Select Case Word
'        Case "adc":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_ADC)
'        Case "add":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_ADD)
'        Case "and":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_AND)
'        Case "bit":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_BIT)
'        Case "call":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CALL)
'        Case "cp":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CP)
'        Case "dec":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DEC)
'        Case "djnz":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DJNZ)
'        Case "ex":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EX)
'        Case "im":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IM)
'        Case "in":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IN)
'        Case "inc":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INC)
'        Case "jp":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_JP)
'        Case "jr":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_JR)
'        Case "ld":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LD)
'        Case "or":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OR)
'        Case "out":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUT)
'        Case "pop":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_POP)
'        Case "push":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_PUSH)
'        Case "res":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RES)
'        Case "ret":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RET)
'        Case "rl":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RL)
'        Case "rlc":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLC)
'        Case "rr":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RR)
'        Case "rrc":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRC)
'        Case "rst":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RST)
'        Case "sbc":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SBC)
'        Case "set":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SET)
'        Case "sla":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SLA)
'        Case "sra":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SRA)
'        Case "sll":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SLL)
'        Case "srl":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SRL)
'        Case "sub":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SUB)
'        Case "xor":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_XOR)
'        Case Else
'            'If none of the above, skip ahead to instructions without parameters
'            GoTo NoParam
'    End Select
'    'The mneomic has been tokenised, move to the next word
'    Let index = index + 1
'
'    'Get the parameter list: _
'     The parser is not concerned with the number of parameters and if they are the _
'     right type, the assembler will handle that
'    Call ContextList
'
'    Exit Sub
'
'NoParam:
'    Select Case Word
'        Case "ccf":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CCF)
'        Case "cpd":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPD)
'        Case "cpdr":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPDR)
'        Case "cpi":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPI)
'        Case "cpir":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPIR)
'        Case "cpl":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_CPL)
'        Case "daa":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DAA)
'        Case "di":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_DI)
'        Case "ei":      'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EI)
'        Case "exx":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_EXX)
'        Case "halt":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_HALT)
'        Case "ind":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_IND)
'        Case "indr":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INDR)
'        Case "ini":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INI)
'        Case "inir":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_INIR)
'        Case "ldd":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDD)
'        Case "lddr":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDDR)
'        Case "ldi":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDI)
'        Case "ldir":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_LDIR)
'        Case "neg":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_NEG)
'        Case "nop":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_NOP)
'        Case "outd":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTD)
'        Case "outdr":   'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTDR)
'        Case "outi":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTI)
'        Case "outir":   'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_OUTIR)
'        Case "reti":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RETI)
'        Case "retn":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RETN)
'        Case "rla":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLA)
'        Case "rlca":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLCA)
'        Case "rld":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RLD)
'        Case "rra":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRA)
'        Case "rrca":    'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRCA)
'        Case "rrd":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_RRD)
'        Case "scf":     'Call TokenStream.AddToken(TOKEN_Z80, TOKEN_Z80_SCF)
'        Case Else
'            'If none of these, the developer may have made a typo or put an assembly _
'             instruction where they shouldn't have
'            'TODO: error / fold back?
'    End Select
'
'    'Since no parameter is needed, we can return to our parent context to continue
'    Let index = index + 1
'End Sub
'
''ContextExpression _
' ======================================================================================
''An expression is anything that results in a value, i.e. a number, a label/property, _
' a calculation, a function call &c.
'Private Sub ContextExpression()
'NextExpr:
'    'Check what type the parser assigned to the current word
'    Select Case Words(index).Kind
'        'Process a number
'        Case OZ80_WORDKIND.Number
'            Call ContextNumber
'
'        'Process parenthesis (a nested expression)
'        Case OZ80_WORDKIND.PAREN
'            'TODO: check if open parenthesis, otherwise fold
'
'        'When a list marker "," is found, the expression is over -- but the comma may _
'         be unexpected by the parent context (if not a list)
'        Case OZ80_WORDKIND.LIST
'            Exit Sub
'
'        Case Else
'            'Not an expression!
'            'TODO: Error out
'
'    End Select
'
'    'Now check for an operator that could extend the expression
'    If Words(index).Kind = OZ80_WORDKIND.OPERATOR Then
'        'Add the operator to the token stream,
''        Call TokenStream.AddToken(TOKEN_OPERATOR, OperatorTokens.Item(Words(index).Word))
'        'Move to the next word,
'        Let index = index + 1
'        'And get the next expression
'        GoTo NextExpr
'    End If
'End Sub
'
''ContextList _
' ======================================================================================
''A list is a common construct and consists of one or more expressions separated _
' by commas. The parameters to any instruction/directive is a list
'Private Sub ContextList()
'    'TODO: Add a list token and remember its position so we can update the item count
''    Call TokenStream.AddToken(TOKEN_LIST, 0)
'
'NextItem:
'    Call ContextExpression
'    'The above call will move to the next word in the list, check if it's a comma _
'     to continue the list, or something else marking the end of the list
'    If Words(index).Kind = OZ80_WORDKIND.LIST Then
'        'Add a token to mark the end of a list item
''        Call TokenStream.AddToken(TOKEN_LIST_NEXT)
'        'TODO: Update the item count in the list
'        GoTo NextItem
'    End If
'    'When the list ends, we fold back to the parent context to deal with it
'End Sub
'
''ContextNumber : Parse numbers (decimal, hexadecimal, binary) _
' ======================================================================================
'Private Sub ContextNumber()
'    Dim Word As String
'    Let Word = Words(index).Word
'
'    On Error Resume Next
'    Dim Number As Long
'
'    'When the source file was exploded into words the number was already validated, _
'     we can be assured that it's parsable as a number, but it could be out of bounds
'
'    'TODO: Check bounds
'    'TODO: Can we handle 32-bit numbers properly?
'
'    If Left$(Word, 1) = SYNTAX_NUMBER_HEX Then
'        'Convert the string from hexadecimal
'        Let Number = CLng("&H" & Mid$(Word, 2) & "&")
'
'    ElseIf Left$(Word, 1) = SYNTAX_NUMBER_BIN Then
'        'Convert the string from binary
'        Dim i As Long
'        For i = 1 To Len(Word)
'            'TODO: optimise this, starting with a 2^X lookup table _
'                -- or easier, have a lookup table of all 8-bit binary numbers
'            Let Number = Number + (Mid$(Word, Len(Word) - i + 1, 1) * (2 ^ (i - 1)))
'        Next i
'
'    Else
'        Let Number = CLng(Word)
'    End If
'
''    Call TokenStream.AddToken(TOKEN_NUMBER, Number)
'    Let index = index + 1
'    On Error GoTo 0
'End Sub

'/// VALIDATION PROCEDURES ////////////////////////////////////////////////////////////

''IsKeyword _
' ======================================================================================
'Private Function IsKeyword(ByVal Word As String)
'    Let IsKeyword = (InStr(KEYWORDS, "|" & Word & "|") > 0)
'End Function
'
''IsLabel _
' ======================================================================================
'Private Function IsLabel(ByVal Word As String) As Boolean
'    If Left$(Word, 1) = SYNTAX_LABEL Then Let IsLabel = True
'End Function
'
''IsNumber _
' ======================================================================================
'Private Function IsNumber(ByVal Word As String) As Boolean
'    'A hexadecimal number?
'    If Left$(Word, 1) = SYNTAX_NUMBER_HEX Then
'        'Ensure that the number consists of 0-9, A-F only
'        Let IsNumber = Not (Word Like "$*[!A-F0-9]*")
'        Exit Function
'    End If
'    'A binary number?
'    If Left$(Word, 1) = SYNTAX_NUMBER_BIN Then
'        Let IsNumber = Not (Word Like "%*[!0-1]*")
'        Exit Function
'    End If
'    'Decimal number:
'    Let IsNumber = Not (Word Like "*[!0-9]*")
'End Function
'
''IsOperator _
' ======================================================================================
'Private Function IsOperator(ByVal Word As String) As Boolean
'    Let IsOperator = (InStr(Operators, Word) > 0)
'End Function
'
''IsRegister _
' ======================================================================================
'Private Function IsRegister(ByVal Word As String) As Boolean
'    'TODO
'End Function
'
''IsValidName : If a string is a valid label/variable/macro/object &c. name _
' ======================================================================================
'Private Function IsValidName(ByVal ItemName As String) As Boolean
'    Let IsValidName = True
'    'Must begin with A-Z or underscore and not a number
'    If ItemName Like "?[!_A-Z]*" Then IsValidName = False: Exit Function
'    'Must only contain A-Z, underscore and 0-9
'    If ItemName Like "??*[!A-Z_0-9]*" Then IsValidName = False: Exit Function
'End Function
