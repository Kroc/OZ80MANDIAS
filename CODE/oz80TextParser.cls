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
Private WordCount As Long
Private index As Long                   'index number of current word being tokenised

'A tokenStream object is passed to us by the `Parse` method. _
 This is where we will feed the result of our text parsing
Private TokenStream As oz80TokenStream

'A look up table of ASCII codes from upper-case to lower-case
Private LowCase(0 To 255) As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise the lookup table of upper-case to lower-case values
    Dim i As Long
    For i = 0 To 255
        If i >= 65 And i <= 90 Then
            Let LowCase(i) = i And 32
        Else
            Let LowCase(i) = i
        End If
    Next i
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the arrays
    Erase Words
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
    
    'Split the source file into separate words for our "Look-Ahead Left Right" parser
    ExplodeFile (FilePath)
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'ExplodeFile : Break a file into all its individual words for tokenising _
 ======================================================================================
Private Function ExplodeFile(ByVal FilePath As String) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices" or _
     even structured code. Gotos are fast, deal with it
     
    Dim Start As Single
    Let Start = Timer
    
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
    Dim AscNumHex As Long, AscNumBin As Long
    Let AscNumHex = AscW(SYNTAX_NUMBER_HEX)
    Let AscNumBin = AscW(SYNTAX_NUMBER_BIN)
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    Dim Text As String
    Dim FileNumber As String
    Let FileNumber = FreeFile
    
    'TODO: We could read a binary string into a SAFEARRAY1D and access the byte _
           values directly for a speed boost
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
    Let WordCount = 0
    
    'The first letter of a word will be remembered for guiding tokenisation
    Dim First As String
    Dim AscFirst As Long
    
    '2nd, 3rd and 4th letters are used in validating registers
    Dim LetterIndex As Long
    Dim Asc2 As Long, Asc3 As Long, Asc4 As Long
    
    '----------------------------------------------------------------------------------
ReadWord:
    'The word as we build it
    Dim Word As String: Let Word = vbNullString
    'Reset the counter and captures for the 2nd, 3rd and 4th characters of the word
    Let LetterIndex = 0
    Let Asc2 = 0: Let Asc3 = 0: Let Asc4 = 0
    
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
    If IsComment Then
        'For single-line comments:
        If IsCommentMulti = False Then
            'Is this a multi-line comment? (two backticks)
            If LenB(Word) Then
                If AscChar = AscComment Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
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
            'Remember the first letter of the word for guiding tokenisation
            Let First = Char
        Else
            'A word is underway, check for end of line (note the `Else`, we don't _
             want to test line endings twice when the Word is currently empty)
            If AscChar = 10 Or AscChar = 13 Then
                'The end of a line ends a single-line comment:
                If IsCommentMulti = False Then
                    If IsComment Then Let IsComment = False: GoTo ReadWord
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
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If LenB(Word) = 0 Then GoTo NextChar
            GoTo WordBreak
            
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
        
        'Capture letters 2,3 & 4 (for high-speed register validation)
        If LetterIndex < 5 Then
            Let LetterIndex = LetterIndex + 1
            If LetterIndex = 2 Then Let Asc2 = LowCase(AscW(Char))
            If LetterIndex = 3 Then Let Asc3 = LowCase(AscW(Char))
            If LetterIndex = 4 Then Let Asc4 = LowCase(AscW(Char))
        End If
        
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
    'Tokenise the word:
    '----------------------------------------------------------------------------------
    'This section is written for pure speed, not for prettiness or flexibility
    Dim Register As Long
    
    'Get the ASCII code of the first letter - we do this here as it wastes time to _
     do it when we set `First` as that may happen multiple times when skipping _
     whitespace before the beginning of a word
    Let AscFirst = LowCase(AscW(First))
    
    'Check the first letter of the word to find the right kind of token:
    
    'Comma (very common!)
    If AscFirst = AscList Then
        'TODO
    
    'Hexadecimal number?
    ElseIf AscFirst = AscNumHex Then
        'TODO
        
    ElseIf AscFirst = 97 Then '--- `a` ---
        'Z80 instructions
        If StrComp(Word, "adc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_ADC): GoTo AddWord
        If StrComp(Word, "add") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_ADD): GoTo AddWord
        If StrComp(Word, "and") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_AND): GoTo AddWord
        'TODO: Keywords
        If StrComp(Word, "as") = 0 Then GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
         
    ElseIf AscFirst = 98 Then '--- `b` ---
        'Z80 instructions
        If StrComp(Word, "bit") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_BIT): GoTo AddWord
        'TODO: Keywords
        If StrComp(Word, "byte") = 0 Then GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 99 Then '--- `c` ---
        'Z80 instructions
        If StrComp(Word, "call") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CALL): GoTo AddWord
        If StrComp(Word, "ccf") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CCF): GoTo AddWord
        If StrComp(Word, "cp") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CP): GoTo AddWord
        If StrComp(Word, "cpd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPD): GoTo AddWord
        If StrComp(Word, "cpdr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPDR): GoTo AddWord
        If StrComp(Word, "cpi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPI): GoTo AddWord
        If StrComp(Word, "cpir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPIR): GoTo AddWord
        If StrComp(Word, "cpl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPL): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 100 Then '--- `d` ---
        'Z80 instructions
        If StrComp(Word, "daa") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DAA): GoTo AddWord
        If StrComp(Word, "dec") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DEC): GoTo AddWord
        If StrComp(Word, "di") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DI): GoTo AddWord
        If StrComp(Word, "djnz") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DJNZ): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 101 Then '--- `e` ---
        'Z80 instructions
        If StrComp(Word, "ei") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EI): GoTo AddWord
        If StrComp(Word, "ex") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EX): GoTo AddWord
        If StrComp(Word, "exx") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EXX): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
    
    ElseIf AscFirst = 104 Then '--- `h` ---
        'Z80 instructions
        If StrComp(Word, "halt") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_HALT): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 105 Then '--- `i` ---
        'Z80 instructions
        If StrComp(Word, "im") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IM): GoTo AddWord
        If StrComp(Word, "in") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IN): GoTo AddWord
        If StrComp(Word, "inc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INC): GoTo AddWord
        If StrComp(Word, "ind") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IND): GoTo AddWord
        If StrComp(Word, "indr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INDR): GoTo AddWord
        If StrComp(Word, "ini") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INI): GoTo AddWord
        If StrComp(Word, "inir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INIR): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 106 Then '--- `j` ---
        'Z80 instructions
        If StrComp(Word, "jp") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_JP): GoTo AddWord
        If StrComp(Word, "jr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_JR): GoTo AddWord
        
    ElseIf AscFirst = 108 Then '--- `l` ---
        'Z80 instructions
        If StrComp(Word, "ld") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LD): GoTo AddWord
        If StrComp(Word, "ldd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDD): GoTo AddWord
        If StrComp(Word, "lddr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDDR): GoTo AddWord
        If StrComp(Word, "ldi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDI): GoTo AddWord
        If StrComp(Word, "ldir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDIR): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
    
    ElseIf AscFirst = 110 Then '--- `n` ---
        'Z80 instructions
        If StrComp(Word, "neg") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_NEG): GoTo AddWord
        If StrComp(Word, "nop") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_NOP): GoTo AddWord
    
    ElseIf AscFirst = 111 Then '--- `o` ---
        'Z80 instructions
        If StrComp(Word, "or") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OR): GoTo AddWord
        If StrComp(Word, "out") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUT): GoTo AddWord
        If StrComp(Word, "outd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTD): GoTo AddWord
        If StrComp(Word, "outdr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTDR): GoTo AddWord
        If StrComp(Word, "outi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTI): GoTo AddWord
        If StrComp(Word, "outir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTIR): GoTo AddWord
        
    ElseIf AscFirst = 112 Then '--- `p` ---
        'Z80 instructions
        If StrComp(Word, "pop") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_POP): GoTo AddWord
        If StrComp(Word, "push") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_PUSH): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 114 Then '--- `r` ---
        'Z80 instructions
        If StrComp(Word, "res") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RES): GoTo AddWord
        If StrComp(Word, "ret") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RET): GoTo AddWord
        If StrComp(Word, "reti") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RETI): GoTo AddWord
        If StrComp(Word, "retn") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RETN): GoTo AddWord
        If StrComp(Word, "rla") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLA): GoTo AddWord
        If StrComp(Word, "rl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RL): GoTo AddWord
        If StrComp(Word, "rlca") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLCA): GoTo AddWord
        If StrComp(Word, "rlc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLC): GoTo AddWord
        If StrComp(Word, "rld") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLD): GoTo AddWord
        If StrComp(Word, "rra") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRA): GoTo AddWord
        If StrComp(Word, "rr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RR): GoTo AddWord
        If StrComp(Word, "rrca") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRCA): GoTo AddWord
        If StrComp(Word, "rrc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRC): GoTo AddWord
        If StrComp(Word, "rrd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRD): GoTo AddWord
        If StrComp(Word, "rst") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RST): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 115 Then '--- `s` ---
        'Z80 instructions
        If StrComp(Word, "sbc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SBC): GoTo AddWord
        If StrComp(Word, "scf") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SCF): GoTo AddWord
        If StrComp(Word, "set") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SET): GoTo AddWord
        If StrComp(Word, "sla") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SLA): GoTo AddWord
        If StrComp(Word, "sra") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SRA): GoTo AddWord
        If StrComp(Word, "sll") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SLL): GoTo AddWord
        If StrComp(Word, "srl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SRL): GoTo AddWord
        If StrComp(Word, "sub") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SUB): GoTo AddWord
        'Is it a register? (a falsey comparison is fastest)
        Let Register = ValidateRegisterAsc(AscFirst, Asc2, Asc3, Asc4)
        If Register = 0 Then
            'ERROR
        Else
            Call TokenStream.AddToken_Register(Register)
        End If
        
    ElseIf AscFirst = 120 Then '--- `x` ---
        'Z80 instructions
        If StrComp(Word, "xor") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_XOR): GoTo AddWord
        'TODO: The `x` operator
        
    'Block (open)
    ElseIf AscFirst = AscBlockOpen Then '--- `{` ---
        'TODO
        GoTo AddWord
    
    'Block (close)
    ElseIf AscFirst = AscBlockClose Then  '--- `}` ---
        'TODO
    
    'Parenthesis (open)
    ElseIf AscFirst = AscParenOpen Then '--- `(` ---
        'TODO
    
    'Parenthesis (close)
    ElseIf AscFirst = AscParenClose Then '--- `)` ---
        'TODO
    
    'A decimal number?
    ElseIf AscFirst >= 48 And AscFirst <= 57 Then
        'TODO:
        
    'Binary number?
    ElseIf AscFirst = AscNumBin Then '--- `%` ---
        'TODO
        
    'String (quote)
    ElseIf AscFirst = AscQuote Then '--- `"` ---
        'TODO
        
    ElseIf AscFirst = AscOpAdd Then '--- `+` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_ADD)
    ElseIf AscFirst = AscOpSub Then '--- `-` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_SUB)
    ElseIf AscFirst = AscOpMul Then '--- `*` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_MUL)
    ElseIf AscFirst = AscOpDiv Then '--- `/` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_DIV)
    ElseIf AscFirst = AscOpPow Then '--- `^` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_POW)
    ElseIf AscFirst = AscOpMod Then '--- `\` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_MOD)
    End If
    
    'Add the word to the array
    '----------------------------------------------------------------------------------
AddWord:
    If WordCount > WordLimit Then
        Let WordLimit = WordLimit + 50000
        ReDim Preserve Words(0 To WordLimit) As String
    End If
    
    Let Words(WordCount) = Word
    Let WordCount = WordCount + 1
    
'    Debug.Print Word
    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
EndOfFile:
    'Trim the buffer of words to the right size
    If WordLimit > WordCount Then ReDim Preserve Words(0 To WordCount) As String
    
    Let Text = vbNullString
    Debug.Print "- Words: " & Format$(WordCount, "#,#")
    Debug.Print "- Time: " & Format$(Timer - Start, "0.000") & "s"
End Function

'/// VALIDATION PROCEDURES ////////////////////////////////////////////////////////////

Public Function ValidateRegisterAsc( _
    ByRef Asc1 As Long, ByRef Asc2 As Long, ByRef Asc3 As Long, ByRef Asc4 As Long _
) As OZ80_TOKEN_REGISTER
    'Ignore anything beginning with a letter before A
    If Asc1 < 97 Then Exit Function
    'No register has a first letter higher than S
    If Asc1 > 115 Then Exit Function
    'If there's only 1 letter total, then it's obviously a 1-letter register
    If Asc2 = 0 Then GoTo OneLetter
    'If a semi-colon occurs at the second letter, then the register must be 1 letter
    If Asc2 = 59 Then GoTo OneLetter
    'Also an apostrophe (shadow register) is the same
    If Asc2 = 39 Then GoTo OneLetter
    
    'If there's only 2 letters in total, it has to be a two-letter register
    If Asc3 = 0 Then GoTo TwoLetters
    'If a semi-colon or apostrophe occurs at the third letter, _
     it could be a two-letter register
    If Asc3 = 59 Then GoTo TwoLetters
    If Asc3 = 39 Then GoTo TwoLetters
    
    Exit Function

OneLetter:
    If Asc1 = 97 Then Let ValidateRegisterAsc = TOKEN_REGISTER_A: Exit Function
    If Asc1 = 98 Then Let ValidateRegisterAsc = TOKEN_REGISTER_B: Exit Function
    If Asc1 = 99 Then Let ValidateRegisterAsc = TOKEN_REGISTER_C: Exit Function
    If Asc1 = 100 Then Let ValidateRegisterAsc = TOKEN_REGISTER_D: Exit Function
    If Asc1 = 101 Then Let ValidateRegisterAsc = TOKEN_REGISTER_E: Exit Function
    If Asc1 = 104 Then Let ValidateRegisterAsc = TOKEN_REGISTER_H: Exit Function
    If Asc1 = 105 Then Let ValidateRegisterAsc = TOKEN_REGISTER_I: Exit Function
    If Asc1 = 108 Then Let ValidateRegisterAsc = TOKEN_REGISTER_L: Exit Function
    If Asc1 = 114 Then Let ValidateRegisterAsc = TOKEN_REGISTER_R: Exit Function
    Exit Function

TwoLetters:
    If Asc1 = 97 Then       '"af"
        If Asc2 = 102 Then Let ValidateRegisterAsc = TOKEN_REGISTER_AF
    ElseIf Asc1 = 98 Then   '"bc"
        If Asc2 = 99 Then Let ValidateRegisterAsc = TOKEN_REGISTER_BC
    ElseIf Asc1 = 100 Then  '"de"
        If Asc2 = 101 Then Let ValidateRegisterAsc = TOKEN_REGISTER_DE
    ElseIf Asc1 = 104 Then  '"hl"
        If Asc2 = 108 Then Let ValidateRegisterAsc = TOKEN_REGISTER_HL
    ElseIf Asc1 = 105 Then  '"ix"
        If Asc2 = 120 Then Let ValidateRegisterAsc = TOKEN_REGISTER_IX: Exit Function
        If Asc2 = 121 Then Let ValidateRegisterAsc = TOKEN_REGISTER_IY
    ElseIf Asc1 = 112 Then  '"pc"
        If Asc2 = 99 Then Let ValidateRegisterAsc = TOKEN_REGISTER_PC
    ElseIf Asc1 = 115 Then  '"sp"
        If Asc2 = 112 Then Let ValidateRegisterAsc = TOKEN_REGISTER_PC
    End If
End Function
