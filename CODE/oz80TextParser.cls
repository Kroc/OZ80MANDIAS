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

Private Const SYNTAX_COMMENT = 96       '`
Private Const SYNTAX_QUOTE = 34         '"
Private Const SYNTAX_LABEL = 58         ':
Private Const SYNTAX_PROPERTY = 46      '.
Private Const SYNTAX_VARIABLE = 35      '#
Private Const SYNTAX_MACRO = 64         '@
Private Const SYNTAX_FUNCT = 63         '?
Private Const SYNTAX_NUMBER_HEX = 36    '$
Private Const SYNTAX_NUMBER_BIN = 37    '%
Private Const SYNTAX_LIST = 44          ',
Private Const SYNTAX_PAREN_OPEN = 40    '(
Private Const SYNTAX_PAREN_CLOSE = 41   ')
Private Const SYNTAX_BLOCK_OPEN = 123   '{
Private Const SYNTAX_BLOCK_CLOSE = 125  '}

Private Const SYNTAX_OPERATOR_ADD = 43  '+
Private Const SYNTAX_OPERATOR_SUB = 45  '-
Private Const SYNTAX_OPERATOR_MUL = 42  '*
Private Const SYNTAX_OPERATOR_DIV = 47  '/
Private Const SYNTAX_OPERATOR_POW = 94  '^
Private Const SYNTAX_OPERATOR_MOD = 92  '\

'--------------------------------------------------------------------------------------

Public Enum OZ80_ERROR
    NOERROR = 0
    INVALID_LABEL = 1
    INVALID_VARIABLE = 2
    UNKNOWN_KEYWORD = 3
End Enum

'/// VARIABLES ////////////////////////////////////////////////////////////////////////

Private WordCount As Long

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
    'Clean up the look up table
    Erase LowCase
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
    
    'Break the source code into words and convert them into machine tokens
    Tokenise (FilePath)
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Tokenise : Break a file into all its individual words and convert to machine tokens _
 ======================================================================================
Private Function Tokenise(ByRef FilePath As String) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices" or _
     even structured code. Gotos are fast, deal with it
     
    Dim Start As Single
    Let Start = Timer
    
    Dim IsQuote As Boolean
    Dim IsComment As Boolean
    Dim IsCommentMulti As Boolean
    
    'Read the whole file into a single string _
     ----------------------------------------------------------------------------------
    Dim Text As String
    Dim Data() As Byte
    Dim FileNumber As String
    Let FileNumber = FreeFile
    
    Debug.Print "Exploding file: " & FilePath
    Open FilePath For Input Access Read Lock Write As #FileNumber
    'Read the file as a binary string with one-byte per character (ASCII). _
     We first normalise to Unicode to deal with locale specifics and then downsample _
     into ASCII to save memory and make stepping through easier
    Let Data = StrConv( _
        StrConv(InputB(LOF(FileNumber), FileNumber), vbUnicode), _
        vbFromUnicode _
    )
    
    'Cache this for speed
    Dim TextLen As Long: Let TextLen = UBound(Data)
    Debug.Print "- Length: " & Format$(TextLen, "#,#") & " bytes"
    Close #FileNumber
    
    Dim TextPos As Long
    Let WordCount = 0
    
    'The first letter of a word will be remembered for guiding tokenisation. _
     The 2nd, 3rd and 4th letters are used in validating registers and number prefixes
    Dim LetterIndex As Long
    Dim Asc1 As Long, Asc2 As Long, Asc3 As Long, Asc4 As Long
    
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
    
    'Read a character:
    Dim AscChar As Long
    Let AscChar = Data(TextPos)
    Let TextPos = TextPos + 1
    
    'skip through comments...
    '----------------------------------------------------------------------------------
    If IsComment Then
        'For single-line comments:
        If IsCommentMulti = False Then
            'Is this a multi-line comment? (two backticks)
            If LenB(Word) Then
                If AscChar = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
            'The end of a line ends a single-line comment:
            If AscChar = 10 Or AscChar = 13 Then
                Let IsComment = False: GoTo ReadWord
            End If
        
        'For multi-line comments, check for backtick:
        ElseIf AscChar = SYNTAX_COMMENT Then
            'Check for a second back-tick following
            Let AscChar = Data(TextPos)
            Let TextPos = TextPos + 1
            If AscChar = SYNTAX_COMMENT Then
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
            Let Asc1 = LowCase(AscChar)
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
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = True: GoTo NextChar
        
        'Check for comment marker
        If AscChar = SYNTAX_COMMENT Then
            Let IsComment = True
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If LenB(Word) = 0 Then GoTo NextChar
            GoTo WordBreak
            
        End If
    
        'Symbols that can break a word in the middle without spaces to separate
        If AscChar = SYNTAX_LIST Then GoTo WordBreak
        If AscChar = SYNTAX_PAREN_OPEN Then GoTo WordBreak
        If AscChar = SYNTAX_PAREN_CLOSE Then GoTo WordBreak
        If AscChar = SYNTAX_BLOCK_OPEN Then GoTo WordBreak
        If AscChar = SYNTAX_BLOCK_CLOSE Then GoTo WordBreak
        
        If AscChar = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
        If AscChar = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
        If AscChar = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
        If AscChar = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
        If AscChar = SYNTAX_OPERATOR_POW Then GoTo WordBreak
        If AscChar = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
        
        'Capture letters 2,3 & 4 (for high-speed register validation)
        If LetterIndex < 5 Then
            Let LetterIndex = LetterIndex + 1
            If LetterIndex = 2 Then Let Asc2 = LowCase(AscChar)
            If LetterIndex = 3 Then Let Asc3 = LowCase(AscChar)
            If LetterIndex = 4 Then Let Asc4 = LowCase(AscChar)
        End If
        
    'Look for the end of a quote
    '----------------------------------------------------------------------------------
    Else
        'TODO: handle escape chars, skip tabs and line breaks
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = False: GoTo EndWord
        
    End If
    
    '----------------------------------------------------------------------------------
    
NextChar:
    Let Word = Word & ChrW$(AscChar)
    GoTo ReadChar

WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    If LenB(Word) Then
        Let TextPos = TextPos - 1
        GoTo EndWord2
    End If
    Let Word = Word & ChrW$(AscChar)
    
EndWord:
    If LenB(Word) = 0 Then GoTo EndOfFile
EndWord2:
    'Tokenise the word: _
     ----------------------------------------------------------------------------------
    'Check the first letter of the word to find the right kind of token:
    
    'Comma (very common!)
    If Asc1 = SYNTAX_LIST Then
        'TODO
        GoTo NextWord
    
    'Hexadecimal number?
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then
        'TODO: Tokenise number
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_LABEL Then
        'TODO: Tokenise
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_PROPERTY Then
        'TODO: Tokenise
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_VARIABLE Then
        'TODO: Tokenise
        GoTo NextWord
        
    ElseIf Asc1 = 97 Then '--- `a` ---
        'Z80 instructions
        If StrComp(Word, "adc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_ADC): GoTo NextWord
        If StrComp(Word, "add") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_ADD): GoTo NextWord
        If StrComp(Word, "and") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_AND): GoTo NextWord
        'Keywords
        If StrComp(Word, "as") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_AS): GoTo NextWord
        If StrComp(Word, "at") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_AT): GoTo NextWord
         
    ElseIf Asc1 = 98 Then '--- `b` ---
        'Z80 instructions
        If StrComp(Word, "bit") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_BIT): GoTo NextWord
        'Keywords
        If StrComp(Word, "bank") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_BANK): GoTo NextWord
        If StrComp(Word, "binary") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_BINARY): GoTo NextWord
        If StrComp(Word, "byte") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_BYTE): GoTo NextWord
        
    ElseIf Asc1 = 99 Then '--- `c` ---
        'Z80 instructions
        If StrComp(Word, "call") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CALL): GoTo NextWord
        If StrComp(Word, "ccf") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CCF): GoTo NextWord
        If StrComp(Word, "cp") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CP): GoTo NextWord
        If StrComp(Word, "cpd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPD): GoTo NextWord
        If StrComp(Word, "cpdr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPDR): GoTo NextWord
        If StrComp(Word, "cpi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPI): GoTo NextWord
        If StrComp(Word, "cpir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPIR): GoTo NextWord
        If StrComp(Word, "cpl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_CPL): GoTo NextWord
        
    ElseIf Asc1 = 100 Then '--- `d` ---
        'Z80 instructions
        If StrComp(Word, "daa") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DAA): GoTo NextWord
        If StrComp(Word, "dec") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DEC): GoTo NextWord
        If StrComp(Word, "di") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DI): GoTo NextWord
        If StrComp(Word, "djnz") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_DJNZ): GoTo NextWord
        'Keywords
        If Word = "DATA" Then Stop
        If StrComp(Word, "data") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_DATA): GoTo NextWord
        If StrComp(Word, "def") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_DEF): GoTo NextWord
        If StrComp(Word, "default") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_DEFAULT): GoTo NextWord
        
    ElseIf Asc1 = 101 Then '--- `e` ---
        'Z80 instructions
        If StrComp(Word, "ei") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EI): GoTo NextWord
        If StrComp(Word, "ex") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EX): GoTo NextWord
        If StrComp(Word, "exx") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_EXX): GoTo NextWord
    
    ElseIf Asc1 = 102 Then '--- `f` ---
        'Keywords
        If StrComp(Word, "fill") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_FILL): GoTo NextWord
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = 104 Then '--- `h` ---
        'Z80 instructions
        If StrComp(Word, "halt") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_HALT): GoTo NextWord
        
    ElseIf Asc1 = 105 Then '--- `i` ---
        'Z80 instructions
        If StrComp(Word, "im") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IM): GoTo NextWord
        If StrComp(Word, "in") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IN): GoTo NextWord
        If StrComp(Word, "inc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INC): GoTo NextWord
        If StrComp(Word, "ind") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_IND): GoTo NextWord
        If StrComp(Word, "indr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INDR): GoTo NextWord
        If StrComp(Word, "ini") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INI): GoTo NextWord
        If StrComp(Word, "inir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_INIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "include") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_INCLUDE): GoTo NextWord
        
    ElseIf Asc1 = 106 Then '--- `j` ---
        'Z80 instructions
        If StrComp(Word, "jp") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_JP): GoTo NextWord
        If StrComp(Word, "jr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_JR): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = 107 Then '--- `k` ---
        'Number suffixes K, KB & Kbit
        If Asc2 = 0 Then    'Just `K`
            Call TokenStream.AddToken_Prefix(TOKEN_PREFIX_K): GoTo NextWord
        'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = 98 Then
            'If only two letters, then it's `KB`
            If Asc3 = 0 Then Call TokenStream.AddToken_Prefix(TOKEN_PREFIX_KB): GoTo NextWord
            'verify ...`it`
            If Asc3 = 105 Then
                If Asc4 = 116 Then Call TokenStream.AddToken_Prefix(TOKEN_PREFIX_KBIT): GoTo NextWord
            End If
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = 108 Then '--- `l` ---
        'Z80 instructions
        If StrComp(Word, "ld") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LD): GoTo NextWord
        If StrComp(Word, "ldd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDD): GoTo NextWord
        If StrComp(Word, "lddr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDDR): GoTo NextWord
        If StrComp(Word, "ldi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDI): GoTo NextWord
        If StrComp(Word, "ldir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_LDIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "length") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_LENGTH): GoTo NextWord
    
    ElseIf Asc1 = 110 Then '--- `n` ---
        'Z80 instructions
        If StrComp(Word, "neg") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_NEG): GoTo NextWord
        If StrComp(Word, "nop") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_NOP): GoTo NextWord
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = 111 Then '--- `o` ---
        'Z80 instructions
        If StrComp(Word, "or") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OR): GoTo NextWord
        If StrComp(Word, "out") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUT): GoTo NextWord
        If StrComp(Word, "outd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTD): GoTo NextWord
        If StrComp(Word, "outdr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTDR): GoTo NextWord
        If StrComp(Word, "outi") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTI): GoTo NextWord
        If StrComp(Word, "outir") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_OUTIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "object") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_OBJECT): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = 112 Then '--- `p` ---
        'Z80 instructions
        If StrComp(Word, "pop") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_POP): GoTo NextWord
        If StrComp(Word, "push") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_PUSH): GoTo NextWord
        'Keywords
        If StrComp(Word, "params") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If StrComp(Word, "proc") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_PROC): GoTo NextWord
        
    ElseIf Asc1 = 114 Then '--- `r` ---
        'Z80 instructions
        If StrComp(Word, "res") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RES): GoTo NextWord
        If StrComp(Word, "ret") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RET): GoTo NextWord
        If StrComp(Word, "reti") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RETI): GoTo NextWord
        If StrComp(Word, "retn") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RETN): GoTo NextWord
        If StrComp(Word, "rla") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLA): GoTo NextWord
        If StrComp(Word, "rl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RL): GoTo NextWord
        If StrComp(Word, "rlca") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLCA): GoTo NextWord
        If StrComp(Word, "rlc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLC): GoTo NextWord
        If StrComp(Word, "rld") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RLD): GoTo NextWord
        If StrComp(Word, "rra") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRA): GoTo NextWord
        If StrComp(Word, "rr") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RR): GoTo NextWord
        If StrComp(Word, "rrca") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRCA): GoTo NextWord
        If StrComp(Word, "rrc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRC): GoTo NextWord
        If StrComp(Word, "rrd") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RRD): GoTo NextWord
        If StrComp(Word, "rst") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If StrComp(Word, "return") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = 115 Then '--- `s` ---
        'Z80 instructions
        If StrComp(Word, "sbc") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SBC): GoTo NextWord
        If StrComp(Word, "scf") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SCF): GoTo NextWord
        If StrComp(Word, "set") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SET): GoTo NextWord
        If StrComp(Word, "sla") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SLA): GoTo NextWord
        If StrComp(Word, "sra") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SRA): GoTo NextWord
        If StrComp(Word, "sll") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SLL): GoTo NextWord
        If StrComp(Word, "srl") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SRL): GoTo NextWord
        If StrComp(Word, "sub") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If StrComp(Word, "slot") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_SLOT): GoTo NextWord
        If StrComp(Word, "start") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_START): GoTo NextWord
        If StrComp(Word, "stop") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_STOP): GoTo NextWord
        If StrComp(Word, "struct") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_STRUCT): GoTo NextWord
        
    ElseIf Asc1 = 119 Then '--- `w` ---
        'Keywords
        If StrComp(Word, "word") = 0 Then Call TokenStream.AddToken_Keyword(TOKEN_KEYWORD_WORD): GoTo NextWord
        'There are no `w` registers
        GoTo WordError
        
    ElseIf Asc1 = 120 Then '--- `x` ---
        'Z80 instructions
        If StrComp(Word, "xor") = 0 Then Call TokenStream.AddToken_Z80(TOKEN_Z80_XOR): GoTo NextWord
        'TODO: The `x` operator
        'There are no `x` registers
        GoTo WordError
        
    'Block (open)
    ElseIf Asc1 = SYNTAX_BLOCK_OPEN Then '--- `{` ---
        'TODO: Tokenise
        GoTo NextWord
    
    'Block (close)
    ElseIf Asc1 = SYNTAX_BLOCK_CLOSE Then  '--- `}` ---
        'TODO: Tokenise
        GoTo NextWord
    
    'Parenthesis (open)
    ElseIf Asc1 = SYNTAX_PAREN_OPEN Then '--- `(` ---
        'TODO: Tokenise
        GoTo NextWord
    
    'Parenthesis (close)
    ElseIf Asc1 = SYNTAX_PAREN_CLOSE Then '--- `)` ---
        'TODO: Tokenise
        GoTo NextWord
    
    'A decimal number?
    ElseIf Asc1 >= 48 And Asc1 <= 57 Then
        'TODO: Tokenise number
        GoTo NextWord
        
    'Binary number?
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then '--- `%` ---
        'TODO: Tokenise number
        GoTo NextWord
        
    'String (quote)
    ElseIf Asc1 = SYNTAX_QUOTE Then '--- `"` ---
        'TODO: Tokenise
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_OPERATOR_ADD Then '--- `+` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_ADD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '--- `-` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_SUB): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '--- `*` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_MUL): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '--- `/` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_DIV): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '--- `^` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_POW): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '--- `\` ---
        Call TokenStream.AddToken_Operator(TOKEN_OPERATOR_MOD): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < 97 Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > 115 Then GoTo WordError
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
    
    GoTo WordError

OneLetter:
    If Asc1 = 97 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_A): GoTo NextWord
    If Asc1 = 98 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_B): GoTo NextWord
    If Asc1 = 99 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_C): GoTo NextWord
    If Asc1 = 100 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_D): GoTo NextWord
    If Asc1 = 101 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_E): GoTo NextWord
    If Asc1 = 104 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_H): GoTo NextWord
    If Asc1 = 105 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_I): GoTo NextWord
    If Asc1 = 108 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_L): GoTo NextWord
    If Asc1 = 114 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_R): GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = 97 Then       '"af"
        If Asc2 = 102 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_AF): GoTo NextWord
    ElseIf Asc1 = 98 Then   '"bc"
        If Asc2 = 99 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_BC): GoTo NextWord
    ElseIf Asc1 = 100 Then  '"de"
        If Asc2 = 101 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_DE): GoTo NextWord
    ElseIf Asc1 = 104 Then  '"hl"
        If Asc2 = 108 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_HL): GoTo NextWord
    ElseIf Asc1 = 105 Then  '"ix"
        If Asc2 = 120 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_IX): GoTo NextWord
        If Asc2 = 121 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_IY): GoTo NextWord
    ElseIf Asc1 = 112 Then  '"pc"
        If Asc2 = 99 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_PC): GoTo NextWord
    ElseIf Asc1 = 115 Then  '"sp"
        If Asc2 = 112 Then Call TokenStream.AddToken_Register(TOKEN_REGISTER_SP): GoTo NextWord
    End If
    
    '----------------------------------------------------------------------------------
WordError:
    'TODO: Error here - unrecognised word
'    Debug.Print "! Unrecognised word: " & Word

    '----------------------------------------------------------------------------------
NextWord:
    Let WordCount = WordCount + 1
    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
EndOfFile:
    Erase Data
    Debug.Print "- Words: " & Format$(WordCount, "#,#")
    Debug.Print "- Time: " & Format$(Timer - Start, "0.000") & "s"
End Function
