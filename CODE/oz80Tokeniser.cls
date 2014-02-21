VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80Tokeniser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS:: oz80TokenStream

'A Token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. This class handles _
 parsing the source code text into tokens; the assembler only works with tokens

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

'These define the various punctiation marks (in ASCII codes) for the language syntax
Private Const SYNTAX_COMMENT = 96       '`
Private Const SYNTAX_HINT1 = 59         ';
Private Const SYNTAX_HINT2 = 59         ''
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

'This makes life a whole lot easier
Private Const A = 97
Private Const B = 98
Private Const C = 99
Private Const D = 100
Private Const E = 101
Private Const F = 102
Private Const G = 103
Private Const H = 104
Private Const I = 105
Private Const J = 106
Private Const K = 107
Private Const L = 108
Private Const M = 109
Private Const N = 110
Private Const O = 111
Private Const P = 112
Private Const Q = 113
Private Const R = 114
Private Const S = 115
Private Const T = 116
Private Const U = 117
Private Const V = 118
Private Const W = 119
Private Const X = 120
Private Const Y = 121
Private Const Z = 122

Private Const aSPC = 32
Private Const aTAB = 9
Private Const aCR = 10
Private Const aLF = 13

'--------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN
    'This token specifies that the data field is a Z80 mnemonic (`OZ80_MNEMONICS`).
     'Note that this is only a mnemonic token, not the Z80 opcode (handled by the
     'assembler) or parameters (handled by the next tokens in the stream)
    TOKEN_Z80 = &H10
    
    'Specifies a Z80 register
    TOKEN_REGISTER = &H20
    
    'OZ80's keywords
    TOKEN_KEYWORD = &H30
    
    'The parser automatically converts hexadecimal/binary numbers, so we only store
     'a 32-bit long (data field) in the token stream
    TOKEN_NUMBER = &H40
    TOKEN_PREFIX = &H41                 'Number prefixes ("K", "KB" & "Kbit")
    
    'Specifies an operator
    TOKEN_OPERATOR = &H50
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_PARENOPEN = &H60
    TOKEN_PARENCLOSE = &H61
    TOKEN_BLOCKOPEN = &H62
    TOKEN_BLOCKCLOSE = &H63
    
    'A list is a series of 1 or more expressions separated by commas,
     'i.e.parameter lists
    TOKEN_LIST = &H70
    TOKEN_NEXT = &H71                   'Sandwiched between list items, i.e. ","
    
    'Debug tokens:
    'These allow us to keep track of where the token originated in the source file(s)
     'so that errors in parsing the tokens can print out friendly error messages
    TOKEN_FILE = &HF0                   'specify a change in file name
    TOKEN_LINE = &HF1                   'specify a change in line number
    TOKEN_COL = &HF2                    'specify a change in column number
End Enum

'Z80 Assembly Mnemonics: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_Z80
    'These are just the mnemonic tokens -- the assembler itself checks the
     'parameters and determines which opcode should be used
    TOKEN_Z80_ADC = &H1                 'Add with Carry
    TOKEN_Z80_ADD = &H2                 'Add
    TOKEN_Z80_AND = &H3                 'Bitwise AND
    TOKEN_Z80_BIT = &H4                 'Bit test
    TOKEN_Z80_CALL = &H5                'Call routine
    TOKEN_Z80_CCF = &H6                 'Clear Carry Flag
    TOKEN_Z80_CP = &H7                  'Compare
    TOKEN_Z80_CPD = &H8                 'Compare and Decrement
    TOKEN_Z80_CPDR = &H9                'Compare, Decrement and Repeat
    TOKEN_Z80_CPI = &HA                 'Compare and Increment
    TOKEN_Z80_CPIR = &HB                'Compare, Increment and Repeat
    TOKEN_Z80_CPL = &HC                 'Complement (bitwise NOT)
    TOKEN_Z80_DAA = &HD                 'Decimal Adjust Accumulator
    TOKEN_Z80_DEC = &HE                 'Decrement
    TOKEN_Z80_DI = &HF                  'Disable Interrupts
    TOKEN_Z80_DJNZ = &H10               'Decrement and Jump if Not Zero
    TOKEN_Z80_EI = &H11                 'Enable Inettupts
    TOKEN_Z80_EX = &H12                 'Exchange
    TOKEN_Z80_EXX = &H13                'Exchange shadow registers
    TOKEN_Z80_HALT = &H14               'Stop CPU (wait for interrupt)
    TOKEN_Z80_IM = &H15                 'Interrupt Mode
    TOKEN_Z80_IN = &H16                 'Input from port
    TOKEN_Z80_INC = &H17                'Increment
    TOKEN_Z80_IND = &H18                'Input and Decrement
    TOKEN_Z80_INDR = &H19               'Input, Decrement and Repeat
    TOKEN_Z80_INI = &H1A                'Input and Increment
    TOKEN_Z80_INIR = &H1B               'Input, Increment and Repeat
    TOKEN_Z80_JP = &H1C                 'Jump
    TOKEN_Z80_JR = &H1D                 'Jump Relative
    TOKEN_Z80_LD = &H1E                 'Load
    TOKEN_Z80_LDD = &H1F                'Load and Decrement
    TOKEN_Z80_LDDR = &H20               'Load, Decrement and Repeat
    TOKEN_Z80_LDI = &H21                'Load and Increment
    TOKEN_Z80_LDIR = &H22               'Load, Increment and Repeat
    TOKEN_Z80_NEG = &H23                'Negate (flip the sign)
    TOKEN_Z80_NOP = &H24                'No Operation (do nothing)
    TOKEN_Z80_OR = &H25                 'Bitwise OR
    TOKEN_Z80_OUT = &H26                'Output to port
    TOKEN_Z80_OUTD = &H27               'Output and Decrement
    TOKEN_Z80_OUTDR = &H28              'Output, Decrement and Repeat
    TOKEN_Z80_OUTI = &H29               'Output and Increment
    TOKEN_Z80_OUTIR = &H2A              'Output, Increment and Repeat
    TOKEN_Z80_POP = &H2B                'Pull from stack
    TOKEN_Z80_PUSH = &H2C               'Push onto stack
    TOKEN_Z80_RES = &H2D                'Reset bit
    TOKEN_Z80_RET = &H2E                'Return from routine
    TOKEN_Z80_RETI = &H2F               'Return from Interrupt
    TOKEN_Z80_RETN = &H30               'Return from NMI
    TOKEN_Z80_RLA = &H31                'Rotate Left (Accumulator)
    TOKEN_Z80_RL = &H32                 'Rotate Left
    TOKEN_Z80_RLCA = &H33               'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLC = &H34                'Rotate Left Circular
    TOKEN_Z80_RLD = &H35                'Rotate Left 4-bits
    TOKEN_Z80_RRA = &H36                'Rotate Right (Accumulator)
    TOKEN_Z80_RR = &H37                 'Rotate Right
    TOKEN_Z80_RRCA = &H38               'Rotate Right Circular (Accumulator)
    TOKEN_Z80_RRC = &H39                'Rotate Right Circular
    TOKEN_Z80_RRD = &H3A                'Rotate Right 4-bits
    TOKEN_Z80_RST = &H3B                '"Restart" -- Call a page 0 routine
    TOKEN_Z80_SBC = &H3C                'Subtract with Carry
    TOKEN_Z80_SCF = &H3D                'Set Carry Flag
    TOKEN_Z80_SET = &H3E                'Set bit
    TOKEN_Z80_SLA = &H3F                'Shift Left Arithmetic
    TOKEN_Z80_SRA = &H40                'Shift Right Arithmetic
    TOKEN_Z80_SLL = &H41                'Shift Left Logical
    TOKEN_Z80_SRL = &H42                'Shift Right Logical
    TOKEN_Z80_SUB = &H43                'Subtract
    TOKEN_Z80_XOR = &H44                'Bitwise XOR
End Enum

'Z80 Registers: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_REGISTER
    TOKEN_REGISTER_A = 1                'Accumulator
    TOKEN_REGISTER_B = 2
    TOKEN_REGISTER_C = 3
    TOKEN_REGISTER_D = 4
    TOKEN_REGISTER_E = 5
    TOKEN_REGISTER_F = 6                'Flags register
    TOKEN_REGISTER_H = 7
    TOKEN_REGISTER_I = 8                'Interrupt - not to be confused with IX & IY
    TOKEN_REGISTER_L = 9
    TOKEN_REGISTER_R = 10               'Refresh register (pseudo-random)

    TOKEN_REGISTER_AF = 11
    TOKEN_REGISTER_BC = 12
    TOKEN_REGISTER_DE = 13
    TOKEN_REGISTER_HL = 14
    TOKEN_REGISTER_IX = 17
    TOKEN_REGISTER_IY = 20
    
    'Undocumented Z80 instructions can access the 8-bit halves of IX & IY
    TOKEN_REGISTER_IXL = 15
    TOKEN_REGISTER_IXH = 16
    TOKEN_REGISTER_IYL = 18
    TOKEN_REGISTER_IYH = 19

    TOKEN_REGISTER_SP = 21              'Stack pointer
    TOKEN_REGISTER_PC = 22              'Program counter
End Enum

'Keywords: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_KEYWORD
    TOKEN_KEYWORD_AT = 1
    TOKEN_KEYWORD_AS = 2
    TOKEN_KEYWORD_BANK = 3
    TOKEN_KEYWORD_BINARY = 4
    TOKEN_KEYWORD_BYTE = 5
    TOKEN_KEYWORD_DATA = 6
    TOKEN_KEYWORD_DEF = 7
    TOKEN_KEYWORD_DEFAULT = 8
    TOKEN_KEYWORD_FILL = 9
    TOKEN_KEYWORD_INCLUDE = 10
    TOKEN_KEYWORD_LENGTH = 11
    TOKEN_KEYWORD_OBJECT = 12
    TOKEN_KEYWORD_PARAMS = 13
    TOKEN_KEYWORD_PROC = 14
    TOKEN_KEYWORD_RETURN = 15
    TOKEN_KEYWORD_SLOT = 16
    TOKEN_KEYWORD_START = 17
    TOKEN_KEYWORD_STOP = 18
    TOKEN_KEYWORD_STRUCT = 19
    TOKEN_KEYWORD_WORD = 20
End Enum

'Number Prefixes: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_PREFIX
    TOKEN_PREFIX_K = 1                  'x1000
    TOKEN_PREFIX_KB = 2                 'x1024
    TOKEN_PREFIX_KBIT = 3               'x128 (1024 bits)
End Enum

'Operators: _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_OPERATOR
    TOKEN_OPERATOR_ADD = 1              'Add "+"
    TOKEN_OPERATOR_SUB = 2              'Subtract "-"
    TOKEN_OPERATOR_MUL = 3              'Multiply "*"
    TOKEN_OPERATOR_DIV = 4              'Divide "/"
    TOKEN_OPERATOR_POW = 5              'Power "^"
    TOKEN_OPERATOR_MOD = 6              'Modulus "\"
    TOKEN_OPERATOR_REP = 7              'Repeat "x"
End Enum

'Grouping: (i.e. parenthesis, braces) _
 --------------------------------------------------------------------------------------
Public Enum OZ80_TOKEN_GROUP
    TOKEN_GROUP_PARENOPEN = OZ80_TOKEN.TOKEN_PARENOPEN
    TOKEN_GROUP_PARENCLOSE = OZ80_TOKEN.TOKEN_PARENCLOSE
    TOKEN_GROUP_BLOCKOPEN = OZ80_TOKEN.TOKEN_BLOCKOPEN
    TOKEN_GROUP_BLOCKCLOSE = OZ80_TOKEN.TOKEN_BLOCKCLOSE
End Enum

'/// DEFS /////////////////////////////////////////////////////////////////////////////

Private Tokens() As Byte                'The token stream will live here
Private TokensLimit As Long             'The length of the above array
Private TokensLength As Long            'The current amount of space used in the array
Private TokensCursor As Long            'Where the current insertion point is
Private TokensCount As Long             'The number of tokens, exluding the data parts

'A look up table of ASCII codes from upper-case to lower-case
Private LowCase(0 To 255) As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Prepare a large buffer (512K) to accept the token stream _
     (using `ReDim` for every token added is *incredibly* slow)
    'If the data overflows this buffer, we'll increase it by _
     another chunk there and then
    Let TokensLimit = 524287
    ReDim Tokens(0 To TokensLimit) As Byte
    
    'Initialise the lookup table of upper-case to lower-case values
    Dim I As Long
    For I = 0 To 255
        If I >= 65 And I <= 90 Then
            Let LowCase(I) = I And 32
        Else
            Let LowCase(I) = I
        End If
    Next I
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Clean up the look up table
    Erase LowCase
    'Empty the token stream
    Erase Tokens
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddToken_Z80 : Add a Z80 instruction to the stream _
 ======================================================================================
Public Sub AddToken_Z80(ByRef Token As OZ80_TOKEN_Z80)
    Call Add(OZ80_TOKEN.TOKEN_Z80)      'Add a byte to say this is a Z80 instructions
    Call Add(Token)                     'And add the Z80 instruction token
End Sub

'AddToken_Register : Add a Z80 register to the stream _
 ======================================================================================
Public Sub AddToken_Register(ByRef Token As OZ80_TOKEN_REGISTER)
    Call Add(OZ80_TOKEN.TOKEN_REGISTER) 'Add a byte to say this is a Z80 register
    Call Add(Token)                     'And add the register token
End Sub

'AddToken_Keyword : Add an OZ80 keyword to the stream _
 ======================================================================================
Public Sub AddToken_Keyword(ByRef Token As OZ80_TOKEN_KEYWORD)
    Call Add(OZ80_TOKEN.TOKEN_KEYWORD)  'Add a byte to say this is a keyword
    Call Add(Token)                     'And add the keyword token
End Sub

'AddToken_Prefix : Add a number prefix to the stream _
 ======================================================================================
Public Sub AddToken_Prefix(ByRef Token As OZ80_TOKEN_PREFIX)
    Call Add(OZ80_TOKEN.TOKEN_PREFIX)   'Add a byte to say this is a number prefix
    Call Add(Token)                     'And add the prefix token
End Sub

'AddToken_Operator : Add an operator to the stream _
 ======================================================================================
Public Sub AddToken_Operator(ByRef Token As OZ80_TOKEN_OPERATOR)
    Call Add(OZ80_TOKEN.TOKEN_OPERATOR) 'Add a byte to say this is an OZ80 operator
    Call Add(Token)                     'And add the operator token
End Sub

'AddToken_Group : Grouping, i.e. parenthesis, braces _
 ======================================================================================
Public Sub AddToken_Group(ByRef Token As OZ80_TOKEN_GROUP)
    'These tokens don't require any data, just a single token is needed
    Call Add(Token)
End Sub

'Tokenise : Break a file into all its individual words and convert to machine tokens _
 ======================================================================================
Public Sub Tokenise(ByRef FilePath As String)
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
    
    Debug.Print "Parsing file: " & FilePath
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
    Let TextPos = -1
    
    'The first letter of a word will be remembered for guiding tokenisation. _
     The 2nd, 3rd and 4th letters are used in validating registers and number prefixes
    Dim LetterIndex As Long
    Dim Asc1 As Long, Asc2 As Long, Asc3 As Long, Asc4 As Long
    
    Dim WordCount As Long

    '----------------------------------------------------------------------------------
ReadWord:
    'The word as we build it
    Dim Word As String: Let Word = vbNullString
    'Reset the counter and captures for the 2nd, 3rd and 4th characters of the word
    Let LetterIndex = 0
    Let Asc2 = 0: Let Asc3 = 0: Let Asc4 = 0

SkipChar:
    Let TextPos = TextPos + 1
ReadChar:
    'If the file ends, treat it as a remaining end of line
    If TextPos >= TextLen Then GoTo EndWord
    
    'Read a character:
    Dim AscChar As Long
    Let AscChar = Data(TextPos)
    
    'Comments:
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
            If AscChar = aCR Or AscChar = aLF Then
                Let IsComment = False: GoTo ReadWord
            End If
        
        'For multi-line comments, check for backtick:
        ElseIf AscChar = SYNTAX_COMMENT Then
            'Check for a second back-tick following
            Let TextPos = TextPos + 1
            Let AscChar = Data(TextPos)
            If AscChar = SYNTAX_COMMENT Then
                Let IsComment = False: Let IsCommentMulti = False
                GoTo ReadWord
            End If
        End If
        'Skip to the next character
        GoTo SkipChar
    
    'Parsable Text:
    '----------------------------------------------------------------------------------
    '(false checks are slightly faster)
    ElseIf IsQuote = False Then
        'Special handling for the beginning of a word
        If LenB(Word) = 0 Then
            If AscChar = aSPC Then GoTo SkipChar
            If AscChar = aTAB Then GoTo SkipChar
            If AscChar = aCR Then GoTo SkipChar
            If AscChar = aLF Then GoTo SkipChar
            'Remember the first letter of the word for guiding tokenisation
            Let Asc1 = LowCase(AscChar)
            
            Dim WordBegin As Long
            Let WordBegin = TextPos
        Else
            'A word is underway, check for end of line (note the `Else`, we don't _
             want to test line endings twice when the Word is currently empty)
            If AscChar = aCR Or AscChar = aLF Then
                'The end of a line ends a single-line comment:
                If IsCommentMulti = False Then
                    If IsComment Then Let IsComment = False: GoTo ReadWord
                End If
                'End of the line ends the word
                GoTo EndWord2
            End If
            
            'End the word on a space instead of at the end of the line
            If AscChar = aSPC Then GoTo EndWord
            If AscChar = aTAB Then GoTo EndWord
        End If
        
        'Is this a quote? (in which case, ignore wordbreaks until quote end). _
         the opening quote is included so that we can recognise it for tokenising
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
        
    'Quotes:
    '----------------------------------------------------------------------------------
    Else
        'TODO: skip tabs and line breaks
        'Skip over the closing quote so that it doesn't get tokenised
        If AscChar = SYNTAX_QUOTE Then Let IsQuote = False: GoTo EndWord2
        
    End If
    
    '----------------------------------------------------------------------------------
    
NextChar:
    Let Word = Word & ChrW$(AscChar)
    GoTo SkipChar

WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    If LenB(Word) Then GoTo EndWord2
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
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        'TODO: Tokenise
        GoTo NextWord
    
    ElseIf Asc1 = SYNTAX_VARIABLE Then
        'TODO: Tokenise
        GoTo NextWord
        
    ElseIf Asc1 = A Then '--- `a` ---
        'Z80 instructions
        If StrComp(Word, "adc") = 0 Then Call AddToken_Z80(TOKEN_Z80_ADC): GoTo NextWord
        If StrComp(Word, "add") = 0 Then Call AddToken_Z80(TOKEN_Z80_ADD): GoTo NextWord
        If StrComp(Word, "and") = 0 Then Call AddToken_Z80(TOKEN_Z80_AND): GoTo NextWord
        'Keywords
        If StrComp(Word, "as") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_AS): GoTo NextWord
        If StrComp(Word, "at") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_AT): GoTo NextWord
         
    ElseIf Asc1 = B Then '--- `b` ---
        'Z80 instructions
        If StrComp(Word, "bit") = 0 Then Call AddToken_Z80(TOKEN_Z80_BIT): GoTo NextWord
        'Keywords
        If StrComp(Word, "bank") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_BANK): GoTo NextWord
        If StrComp(Word, "binary") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_BINARY): GoTo NextWord
        If StrComp(Word, "byte") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_BYTE): GoTo NextWord
        
    ElseIf Asc1 = C Then '--- `c` ---
        'Z80 instructions
        If StrComp(Word, "call") = 0 Then Call AddToken_Z80(TOKEN_Z80_CALL): GoTo NextWord
        If StrComp(Word, "ccf") = 0 Then Call AddToken_Z80(TOKEN_Z80_CCF): GoTo NextWord
        If StrComp(Word, "cp") = 0 Then Call AddToken_Z80(TOKEN_Z80_CP): GoTo NextWord
        If StrComp(Word, "cpd") = 0 Then Call AddToken_Z80(TOKEN_Z80_CPD): GoTo NextWord
        If StrComp(Word, "cpdr") = 0 Then Call AddToken_Z80(TOKEN_Z80_CPDR): GoTo NextWord
        If StrComp(Word, "cpi") = 0 Then Call AddToken_Z80(TOKEN_Z80_CPI): GoTo NextWord
        If StrComp(Word, "cpir") = 0 Then Call AddToken_Z80(TOKEN_Z80_CPIR): GoTo NextWord
        If StrComp(Word, "cpl") = 0 Then Call AddToken_Z80(TOKEN_Z80_CPL): GoTo NextWord
        
    ElseIf Asc1 = D Then '--- `d` ---
        'Z80 instructions
        If StrComp(Word, "daa") = 0 Then Call AddToken_Z80(TOKEN_Z80_DAA): GoTo NextWord
        If StrComp(Word, "dec") = 0 Then Call AddToken_Z80(TOKEN_Z80_DEC): GoTo NextWord
        If StrComp(Word, "di") = 0 Then Call AddToken_Z80(TOKEN_Z80_DI): GoTo NextWord
        If StrComp(Word, "djnz") = 0 Then Call AddToken_Z80(TOKEN_Z80_DJNZ): GoTo NextWord
        'Keywords
        If Word = "DATA" Then Stop
        If StrComp(Word, "data") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_DATA): GoTo NextWord
        If StrComp(Word, "def") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_DEF): GoTo NextWord
        If StrComp(Word, "default") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_DEFAULT): GoTo NextWord
        
    ElseIf Asc1 = E Then '--- `e` ---
        'Z80 instructions
        If StrComp(Word, "ei") = 0 Then Call AddToken_Z80(TOKEN_Z80_EI): GoTo NextWord
        If StrComp(Word, "ex") = 0 Then Call AddToken_Z80(TOKEN_Z80_EX): GoTo NextWord
        If StrComp(Word, "exx") = 0 Then Call AddToken_Z80(TOKEN_Z80_EXX): GoTo NextWord
    
    ElseIf Asc1 = F Then '--- `f` ---
        'Keywords
        If StrComp(Word, "fill") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_FILL): GoTo NextWord
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = H Then '--- `h` ---
        'Z80 instructions
        If StrComp(Word, "halt") = 0 Then Call AddToken_Z80(TOKEN_Z80_HALT): GoTo NextWord
        
    ElseIf Asc1 = I Then '--- `i` ---
        'Z80 instructions
        If StrComp(Word, "im") = 0 Then Call AddToken_Z80(TOKEN_Z80_IM): GoTo NextWord
        If StrComp(Word, "in") = 0 Then Call AddToken_Z80(TOKEN_Z80_IN): GoTo NextWord
        If StrComp(Word, "inc") = 0 Then Call AddToken_Z80(TOKEN_Z80_INC): GoTo NextWord
        If StrComp(Word, "ind") = 0 Then Call AddToken_Z80(TOKEN_Z80_IND): GoTo NextWord
        If StrComp(Word, "indr") = 0 Then Call AddToken_Z80(TOKEN_Z80_INDR): GoTo NextWord
        If StrComp(Word, "ini") = 0 Then Call AddToken_Z80(TOKEN_Z80_INI): GoTo NextWord
        If StrComp(Word, "inir") = 0 Then Call AddToken_Z80(TOKEN_Z80_INIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "include") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_INCLUDE): GoTo NextWord
        
    ElseIf Asc1 = J Then '--- `j` ---
        'Z80 instructions
        If StrComp(Word, "jp") = 0 Then Call AddToken_Z80(TOKEN_Z80_JP): GoTo NextWord
        If StrComp(Word, "jr") = 0 Then Call AddToken_Z80(TOKEN_Z80_JR): GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = K Then '--- `k` ---
        'Number suffixes K, KB & Kbit
        If Asc2 = 0 Then    'Just `K`
            Call AddToken_Prefix(TOKEN_PREFIX_K): GoTo NextWord
        'B must be next for `KB` and `Kbit`
        ElseIf Asc2 = B Then
            'If only two letters, then it's `KB`
            If Asc3 = 0 Then Call AddToken_Prefix(TOKEN_PREFIX_KB): GoTo NextWord
            'verify ...`Kbit`, we don't have Asc5 so avoid things like `Kbitty`
            If StrComp(Word, "kbit") = 0 Then Call AddToken_Prefix(TOKEN_PREFIX_KBIT): GoTo NextWord
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = L Then '--- `l` ---
        'Z80 instructions
        If StrComp(Word, "ld") = 0 Then Call AddToken_Z80(TOKEN_Z80_LD): GoTo NextWord
        If StrComp(Word, "ldd") = 0 Then Call AddToken_Z80(TOKEN_Z80_LDD): GoTo NextWord
        If StrComp(Word, "lddr") = 0 Then Call AddToken_Z80(TOKEN_Z80_LDDR): GoTo NextWord
        If StrComp(Word, "ldi") = 0 Then Call AddToken_Z80(TOKEN_Z80_LDI): GoTo NextWord
        If StrComp(Word, "ldir") = 0 Then Call AddToken_Z80(TOKEN_Z80_LDIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "length") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_LENGTH): GoTo NextWord
    
    ElseIf Asc1 = N Then '--- `n` ---
        'Z80 instructions
        If StrComp(Word, "neg") = 0 Then Call AddToken_Z80(TOKEN_Z80_NEG): GoTo NextWord
        If StrComp(Word, "nop") = 0 Then Call AddToken_Z80(TOKEN_Z80_NOP): GoTo NextWord
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = O Then '--- `o` ---
        'Z80 instructions
        If StrComp(Word, "or") = 0 Then Call AddToken_Z80(TOKEN_Z80_OR): GoTo NextWord
        If StrComp(Word, "out") = 0 Then Call AddToken_Z80(TOKEN_Z80_OUT): GoTo NextWord
        If StrComp(Word, "outd") = 0 Then Call AddToken_Z80(TOKEN_Z80_OUTD): GoTo NextWord
        If StrComp(Word, "outdr") = 0 Then Call AddToken_Z80(TOKEN_Z80_OUTDR): GoTo NextWord
        If StrComp(Word, "outi") = 0 Then Call AddToken_Z80(TOKEN_Z80_OUTI): GoTo NextWord
        If StrComp(Word, "outir") = 0 Then Call AddToken_Z80(TOKEN_Z80_OUTIR): GoTo NextWord
        'Keywords
        If StrComp(Word, "object") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_OBJECT): GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = P Then '--- `p` ---
        'Z80 instructions
        If StrComp(Word, "pop") = 0 Then Call AddToken_Z80(TOKEN_Z80_POP): GoTo NextWord
        If StrComp(Word, "push") = 0 Then Call AddToken_Z80(TOKEN_Z80_PUSH): GoTo NextWord
        'Keywords
        If StrComp(Word, "params") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_PARAMS): GoTo NextWord
        If StrComp(Word, "proc") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_PROC): GoTo NextWord
        
    ElseIf Asc1 = R Then '--- `r` ---
        'Z80 instructions
        If StrComp(Word, "res") = 0 Then Call AddToken_Z80(TOKEN_Z80_RES): GoTo NextWord
        If StrComp(Word, "ret") = 0 Then Call AddToken_Z80(TOKEN_Z80_RET): GoTo NextWord
        If StrComp(Word, "reti") = 0 Then Call AddToken_Z80(TOKEN_Z80_RETI): GoTo NextWord
        If StrComp(Word, "retn") = 0 Then Call AddToken_Z80(TOKEN_Z80_RETN): GoTo NextWord
        If StrComp(Word, "rla") = 0 Then Call AddToken_Z80(TOKEN_Z80_RLA): GoTo NextWord
        If StrComp(Word, "rl") = 0 Then Call AddToken_Z80(TOKEN_Z80_RL): GoTo NextWord
        If StrComp(Word, "rlca") = 0 Then Call AddToken_Z80(TOKEN_Z80_RLCA): GoTo NextWord
        If StrComp(Word, "rlc") = 0 Then Call AddToken_Z80(TOKEN_Z80_RLC): GoTo NextWord
        If StrComp(Word, "rld") = 0 Then Call AddToken_Z80(TOKEN_Z80_RLD): GoTo NextWord
        If StrComp(Word, "rra") = 0 Then Call AddToken_Z80(TOKEN_Z80_RRA): GoTo NextWord
        If StrComp(Word, "rr") = 0 Then Call AddToken_Z80(TOKEN_Z80_RR): GoTo NextWord
        If StrComp(Word, "rrca") = 0 Then Call AddToken_Z80(TOKEN_Z80_RRCA): GoTo NextWord
        If StrComp(Word, "rrc") = 0 Then Call AddToken_Z80(TOKEN_Z80_RRC): GoTo NextWord
        If StrComp(Word, "rrd") = 0 Then Call AddToken_Z80(TOKEN_Z80_RRD): GoTo NextWord
        If StrComp(Word, "rst") = 0 Then Call AddToken_Z80(TOKEN_Z80_RST): GoTo NextWord
        'Keywords
        If StrComp(Word, "return") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_RETURN): GoTo NextWord
        
    ElseIf Asc1 = S Then '--- `s` ---
        'Z80 instructions
        If StrComp(Word, "sbc") = 0 Then Call AddToken_Z80(TOKEN_Z80_SBC): GoTo NextWord
        If StrComp(Word, "scf") = 0 Then Call AddToken_Z80(TOKEN_Z80_SCF): GoTo NextWord
        If StrComp(Word, "set") = 0 Then Call AddToken_Z80(TOKEN_Z80_SET): GoTo NextWord
        If StrComp(Word, "sla") = 0 Then Call AddToken_Z80(TOKEN_Z80_SLA): GoTo NextWord
        If StrComp(Word, "sra") = 0 Then Call AddToken_Z80(TOKEN_Z80_SRA): GoTo NextWord
        If StrComp(Word, "sll") = 0 Then Call AddToken_Z80(TOKEN_Z80_SLL): GoTo NextWord
        If StrComp(Word, "srl") = 0 Then Call AddToken_Z80(TOKEN_Z80_SRL): GoTo NextWord
        If StrComp(Word, "sub") = 0 Then Call AddToken_Z80(TOKEN_Z80_SUB): GoTo NextWord
        'Keywords
        If StrComp(Word, "slot") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_SLOT): GoTo NextWord
        If StrComp(Word, "start") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_START): GoTo NextWord
        If StrComp(Word, "stop") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_STOP): GoTo NextWord
        If StrComp(Word, "struct") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_STRUCT): GoTo NextWord
        
    ElseIf Asc1 = W Then '--- `w` ---
        'Keywords
        If StrComp(Word, "word") = 0 Then Call AddToken_Keyword(TOKEN_KEYWORD_WORD): GoTo NextWord
        'There are no `w` registers
        GoTo WordError
        
    ElseIf Asc1 = X Then '--- `x` ---
        'Z80 instructions
        If StrComp(Word, "xor") = 0 Then Call AddToken_Z80(TOKEN_Z80_XOR): GoTo NextWord
        'TODO: The `x` operator
        'There are no `x` registers
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_BLOCK_OPEN Then '--- `{` ---
        Call AddToken_Group(TOKEN_GROUP_BLOCKOPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_BLOCK_CLOSE Then  '--- `}` ---
        Call AddToken_Group(TOKEN_GROUP_BLOCKCLOSE): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_OPEN Then '--- `(` ---
        Call AddToken_Group(TOKEN_GROUP_PARENOPEN): GoTo NextWord
    ElseIf Asc1 = SYNTAX_PAREN_CLOSE Then '--- `)` ---
        Call AddToken_Group(TOKEN_GROUP_PARENCLOSE): GoTo NextWord
    
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
        Call AddToken_Operator(TOKEN_OPERATOR_ADD): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '--- `-` ---
        Call AddToken_Operator(TOKEN_OPERATOR_SUB): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '--- `*` ---
        Call AddToken_Operator(TOKEN_OPERATOR_MUL): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '--- `/` ---
        Call AddToken_Operator(TOKEN_OPERATOR_DIV): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '--- `^` ---
        Call AddToken_Operator(TOKEN_OPERATOR_POW): GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '--- `\` ---
        Call AddToken_Operator(TOKEN_OPERATOR_MOD): GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
    'Ignore anything beginning with a letter before A
    If Asc1 < A Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > S Then GoTo WordError
    'If there's only 1 letter total, then it's obviously a 1-letter register
    If Asc2 = 0 Then GoTo OneLetter
    'If a semi-colon occurs at the second letter, then the register must be 1 letter
    If Asc2 = SYNTAX_HINT1 Then GoTo OneLetter
    'Also an apostrophe (shadow register) is the same
    If Asc2 = SYNTAX_HINT2 Then GoTo OneLetter
    
    'If there's only 2 letters in total, it has to be a two-letter register
    If Asc3 = 0 Then GoTo TwoLetters
    'If a semi-colon or apostrophe occurs at the third letter, _
     it could be a two-letter register
    If Asc3 = SYNTAX_HINT1 Then GoTo TwoLetters
    If Asc3 = SYNTAX_HINT2 Then GoTo TwoLetters
    
    GoTo WordError

OneLetter:
    If Asc1 = A Then Call AddToken_Register(TOKEN_REGISTER_A): GoTo NextWord
    If Asc1 = B Then Call AddToken_Register(TOKEN_REGISTER_B): GoTo NextWord
    If Asc1 = C Then Call AddToken_Register(TOKEN_REGISTER_C): GoTo NextWord
    If Asc1 = D Then Call AddToken_Register(TOKEN_REGISTER_D): GoTo NextWord
    If Asc1 = E Then Call AddToken_Register(TOKEN_REGISTER_E): GoTo NextWord
    If Asc1 = H Then Call AddToken_Register(TOKEN_REGISTER_H): GoTo NextWord
    If Asc1 = I Then Call AddToken_Register(TOKEN_REGISTER_I): GoTo NextWord
    If Asc1 = L Then Call AddToken_Register(TOKEN_REGISTER_L): GoTo NextWord
    If Asc1 = R Then Call AddToken_Register(TOKEN_REGISTER_R): GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = A Then       '"af"
        If Asc2 = F Then Call AddToken_Register(TOKEN_REGISTER_AF): GoTo NextWord
    ElseIf Asc1 = B Then   '"bc"
        If Asc2 = C Then Call AddToken_Register(TOKEN_REGISTER_BC): GoTo NextWord
    ElseIf Asc1 = D Then  '"de"
        If Asc2 = E Then Call AddToken_Register(TOKEN_REGISTER_DE): GoTo NextWord
    ElseIf Asc1 = H Then  '"hl"
        If Asc2 = L Then Call AddToken_Register(TOKEN_REGISTER_HL): GoTo NextWord
    ElseIf Asc1 = I Then  '"ix"/"iy"
        If Asc2 = X Then Call AddToken_Register(TOKEN_REGISTER_IX): GoTo NextWord
        If Asc2 = Y Then Call AddToken_Register(TOKEN_REGISTER_IY): GoTo NextWord
    ElseIf Asc1 = P Then  '"pc"
        If Asc2 = C Then Call AddToken_Register(TOKEN_REGISTER_PC): GoTo NextWord
    ElseIf Asc1 = S Then  '"sp"
        If Asc2 = P Then Call AddToken_Register(TOKEN_REGISTER_SP): GoTo NextWord
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
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'Add : Add a token (any) to the stream _
 ======================================================================================
Private Sub Add(ByRef Token As Long)
    'The param MUST be 0-255, but is given as long for speed reasons; _
     the rest of this class will ensure that this is never out of bounds
    Let Tokens(TokensCursor) = Token
    Let TokensCursor = TokensCursor + 1
End Sub
