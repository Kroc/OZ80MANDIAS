VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_Assembler"
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
'CLASS :: oz80_Assembler

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

'This structure is used when parsing parameters of Z80 instructions
Private Type oz80Param
    Register As OZ80_TOKEN_DATA         'If a register, which one
    Value As Long                       'The numerical value of the parameter
    ValueUnknown As Boolean             'If the parameter has unknown values (labels)
    Kind As OZ80_PARAM                  'As below, a brief hint at the parameter type
    Mask As OZ80_MASK                   'A detailed bit-mask of parameter features
End Type

Private Enum OZ80_PARAM
    PARAM_REGISTER = 1
    PARAM_NUMBER = 2
    PARAM_MEMORY = 3
End Enum

'Specific values used with opcodes for various instruction features: _
 E.g. the use of IX/IY often requires extra bytes to denote these
Private Enum OZ80_OPCODE
    'Opcodes that accept registers usually add these offsets to the opcode
    OPCODE_A = 7
    OPCODE_B = 0
    OPCODE_C = 1
    OPCODE_D = 2
    OPCODE_E = 3
    OPCODE_H = 4
    OPCODE_L = 5
    
    'Register-pairs work in the same manner
    OPCODE_BC = &H0
    OPCODE_DE = &H10
    OPCODE_HL = &H20: OPCODE_IX = OPCODE_HL: OPCODE_IY = OPCODE_IX
    OPCODE_SP = &H30: OPCODE_AF = OPCODE_SP
    
    'For memory references such as `(ix+$8)`, these bytes are added as extra opcodes
    OPCODE_IX8 = &HDD
    OPCODE_IY8 = &HFD
    
    'The flags add these offsets to the base opcode
    OPCODE_FLAG_NZ = 0                  '..---...
    OPCODE_FLAG_Z = &H8                 '..--x...
    OPCODE_FLAG_NC = &H10               '..-x-...
    OPCODE_FLAG_C = &H18                '..-xx...
    OPCODE_FLAG_PO = &H20               '..x--...
    OPCODE_FLAG_PE = &H28               '..x-x...
    OPCODE_FLAG_P = &H30                '..xx-...
    OPCODE_FLAG_M = &H38                '..xxx...
End Enum

'When fetching parameters for Z80 instructions a bit-pattern is used to relay the _
 type and features of the paramter, i.e. whether it's a register or memory reference, _
 what register it is and if it has an associated value e.g. `(ix+$8)`
Private Enum OZ80_MASK
    'Single registers
    MASK_A = 2 ^ 0
    MASK_B = 2 ^ 1
    MASK_C = 2 ^ 2
    MASK_D = 2 ^ 3
    MASK_E = 2 ^ 4
    MASK_H = 2 ^ 5
    MASK_L = 2 ^ 6
    
    'Register pairs
    MASK_AF = 2 ^ 7
    MASK_BC = 2 ^ 8
    MASK_DE = 2 ^ 9
    MASK_HL = 2 ^ 10
    MASK_IX = 2 ^ 11
    MASK_IY = 2 ^ 12
    MASK_SP = 2 ^ 13
    
    'Values
    MASK_8 = 2 ^ 14                     'An 8-bit number
    MASK_16 = 2 ^ 15                    'A 16-bit number
    
    'Memory reference
    MASK_MEM8 = 2 ^ 16                  '`in ($8)`
    MASK_MEM16 = 2 ^ 17                 '`jp ($16)`
    MASK_MEMHL = 2 ^ 18                 '`ld a, (hl)`
    MASK_MEMIX8 = 2 ^ 19                '`ld a, (ix+$8)`
    MASK_MEMIY8 = 2 ^ 20                '`ld a, (iy+$8)`
    MASK_MEMSP = 2 ^ 21                 '`ld a, (sp)`
    MASK_MEMC = 2 ^ 22                  '`in (c)`
    
    'Condition flags (e.g. `ret`, `call`)
    MASK_FLAGS_CZ = 2 ^ 23              'Flags c, nc, z or nz
    MASK_FLAGS_MP = 2 ^ 24              'Flags m, p, pe or po
    MASK_FLAGS = MASK_FLAGS_CZ Or MASK_FLAGS_MP
    
    MASK_7 = 2 ^ 25                     '0-7, e.g. `BIT` instruction
    MASK_2 = 2 ^ 26                     '0-2, e.g. `IM` instruction
    
    'Some shorthand:
    MASK_IXY = MASK_IX Or MASK_IY
    MASK_IXY8 = MASK_MEMIX8 Or MASK_MEMIY8
    
    MASK_HL_IXY = MASK_HL Or MASK_IXY
    MASK_A_HL_IXY = MASK_A Or MASK_HL_IXY
    MASK_ABCDEHL = MASK_A Or MASK_B Or MASK_C Or MASK_D Or MASK_E Or MASK_H Or MASK_L
    MASK_ABCDEHL8 = MASK_ABCDEHL Or MASK_8
    MASK_MEM_HL_IXY8 = MASK_MEMHL Or MASK_IXY8
    MASK_BC_DE_HL = MASK_BC Or MASK_DE Or MASK_HL
End Enum

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

'For timing this process
Private Stopwatch As PerformanceCounter

'The current source file being processed
Private WithEvents TokenFile As oz80_TokenStream
Attribute TokenFile.VB_VarHelpID = -1

'The line number of the previous token, so that we can tell when a label/property _
 definition occurs on a new line and isn't therefore part of a list
Private PrevLine As Long

Private At As Long                      'Current position within the ROM (absolute)
Private Bank As Long                    'Current bank number
Private Slot As Long                    'Current slot number the bank applies to

Private ROM() As Byte                   'The output!
Private ROMSize As Long                 'Size of the ROM (32/64/128/256/512/1024 KB)

'The token stream doesn't store the text of named items, merely a reference to the _
 string stored here. We use this for error messages and concatenating properties to _
 their parent variable / label
Private Names As Scripting.Dictionary

'Label data will be stored here, along with the label's properties, _
 e.g. `:here`, `:here.bank`, `:here.slot`
Private Labels As Scripting.Dictionary

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message(ByRef Text As String)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Set Stopwatch = New PerformanceCounter
    Set Labels = New Scripting.Dictionary
    
    'NOTE:
    '   +rr     BC = +00
    '           DE = +10
    '           HL = +20 (Also IX, IY)
    '           SP = +30 (Also AF)
    '   rr      IX = DD
    '           IY = FD
    '           HL = rr and ?? is omitted
    '   +f      nz = +0
    '           z  = +8
    '           nc = +10
    '           c  = +18
    '           po = +20
    '           pe = +28
    '           p  = +30
    '           m  = +38
    '   +b      bit * 8
    '   +im     0  = +0
    '           1  = +10
    '           2  = +18
        
    'Z80    Param 1                 Param 2             Opcode          type    done _
     ----------------------------------------------------------------------------------
    'ADC    a                       a|b|c|d|e|h|l       88+r            reg     x
    '                               $8                  CE ??           val     x
    '                               (hl|ix+$8|iy+$8)    rr 8E ??        mem     x
    '       hl                      bc|de|hl|sp         ED 4A+rr        reg     x
                                
    'ADD    a                       a|b|c|d|e|h|l       80+r            reg     x
    '                               $8                  C6 ??           val     x
    '                               (hl|ix+$8|iy+$8)    rr 86 ??        mem     x
    '       hl                      bc|de|hl|sp         09+rr           reg     x
    '       ix                      bc|de|ix|sp         rr 09+rr        reg     x
    '       iy                      bc|de|iy|sp         rr 09+rr        reg     x
    
    'AND    a|b|c|d|e|h|l                               A0+r            reg     x
    '       $8                                          E6 ??           val     x
    '       (hl|ix+$8|iy+$8)                            rr A6 ??        mem     x
    
    'BIT    0-7                     a|b|c|d|e|h|l       CB 40+r+b       reg bit
    '                               (hl|ix+$8|iy+$8)    rr CB ?? 40+r+b mem bit
    
    'CALL   $16                                         CD ?? ??        val
    '       c|nc|m|p|z|nz|pe|po     $16                 C4+f            con
    
    'CCF                                                3F              ---
    
    'CP     a|b|c|d|e|h|l                               B8+r            reg
    '       (hl|ix+$8|iy+$8)                            rr BE ??        mem
    '       $8                                          FE ??           val
    
    'CPD                                                ED A9           ---
    'CPDR                                               ED B9           ---
    'CPI                                                ED A1           ---
    'CPIR                                               ED B1           ---
    'CPL                                                2F              ---
    'DAA                                                27              ---
    
    'DEC    a|b|c|d|e|h|l                               05+r*8          reg
    '       (hl|ix+$8|iy+$8)                            rr 35 ??        mem
    '       bc|de|hl|sp                                 0B+rr           reg
    '       ix|iy                                       rr 2B           ---
    
    'DI                                                 F3              ---
    
    'DJNZ   $8                                          10 ??           val
    
    'EI                                                 FB              ---
    
    'EX     (sp)                    hl|ix|iy            rr E3           ---
    '       af                      af                  08              ---
    '       de                      hl                  EB              ---
    
    'EXX                                                D9              ---
    'HALT                                               76              ---
    
    'IM     0-2                                         46+im
    
    'IN     a                       ($8)                DB ??
    '       (c)                                         ED 70
    '       a|b|c|d|e|h|l           (c)                 ED 40+r
    
    'INC    a|b|c|d|e|h|l                               04+r*8
    '       bc|de|hl|sp|ix|iy                           rr 03+rr
    '       (hl|ix+$8|iy+$8)                            rr 34 ??
    
    'IND                                                ED AA
    'INDR                                               ED BA
    'INI                                                ED A2
    'INIR                                               ED B2
    
    'JP     $16                                         C3 ?? ??
    '       (hl|ix|iy)                                  rr E9
    '       c|nc|m|p|z|nz|pe|po     $16                 C2+f ?? ??
    
    'JR     $8                                          18 ??
    '       c|nc|z|nz               $8                  20+f
    
    'LD     i                       a                   ED 47
    '       r                       a                   ED 4F
    '       a                       i                   ED 57
    '                               r                   ED 5F
    '                               a|b|c|d|e|h|l       78+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 7E ??        mem
    '                               $8                  3E ??           val
    '                               (bc|de)             0A+rr           mem !
    '                               ($16)               3A ?? ??        mem
    '       b                       a|b|c|d|e|h|l       40+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 46 ??        mem
    '                               $8                  06 ??           val
    '       c                       a|b|c|d|e|h|l       48+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 4E ??        mem
    '                               $8                  0E ??           val
    '       d                       a|b|c|d|e|h|l       50+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 56 ??        mem
    '                               $8                  16 ??           val
    '       e                       a|b|c|d|e|h|l       58+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 5E ??        mem
    '                               $8                  1E ??           val
    '       h                       a|b|c|d|e|h|l       60+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 66 ??        mem
    '                               $8                  26 ??           val
    '       l                       a|b|c|d|e|h|l       68+r            reg
    '                               (hl|ix+$8|iy+$8)    rr 6E ??        mem
    '                               $8                  2E ??           val
    '       bc                      ($16)               ED 4B ?? ??     mem
    '                               $16                 01+rr ?? ??     reg val
    '       de                      ($16)               ED 5B ?? ??     mem
    '                               $16                 01+rr ?? ??     reg val
    '       hl                      ($16)               2A ?? ??        mem
    '                               $16                 01+rr ?? ??     reg val
    '       sp                      ($16)               ED 7B ?? ??     mem
    '                               hl|ix|iy            rr F9
    '                               $16                 01+rr ?? ??     reg val
    '       ix|iy                   ($16)               rr 2A ?? ??
    '                               $16                 rr 21 ?? ??
    '       (hl)                    a|b|c|d|e|h|l       70+r
    '                               $8                  36 ??
    '       (bc)                    a                   02+rr
    '       (de)                    a                   02+rr
    '       ($16)                   a                   32 ?? ??
    '                               bc                  ED 43 ?? ??
    '                               de                  ED 53 ?? ??
    '                               hl|ix|iy            rr 22 ?? ??
    '                               sp                  ED 73 ?? ??
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l       rr 70+r ??
    '                               $8                  rr 36 ?? ??
    
    'LDD                                                ED A8
    'LDDR                                               ED B8
    'LDI                                                ED A0
    'LDIR                                               ED B0
    'NEG                                                ED 44
    'NOP                                                00
    
    'OR     a|b|c|d|e|h|l                               B0+r
    '       (hl|ix+$8|iy+$8)                            rr B6 ??
    '       $n                                          F6 ??
    
    'OUT    ($8)                    a                   D3 ??
    '       (C)                     0                   ED 71
    '                               a|b|c|d|e|h|l       ED 41+r*8
    
    'OUTD                                               ED AB
    'OTDR                                               ED BB
    'OUTI                                               ED A3
    'OTIR                                               ED B3
    
    'POP    af|bc|de|hl|ix|iy                           rr C1+rr
    
    'PUSH   af|bc|de|hl|ix|iy                           rr C5+rr
    
    'RES    0-7                     a|b|c|d|e|h|l       CB 80+r+b
    '                               (hl|ix+$8|iy+$8)    rr CB ?? 86+b
    
    'RET                                                C9
    '       c|nc|m|p|z|nz|pe|po                         C0+f
    
    'RETI                                               ED 4D
    'RETN                                               ED 45
    'RLA                                                17
    
    'RL     a|b|c|d|e|h|l                               CB 10+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 16
    
    'RLCA                                               07
    
    'RLC    a|b|c|d|e|h|l                               CB 00+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 06
    
    'RLD                                                ED 6F
    'RRA                                                1F
    
    'RR     a|b|c|d|e|h|l                               CB 18+r
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 1E
    
    'RRCA                                               0F              ---     x
    
    'RRC    a|b|c|d|e|h|l                               CB 08+r         reg     x
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 0E     mem     x
    
    'RRD                                                ED 67           ---     x
    
    'RST    0|08|10|18|20|28|30|38                      C7+num
    
    'SBC    a|b|c|d|e|h|l                               98+r            reg     x
    '       (hl|ix+$8|iy+S8)                            rr 9E ??        mem     x
    '       a                       $8                  DE ??
    '       hl                      bc|de|hl|sp         ED+rr
    
    'SCF                                                37              ---     x
    
    'SET    0-7                     a|b|c|d|e|h|l       CB C0+b+r
    '                               (hl|ix+$8|iy+$8)    rr CB ?? C6+b+r
    
    'SLA    a|b|c|d|e|h|l                               CB 20+r         reg     x
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 26     mem     x
    
    'SRA    a|b|c|d|e|h|l                               CB 28+r         reg     x
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 2E     mem     x
    
    'SLL    a|b|c|d|e|h|l                               CB 30+r         reg     x
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 36     mem     x
    
    'SRL    a|b|c|d|e|h|l                               CB 38+r         reg     x
    '       (hl|ix+$8|iy+$8)                            rr CB ?? 3E     mem     x
    
    'SUB    a|b|c|d|e|h|l                               90+r            reg     x
    '       (hl|ix+$8|iy+$8)                            rr 96 ??        mem     x
    '       $8                                          D6 ??           val     x
    
    'XOR    a|b|c|d|e|h|l                               A8+r            reg     x
    '       (hl|ix+$8|iy+$8)                            rr AE ??        mem     x
    '       $8                                          EE ??           val     x
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Labels = Nothing
    Set Stopwatch = Nothing
    Erase ROM
End Sub

'EVENT <TokenFile> Error : Detailed error message from a TokenStream object _
 ======================================================================================
Private Sub TokenFile_Error( _
    ByVal Number As OZ80_ERROR, Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)
    RaiseEvent Error(Number, Description, Line, Col)
End Sub

'EVENT <TokenFile> Message : Basic progress message from a TokenStream object _
 ======================================================================================
Private Sub TokenFile_Message(Text As String)
    RaiseEvent Message(Text)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Process : Assemble a Z80 binary from a given token stream _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    RaiseEvent Message("OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision)
    RaiseEvent Message("* Assembling token stream")
    
    'We'll measure the time this takes
    Call Stopwatch.Reset
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
    'Create a tokeniser object to hold the machine representation of the text files; _
     the assembler doesn't work with the original text directly
    Set TokenFile = New oz80_TokenStream
    'Explode the source code file into tokens
    Let Assemble = TokenFile.TokeniseFile(FilePath)
    If Assemble <> OZ80_ERROR_NONE Then Stop
    
    'Stage 2: Assemble _
     ----------------------------------------------------------------------------------
    'Set the assembly point to the beginning
    Let Bank = 0
    Let Slot = 0
    Let At = 0
    
    'Default to 32 KB ROM (the ROM size will double each time it overflows)
    Let ROMSize = 2 ^ 15
    ReDim ROM(0 To ROMSize - 1) As Byte
    
    'Phase 1 : Determine the size of the assembly and remember all label positions
    Let Assemble = ScopeRoot()
    
    'Phase 2 : 'Fill in the blanks' -- go back and evaluate the unknowns
    'TODO
    
    RaiseEvent Message("- Final ROM size is " & (ROMSize \ 1024) & " KB")
    RaiseEvent Message("- Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
End Function

'/// SCOPE HANDLERS ///////////////////////////////////////////////////////////////////

'This is a structured map of the OZ80 syntax. _
 The first column on the left is the name of a scope. For each scope, the columns to _
 the right represent the heirarchy of allowable items in that scope. Whatever is in _
 the third column is only allowed to follow what is in the second column. Items below _
 one another within a column represent more than one item being allowed within a scope.
'Another scope name within a column represents a nested scope. Items within brackets _
 show optional items. _
'An example:

'           INCLUDE     <list>
'                       BINARY      <text>  [START <expr>]  [LENGTH <expr>]
'                                                           [STOP   <expr>]

'Here the `INCLUDE` keyword can be followed by a list, *or* the keyword `BINARY`. _
 If `BINARY` is used then it may be optionally followed by the `START` keyword and an _
 expression. Only if the `START` keyword is included can `LENGTH` *or* `STOP` be used, _
 but not both. `INCLUDE <list>` cannot be followed by `START` &c.

'NOTES: *   You can take `{ASCII}` to mean any ASCII code 0-127 _
        *   `{operator}` can be any one of `+ - * / ^ \`

'Scope:     1           2           3       4       5       6       7       8 _
 --------------------------------------------------------------------------------------
'<text>     "           {ASCII}     "
'
'<expr>     <str>
'           <number>    [{operator} <expr>]
'           #var        [{operator} <expr>]
'           :label      [{operator} <expr>]
'           {           <expr>      }
'
'
'<list>     <text>      [,          <list>]
'           <expr>      [,          <list>]
'
'<root>     DEF         #variable   [<list>]
'
'           SECTION     [BANK       <expr>]
'                       [SLOT       <expr>]
'                       [AT         <expr>]
'                       {           <section>
'                       }
'
'           FILL        BINARY      <text>
'                       <list>
'
'           INCLUDE     <list>
'                       BINARY      <text>  [START <expr>]  [LENGTH <expr>]
'                                                           [STOP   <expr>]
'
'           OBJECT      #object     <object>
'
'<section>  PROC        :label      [PARAMS register
'                                           number]
'                                   [RETURN register
'                                           number]
'                                   {       <code>
'                                   }
'
'           STRUCT      :label      AS      #object         <struct>
'
'           TABLE       :label      <table>
'
'           :label
'
'           <code>
'
'<code>     .label
'           DATA        <list>
'           BYTE        <list>
'           WORD        <list>
'
'           DEF         #variable   [<list>]
'
'           <asm>

'ScopeRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Function ScopeRoot() As OZ80_ERROR
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
Continue:
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_LABEL Then
'        Debug.Print TokenFile.Line & "," & TokenFile.Col & ": " _
'            & TokenFile.StringData(TokenFile.Data)
        
        'If the label is already defined, we have a problem
        If Labels.Exists(TokenFile.Data) = True Then
            'TODO: Error - label already defined
'            Let ScopeRoot = RaiseError(OZ80_ERROR_DUPLICATE, _
'                "Duplicate label definition! '" & _
'                TokenFile.StringData(TokenFile.Data) & "' is already defined.", _
'                TokenFile.Line, TokenFile.Col _
'            ): Exit Function
            Stop
        Else
            'Add a label using the token's data value for an index key as this will _
             be the string hash generated in the token stream
            Call Labels.Add(TokenFile.Data, At)
        End If
        
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'Check allowed keywords at this Scope
        If TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DEF Then
            'TODO
            
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_FILL Then
            'TODO
            
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_INCLUDE Then
            'TODO
            
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_TABLE Then
            'TODO
            
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_STRUCT Then
            'TODO
            
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_PROC Then
            'TODO
            
        End If
    Else
        'The root Scope also allows for source code, a Scope shared by blocks
        Let ScopeRoot = ScopeCode()
        GoTo Continue
    End If
    
    If NextToken() = True Then GoTo Continue
End Function

'ScopeCode : Process areas where source code is allowed _
 ======================================================================================
Private Function ScopeCode() As OZ80_ERROR
    'This scope allows data statements and Z80 source code, e.g. root and procedures
    
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'TODO: Which keywords are allowed here
        If TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_BYTE Then
            'TODO
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DATA Then
            'TODO
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DEF Then
            'TODO
        ElseIf TokenFile.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_WORD Then
            'TODO
        Else
            'TODO: Error - keyword not allowed at this Scope
'            Stop
        End If
        
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_Z80 Then
        'Process a Z80 instruction ...
        Let ScopeCode = ScopeZ80()
        Exit Function
    
    Else
        'TODO: Error - unexpected content at this point
        Stop
    End If
    
    Call NextToken
End Function

'ScopeExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ScopeExpression( _
    ByRef ReturnValue As Long, ByRef ReturnValueIsUnknown As Boolean _
) As OZ80_ERROR
    Dim Operand1 As Long                'First value in an expression
    Dim Operand2 As Long                'Next value in an expression
    Dim Operator As OZ80_TOKEN_DATA     'Operator to use with the expressions
    
    Let ReturnValueIsUnknown = False
    
    'Begin with fetching the first operand
    Let ScopeExpression = ScopeOperand(Operand1, ReturnValueIsUnknown)
    If ScopeExpression <> OZ80_ERROR_NONE Then Stop
    Let ReturnValue = Operand1
    
'    'The file could legitimately end with an expression, so don't error out if the _
'     token stream comes to an end since the expression is valid. The parent scopes _
'     will handle an unexpected end error
'    If NextToken() = False Then Exit Function
    
    'If next token is not an operator assume the expression has finished, _
      the parent will handle whatever the next token is used for
    If TokenFile.Kind <> OZ80_TOKEN.TOKEN_OPERATOR Then Exit Function
    
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = TokenFile.Data
    
    'Get the second operand
    'TODO: Error - unexpected end
    If NextToken() = False Then Stop
    Let ScopeExpression = ScopeOperand(Operand2, ReturnValueIsUnknown)
    'Cannot have an operator hanging there without an operand following
    If ScopeExpression <> OZ80_ERROR_NONE Then
        'TODO: Error - invalid use of operator
        Stop
    End If
    
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If ReturnValueIsUnknown = True Then
        'Do nothing here, just fall out of the if block, skipping operator processing
    
    ElseIf Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 + Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 - Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let ReturnValue = Operand1 * Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        'TODO: Should we use Doubles so as to be accurate?
        Let ReturnValue = Operand1 / Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let ReturnValue = Operand1 ^ Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        Let ReturnValue = Operand1 Mod Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        Let ReturnValue = Operand1 And Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        Let ReturnValue = Operand1 Or Operand2
        
    End If
    
'    'The file could legitimately end with an expression, so don't error out if the _
'     token stream comes to an end since the expression is valid. The parent scopes _
'     will handle an unexpected end error
'    If NextToken() = False Then Exit Function
    
    'Does the expression continue?
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_OPERATOR Then
        'Set the current result as the first operand and go around again
        Let Operand1 = ReturnValue
        GoTo Continue
    End If
End Function

'ScopeMemory : Work out a memory reference, i.e. `($????)` / `(ix+$??)` _
 ======================================================================================
Private Function ScopeMemory( _
    ByRef Register As OZ80_TOKEN_DATA, _
    ByRef Value As Long, ByRef Unknown As Boolean _
) As OZ80_ERROR
    'If the current token is an open parentheses move forward into the expression
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_PARENOPEN Then
        If NextToken() = False Then Stop
    End If
    
    'The first item in the expression can be a register, typically HL, IX or IY
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
        'The `IN` & `OUT` instructions accept only the C register
        If TokenFile.Data = TOKEN_REGISTER_C Then
            Let Register = TOKEN_REGISTER_C
        
        '`LD` accepts registers BC & DE
        ElseIf TokenFile.Data = TOKEN_REGISTER_BC Then
            Let Register = TOKEN_REGISTER_BC
        ElseIf TokenFile.Data = TOKEN_REGISTER_DE Then
            Let Register = TOKEN_REGISTER_DE
            
        ElseIf TokenFile.Data = TOKEN_REGISTER_HL Then
            Let Register = TOKEN_REGISTER_HL
            
        ElseIf TokenFile.Data = TOKEN_REGISTER_IX Then
            Let Register = TOKEN_REGISTER_IX
            'IX can have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf TokenFile.Data = TOKEN_REGISTER_IY Then
            Let Register = TOKEN_REGISTER_IY
            'IY can also have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf TokenFile.Data = TOKEN_REGISTER_SP Then
            Let Register = TOKEN_REGISTER_SP
        Else
            'ERROR: register not allowed here
            Stop
        End If
        
        'Move to the next token, then check that it's the end of the memory reference
        'TODO: out of tokens
        If NextToken() = False Then Stop
        GoTo EndScope
        
AddExpr:
        'Look for the `+` following the register name
        If NextToken() = False Then Stop
        'If present, skip over it and check for an expression
        If TokenFile.Kind = OZ80_TOKEN.TOKEN_OPERATOR Then
            If TokenFile.Data = TOKEN_OPERATOR_ADD Then
                If NextToken() = False Then Stop
            End If
            
        'If it's just a register in parens, then finish
        ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
            Exit Function
        End If
    End If
    
    'Try and calculate the expression. It might come back unknown and this will be _
     passed on to the parent to decide if that's a problem or not
    'TODO: handle overflow
    Let ScopeMemory = ScopeExpression(Value, Unknown)

EndScope:
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
        Call NextToken
        Exit Function
    End If
    'ERROR: Unexpected token in memory expression
    Stop
End Function

'ScopeOperand : Get the value of one half of a calculation _
 ======================================================================================
Private Function ScopeOperand( _
    ByRef ReturnValue As Long, ByRef ReturnValueIsUnknown As Boolean _
) As OZ80_ERROR
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: If the label is not yet known, the expression cannot be evaluated _
         until the label is reached
        Let ReturnValueIsUnknown = True
         
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_NUMBER Then
        'This is an easy one
        Let ReturnValue = TokenFile.Data
        
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PROPERTY Then
        'If a property, check the scope to see which variable / label it belongs to
        'TODO
        Let ReturnValueIsUnknown = True
                
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_VARIABLE Then
        'TODO: If the variable is not yet defined, error
        Let ReturnValueIsUnknown = True
        
    Else
        'ERROR: Not a valid operand. We won't produce an error message here as the _
         error might be exepected (i.e. an optional parameter). We fold upwards and _
         whomever knows best will deal with the error
        Debug.Print TokenFile.Line & ", " & TokenFile.Col
        Stop
        Let ScopeOperand = OZ80_ERROR_OPERAND
    End If
    
    'Move beyond the current token now it's been processed
    Call NextToken
End Function

'ScopeZ80 : Process a Z80 instruction _
 ======================================================================================
Private Function ScopeZ80() As OZ80_ERROR
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN_DATA: Let Instruction = TokenFile.Data
    'Remember the line number of this since a label or property on a different line _
     will be a definition, not a parameter
    Dim Line As Long: Let Line = TokenFile.Line
    'Skip over the Z80 instruction token now that we know it
    If NextToken() = False Then Stop
    
    'If there appears to be a parameter, skip over the parameterless instructions
    If IsParameter() = True Then GoTo With1Parameter
    
    '----------------------------------------------------------------------------------
    
    'Z80 instructions without parameters:
    If Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CCF Then
        Call AddByte(&H3F)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CPD Then
        Call AddByte(&HED): Call AddByte(&HA9)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CPDR Then
        Call AddByte(&HED): Call AddByte(&HB9)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CPI Then
        Call AddByte(&HED): Call AddByte(&HA1)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CPIR Then
        Call AddByte(&HED): Call AddByte(&HB1)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CPL Then
        Call AddByte(&H2F)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_DAA Then
        Call AddByte(&H27)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_DI Then
        Call AddByte(&HF3)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_EI Then
        Call AddByte(&HFB)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_EXX Then
        Call AddByte(&HD9)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_HALT Then
        Call AddByte(&H76)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_IND Then
        Call AddByte(&HED): Call AddByte(&HAA)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_INDR Then
        Call AddByte(&HED): Call AddByte(&HBA)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_INI Then
        Call AddByte(&HED): Call AddByte(&HA2)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_INIR Then
        Call AddByte(&HED): Call AddByte(&HB2)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_LDD Then
        Call AddByte(&HED): Call AddByte(&HA8)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_LDDR Then
        Call AddByte(&HED): Call AddByte(&HB8)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_LDI Then
        Call AddByte(&HED): Call AddByte(&HA0)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_LDIR Then
        Call AddByte(&HED): Call AddByte(&HB0)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_NEG Then
        Call AddByte(&HED): Call AddByte(&H44)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_NOP Then
        Call AddByte(&H0)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OUTD Then
        Call AddByte(&HED): Call AddByte(&HAB)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OTDR Then
        Call AddByte(&HED): Call AddByte(&HBB)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OUTI Then
        Call AddByte(&HED): Call AddByte(&HA3)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OTIR Then
        Call AddByte(&HED): Call AddByte(&HB3)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RET Then
        Call AddByte(&HC9)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RETI Then
        Call AddByte(&HED): Call AddByte(&H4D)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RETN Then
        Call AddByte(&HED): Call AddByte(&H45)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RLA Then
        Call AddByte(&H17)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RLCA Then
        Call AddByte(&H7)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RLD Then
        Call AddByte(&HED): Call AddByte(&H6F)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RRA Then
        Call AddByte(&H1F)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RRCA Then
        Call AddByte(&HF)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RRD Then
        Call AddByte(&HED): Call AddByte(&H67)
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SCF Then
        Call AddByte(&H37)
    Else
        Debug.Print TokenFile.Line, TokenFile.Col
        Stop
    End If
    Exit Function
    
    '----------------------------------------------------------------------------------
With1Parameter:
    'Parameter expressions might return an unknown value _
     (such as when a label, which has not yet been reached, is used)
    Dim Param1 As oz80Param, Param2 As oz80Param
    'Get the first parameter
    Let ScopeZ80 = GetParameter(Param1)
    If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
    
    'Check if a second parameter exists
    If IsParameter() = True Then GoTo With2Parameters
    
    If Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_AND Then '...................... AND ...
        If TestOpcode(&HA0, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HA6, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        If TestOpcode(&HE6, Param1, MASK_8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CALL Then '................ CALL ...
        If TestOpcode(&HCD, Param1, MASK_16) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CP Then '.................... CP ...
        If TestOpcode(&HB8, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HBE, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        If TestOpcode(&HFE, Param1, MASK_8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_DEC Then '.................. DEC ...
        If TestOpcode(&H5, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HB, Param1, MASK_BC_DE_HL Or MASK_SP) Then Exit Function
        If TestOpcode(&H2B, Param1, MASK_IXY) Then Exit Function
        If TestOpcode(&H35, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_DJNZ Then '................ DJNZ ...
        If TestOpcode(&H10, Param1, MASK_8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_IM Then '.................... IM ...
        'IM 0-2
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_IN Then '.................... IN ...
        If TestOpcode(&HED70, Param1, MASK_MEMC) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_INC Then '.................. INC ...
        If TestOpcode(&H4, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&H3, Param1, MASK_BC_DE_HL Or MASK_SP Or MASK_IXY) Then Exit Function
        If TestOpcode(&H34, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_JP Then '.................... JP ...
        If TestOpcode(&HC3, Param1, MASK_16) Then Exit Function
        'NOTE: Opcode does not support a value for IX/IY. _
               Need to get `AddOpcode` to handle this
        If TestOpcode(&HE9, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_JR Then '.................... JR ...
        If TestOpcode(&H18, Param1, MASK_8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OR Then '.................... OR ...
        If TestOpcode(&HB0, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HB6, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        If TestOpcode(&HF6, Param1, MASK_8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_POP Then '.................. POP ...
        'This is handled further down because `pop` may appear as having more than _
         one paramater as OZ80 allows concatenation of `pop` instructions
        GoTo pop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_PUSH Then '................ PUSH ...
        'This is handled further down because `push` may appear as having more than _
         one paramater as OZ80 allows concatenation of `push` instructions
        GoTo push
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RET Then '.................. RET ...
        'Of instructions with 1 parameter, only `ret` uses flags. _
         Since we cannot tell when parsing the text if "c" is a register or a flag, _
         we check for the `c` flag here and provide the correct opcode offset
        If Param1.Kind = PARAM_REGISTER And Param1.Value = OZ80_OPCODE.OPCODE_C Then
            'Change register C to flag C
            Let Param1.Value = OZ80_OPCODE.OPCODE_FLAG_C
            Let Param1.Mask = Param1.Mask Or MASK_FLAGS_CZ
            Let Param1.Mask = Param1.Mask And Not MASK_C
        End If
        If TestOpcode(&HC0, Param1, MASK_FLAGS) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RL Then '.................... RL ...
        If TestOpcode(&HCB10, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB16, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RLC Then '.................. RLC ...
        If TestOpcode(&HCB00, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB06, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RR Then '.................... RR ...
        If TestOpcode(&HCB18, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB1E, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RRC Then '.................. RRC ...
        If TestOpcode(&HCB08, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB0E, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RST Then '.................. RST ...
        'For `rst` the values allowed are $00, $08, $10, $18, $20, $28, $30 & $38
        If (Param1.Value And &H38) = Param1.Value Then
            If TestOpcode(&HC7 + Param1.Value, Param1, MASK_8) Then Exit Function
        End If
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SBC Then '.................. SBC ...
        If TestOpcode(&H98, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&H9E, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SLA Then '.................. SLA ...
        If TestOpcode(&HCB20, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB26, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SRA Then '.................. SRA ...
        If TestOpcode(&HCB28, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB2E, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SLL Then '.................. SLL ...
        If TestOpcode(&HCB30, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB36, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SRL Then '.................. SRL ...
        If TestOpcode(&HCB38, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HCB3E, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SUB Then '.................. SUB ...
        If TestOpcode(&H90, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&H96, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        If TestOpcode(&HD6, Param1, MASK_8) Then Exit Function
        Stop
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_XOR Then '.................. XOR ...
        If TestOpcode(&HA8, Param1, MASK_ABCDEHL) Then Exit Function
        If TestOpcode(&HAE, Param1, MASK_MEM_HL_IXY8) Then Exit Function
        If TestOpcode(&HEE, Param1, MASK_8) Then Exit Function
        Stop
    
    Else
        'ERROR: Unrecognised Z80 instruction
        Stop
    End If
    Exit Function

    '----------------------------------------------------------------------------------
With2Parameters:
    'Fetch the second parameter
    Let ScopeZ80 = GetParameter(Param2)
    If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
            
    If Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_ADC Then '...................... ADC ...
        'If the first parameter is A: _
         `ADC a, a|b|c|d|e|h|l|$8|(hl|ix+$8|iy+$8)`
        If Param1.Register = TOKEN_REGISTER_A Then
            If TestOpcode(&H88, Param2, MASK_ABCDEHL) Then Exit Function
            If TestOpcode(&H8E, Param2, MASK_MEM_HL_IXY8) Then Exit Function
            If TestOpcode(&HCE, Param2, MASK_8) Then Exit Function
            Stop
            
        'If the first parameter is HL: _
         `ADC hl, bc|de|hl|sp`
        ElseIf Param1.Register = TOKEN_REGISTER_HL Then
            'Add opcodes for `ADC hl, bc|de|hl|sp`
            If TestOpcode(&HED4A, Param2, MASK_BC_DE_HL Or MASK_SP) Then Exit Function
            Stop
        Else
            Stop
        End If
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_ADD Then '.................. ADD ...
        '`ADD a, a|b|c|d|e|h|l|$8|(hl|ix+$8|iy+$8)`
        If Param1.Register = TOKEN_REGISTER_A Then
            If TestOpcode(&H80, Param2, MASK_ABCDEHL) Then Exit Function
            If TestOpcode(&H86, Param2, MASK_MEM_HL_IXY8) Then Exit Function
            If TestOpcode(&HC6, Param2, MASK_8) Then Exit Function
            Stop
            
        '`ADD hl|ix|iy, bc|de|hl/ix/iy|sp`
        Else
            'Add opcodes
            Call AddOpcode(&H9, Param2.Mask, Param2.Value)
            
        End If
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_BIT Then '.................. BIT ...
        'BIT 0-7, a|b|c|d|e|h|l
        'BIT 0-7, (hl|ix+$8|iy+$8)
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_CALL Then '................ CALL ...
        'CALL c|nc|m|p|z|nz|pe|po, $16
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_EX Then '.................... EX ...
        'EX (sp), hl|ix|iy
        'EX af, af'
        'EX de, hl
       
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_IN Then '.................... IN ...
        'IN a, ($8)
        'IN a|b|c|d|e|h|l, (c)
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_JP Then '.................... JP ...
        'JP c|nc|m|p|z|nz|pe|po, $16
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_JR Then '.................... JR ...
        'JR c|nc|z|nz, $8
       
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_LD Then '.................... LD ...
        'LD i|r, a
        'LD a, i|r
        'LD a|b, b|c|d|e|h|l
        'LD a|b|c|d|e|h|l, $8
        'LD a, (bc|de|$16)
        'LD a|b|c|d|e|h|l, (hl|ix+$8|iy+$8)
        'LD bc|de|hl|sp|ix|iy, $16|($16)
        'LD sp, hl|ix|iy
        'LD (hl|ix+$8|iy+$8), a|b|c|d|e|h|l|$8
        'LD (bc|de), a
        'LD ($16), a|bc|de|hl|ix|iy|sp
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_OUT Then '.................. OUT ...
        'OUT ($8), a
        'OUT (c), 0|a|b|c|d|e|h|l
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_RES Then '.................. RES ...
        'RES 0-7, a|b|c|d|e|h|l
        'RES 0-7, (hl|ix+$8|iy+$8)
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SBC Then '.................. SBC ...
        'SBC a, $8
        'SBC a, (ix+$8|iy+$8)
        'SBC hl, bc|de|hl|sp
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_SET Then '.................. SET ...
        'SET 0-7, a|b|c|d|e|h|l
        'SET 0-7, (hl|ix+$8|iy+$8)
        
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_POP Then '.................. POP ...
        'Add the opcode in question
pop:    If TestOpcode(&HC1, Param1, MASK_AF Or MASK_BC_DE_HL Or MASK_IXY) Then
            '(If the token stream ends, it's a valid cut-off point)
'            If NextToken() = False Then Exit Function
            'Check for another parameter:
            If IsParameter() = True Then
                'Fetch the parameter value
                Let ScopeZ80 = GetParameter(Param1)
                If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
                'Add the pop opcode again, and check for another parameter
                GoTo pop
            End If
        End If
    
    ElseIf Instruction = OZ80_TOKEN_DATA.TOKEN_Z80_PUSH Then '................ PUSH ...
        'Add the opcode in question
push:   If TestOpcode(&HC5, Param1, MASK_AF Or MASK_BC_DE_HL Or MASK_IXY) Then
            '(If the token stream ends, it's a valid cut-off point)
'            If NextToken() = False Then Exit Function
            'Check for another parameter:
            If IsParameter() = True Then
                'Fetch the parameter value
                Let ScopeZ80 = GetParameter(Param1)
                If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
                'Add the push opcode again, and check for another parameter
                GoTo push
            End If
        End If
        
    Else
        Debug.Print TokenFile.Line, TokenFile.Col
        Stop
    End If
    Exit Function
    
OutOfTokens:
    'TODO: handle the token stream ending unexpectedly
    Stop
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'AddByte : Add a byte to the assembled output and move along _
 ======================================================================================
Private Sub AddByte(ByRef Value As Byte)
    'Add the byte to the ROM and move to the next byte
    'TODO: Handle banking overflows
    Let ROM(At) = Value
    Let At = At + 1
    'Handle ROM overflowing
    If At > ROMSize Then
        Let ROMSize = ROMSize * 2
        ReDim Preserve ROM(0 To ROMSize) As Byte
        RaiseEvent Message("! ROM size was increased to " & (ROMSize \ 1024) & " KB")
    End If
End Sub

'AddOpcode : Write out an assembled Z80 instruction _
 ======================================================================================
Private Sub AddOpcode( _
    ByRef Opcode As Integer, ByRef Kind As OZ80_MASK, _
    Optional ByRef Value As Long = 0 _
)
    'Opcodes that accept registers/flags add an offset to the given opcode; _
     this offset is provided for us in the `Value` parameter
    'NOTE: `inc` is unlike other instructions and multiplies the register offset by 8
    If Opcode = &H4 Then
        Let Opcode = Opcode + Value * 8
    ElseIf Kind And ( _
        MASK_ABCDEHL Or MASK_AF Or MASK_BC_DE_HL Or MASK_SP Or MASK_IXY Or _
        MASK_FLAGS _
    ) Then
        Let Opcode = Opcode + Value
    End If
    
    'TODO: Bits
    
    'IX/IY have their own opcodes that proceed the main one
    If Kind And (MASK_IX Or MASK_MEMIX8) Then Call AddByte(OZ80_OPCODE.OPCODE_IX8)
    If Kind And (MASK_IY Or MASK_MEMIY8) Then Call AddByte(OZ80_OPCODE.OPCODE_IY8)
    'Are we using a pair of opcodes?
    If (Opcode And &HFF) <> Opcode Then
        'Add the first opcode (hi-byte of our parameter)
        Call AddByte((Opcode And &HFF00&) \ &H100)
        'When a pair of opcodes are used, the presence of IX/IY causes the value to _
         be placed between the pair of opcodes, e.g. `IX CB ?? 0E`
        If Kind And (MASK_MEMIY8) Then
            'TODO: Check for overflow
            'Add the value in between
            Call AddByte(Value And &HFF)
            'Add the second opcode
            Call AddByte(Opcode And &HFF)
        Else
            'Add the second opcode
            Call AddByte(Opcode And &HFF)
        End If
    'Using a single opcode
    Else
        'Add the opcode
        Call AddByte(Opcode And &HFF)
        'Then the value (low-byte first)
        If Kind And (MASK_16 Or MASK_8) Then Call AddByte(Value And &HFF)
        'If a 16-bit value the high-byte comes next
        'TODO: Overflow checking
        If Kind And MASK_16 Then Call AddByte((Value And &HFF00&) \ &H100)
    End If
End Sub

'GetParameter : Fetch the type / value of a Z80 instruction parameter _
 ======================================================================================
Private Function GetParameter(ByRef Result As oz80Param) As OZ80_ERROR
    'Clear any previous result since this is ByRef
    With Result
        Let .Kind = 0
        Let .Mask = 0
        Let .Register = 0
        Let .Value = 0
        Let .ValueUnknown = False
    End With
    
     'If the current token is a register, map this to the opcode offset required
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_REGISTER Then '...................................
        'Set the return field to be a register/flag and capture which
        Let Result.Kind = PARAM_REGISTER
        Let Result.Register = TokenFile.Data
        'Return the opcode modifier in the value field. _
         The `AddOpcode` function will do the right thing with it
        If Result.Register = TOKEN_REGISTER_A Then
            Let Result.Mask = MASK_A: Let Result.Value = OPCODE_A
        ElseIf Result.Register = TOKEN_REGISTER_B Then
            Let Result.Mask = MASK_B: Let Result.Value = OPCODE_B
        ElseIf Result.Register = TOKEN_REGISTER_C Then
            Let Result.Mask = MASK_C: Let Result.Value = OPCODE_C
        ElseIf Result.Register = TOKEN_REGISTER_D Then
            Let Result.Mask = MASK_D: Let Result.Value = OPCODE_D
        ElseIf Result.Register = TOKEN_REGISTER_E Then
            Let Result.Mask = MASK_E: Let Result.Value = OPCODE_E
        ElseIf Result.Register = TOKEN_REGISTER_H Then
            Let Result.Mask = MASK_H: Let Result.Value = OPCODE_H
        ElseIf Result.Register = TOKEN_REGISTER_L Then
            Let Result.Mask = MASK_L: Let Result.Value = OPCODE_L
        ElseIf Result.Register = TOKEN_REGISTER_AF Then
            Let Result.Mask = MASK_AF: Let Result.Value = OPCODE_AF
        ElseIf Result.Register = TOKEN_REGISTER_BC Then
            Let Result.Mask = MASK_BC: Let Result.Value = OPCODE_BC
        ElseIf Result.Register = TOKEN_REGISTER_DE Then
            Let Result.Mask = MASK_DE: Let Result.Value = OPCODE_DE
        ElseIf Result.Register = TOKEN_REGISTER_HL Then
            Let Result.Mask = MASK_HL: Let Result.Value = OPCODE_HL
        ElseIf Result.Register = TOKEN_REGISTER_SP Then
            Let Result.Mask = MASK_SP: Let Result.Value = OPCODE_SP
        ElseIf Result.Register = TOKEN_REGISTER_IX Then
            Let Result.Mask = MASK_IX: Let Result.Value = OPCODE_IX
        ElseIf Result.Register = TOKEN_REGISTER_IY Then
            Let Result.Mask = MASK_IY: Let Result.Value = OPCODE_IY
        
        'Condition flags:
        ElseIf Result.Register = TOKEN_FLAG_NZ Then
            Let Result.Mask = MASK_FLAGS_CZ: Let Result.Value = OPCODE_FLAG_NZ
        ElseIf Result.Register = TOKEN_FLAG_Z Then
            Let Result.Mask = MASK_FLAGS_CZ: Let Result.Value = OPCODE_FLAG_Z
        ElseIf Result.Register = TOKEN_FLAG_NC Then
            Let Result.Mask = MASK_FLAGS_CZ: Let Result.Value = OPCODE_FLAG_NC
        'NOTE: Flag C is not included because it cannot be diferentiated from _
               Register C. The parent will fix this limitation for us
        ElseIf Result.Register = TOKEN_FLAG_PO Then
            Let Result.Mask = MASK_FLAGS_MP: Let Result.Value = OPCODE_FLAG_PO
        ElseIf Result.Register = TOKEN_FLAG_PE Then
            Let Result.Mask = MASK_FLAGS_MP: Let Result.Value = OPCODE_FLAG_PE
        ElseIf Result.Register = TOKEN_FLAG_P Then
            Let Result.Mask = MASK_FLAGS_MP: Let Result.Value = OPCODE_FLAG_P
        ElseIf Result.Register = TOKEN_FLAG_M Then
            Let Result.Mask = MASK_FLAGS_MP: Let Result.Value = OPCODE_FLAG_M
        End If
        'Now move beyond the register token
        Call NextToken
    
    'If the token is a memory reference, such as `(ix+$8)`, _
     then get the associated register and value
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PARENOPEN Then '..............................
        Let GetParameter = ScopeMemory( _
            Result.Register, Result.Value, Result.ValueUnknown _
        )
        'TODO: Error - the memory reference is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
        
        'Set which register is used in the memory reference
        If Result.Register = TOKEN_REGISTER_HL Then
            Let Result.Mask = MASK_MEMHL
        ElseIf Result.Register = TOKEN_REGISTER_IX Then
            Let Result.Mask = MASK_MEMIX8
        ElseIf Result.Register = TOKEN_REGISTER_IY Then
            Let Result.Mask = MASK_MEMIY8
        ElseIf Result.Register = TOKEN_REGISTER_SP Then
            Let Result.Mask = MASK_MEMSP
        ElseIf Result.Register = TOKEN_REGISTER_C Then
            Let Result.Mask = MASK_MEMC
        End If
        
    'Anything else classes as an expression (number / labels / variables / operators)
    ElseIf IsExpression() Then '........................................................
        Let Result.Kind = PARAM_NUMBER
        Let GetParameter = ScopeExpression(Result.Value, Result.ValueUnknown)
        'Check if it's in 8-bit range and 16-bit range
        If (Result.Value And &HFF) = Result.Value Then _
            Let Result.Mask = Result.Mask Or MASK_8
        If (Result.Value And &HFF00) = Result.Value Then _
            Let Result.Mask = Result.Mask Or MASK_16
        'TODO: Error - the expression is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
    
    Else
        'ERROR: Not an expression
        Let GetParameter = RaiseError(OZ80_ERROR_EXPRESSION, _
            "Z80 instruction parameter expected. Not a valid parameter expression.", _
            TokenFile.Line, TokenFile.Col _
        )
    End If
End Function

'IsParameter : Check if the current token would be considered a Z80 parameter _
 ======================================================================================
Private Function IsParameter() As Boolean
    'A parameter can be a number, label/variable/property, register _
     or memory reference
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'A label on a new line is a definition, not a parameter
        If TokenFile.Line = PrevLine Then Let IsParameter = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_NUMBER Then
        Let IsParameter = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PARENOPEN Then
        Let IsParameter = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PROPERTY Then
        'A property on a new line is a definition, not a parameter
        If TokenFile.Line = PrevLine Then Let IsParameter = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
        Let IsParameter = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_VARIABLE Then
        Let IsParameter = True
    End If
End Function

'IsExpression : Check if the current token would be considered an expression _
 ======================================================================================
Private Function IsExpression() As Boolean
    'A parameter can be a number or label/variable/property
    If TokenFile.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        Let IsExpression = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_NUMBER Then
        Let IsExpression = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_PROPERTY Then
        Let IsExpression = True
    ElseIf TokenFile.Kind = OZ80_TOKEN.TOKEN_VARIABLE Then
        Let IsExpression = True
    End If
End Function

'NextToken : Move to the next token in the token stream _
 ======================================================================================
Private Function NextToken() As Boolean
    'This function is a mere wrapper simply because we need to track when a new line _
     occurs between tokens to diferentiate label definitions and parameters
    Let PrevLine = TokenFile.Line
    Let NextToken = TokenFile.Forward()
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Function RaiseError( _
    ByRef Number As OZ80_ERROR, _
    ByRef Description As String, _
    Optional ByRef Line As Long = 0, Optional ByRef Col As Long = 0 _
) As OZ80_ERROR
    
    RaiseEvent Error(Number, Description, Line, Col)
    
    'We return the same error number we were given only to make the tail recursion _
     simpler, e.g. `let ScopeRoot = RaiseError(OZ80_ERROR_...`
    Let RaiseError = Number
End Function

'TestOpcode : Outputs an opcode if the parameter has values within the given mask _
 ======================================================================================
Private Function TestOpcode( _
    ByRef Opcode As Integer, ByRef Param As oz80Param, ByRef Test As OZ80_MASK _
) As Boolean
    If (Param.Mask And Test) <> 0 Then
        Call AddOpcode(Opcode, Param.Mask, Param.Value)
        Let TestOpcode = True
    End If
End Function
