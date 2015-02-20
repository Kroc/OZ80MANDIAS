VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OZ80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'OZ80MANDIAS: a Z80 assembler; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: OZ80

'In any normal programming language you would want lots of small classes to handle _
 something as big and complex as parsing and assembling. Class dereferencing (that is, _
 crossing the boundary into the class for a variable containing an instance) can be _
 a little slow in VB6. For maximum speed, OZ80 forgoes almost any 'sane' architecture _
 and does everything "flat" -- using data structures instead of classes
 
'Therefore the code in here is very, very long. I've tried to arrange it in logical _
 progression order -- source code parsing & tokenising, assembling, and output

'/// DEBUG ////////////////////////////////////////////////////////////////////////////

Private Profiler As New bluProfiler

Private Enum PROFILER_EVENTS
    EVENT_INITIALIZE                    'Class initialisation -- caching
    EVENT_TOKENISE                      'Tokenise File
    EVENT_TOKENISE_READWORD             '- Parse out a single word
    EVENT_TOKENISE_TOKENWORD            '- Tokenise a single word
    EVENT_FORMATTOKEN                   '- Format token data for logging
    EVENT_PROCESSZ80                    'ProcessZ80
End Enum

'/// API //////////////////////////////////////////////////////////////////////////////

'Undocumented VB6 APIs to write a number of bytes to memory. _
 We'll be using this to manually create some Double values
Private Declare Sub api_PutMem1 Lib "msvbvm60" Alias "PutMem1" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Byte _
)
Private Declare Sub api_PutMem2 Lib "msvbvm60" Alias "PutMem2" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Integer _
)
Private Declare Sub api_PutMem4 Lib "msvbvm60" Alias "PutMem4" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Long _
)

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub api_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal DestinationPointer As Long, _
    ByVal SourcePointer As Long, _
    ByVal Length As Long _
)

'--------------------------------------------------------------------------------------

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Structure for obtaining the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724834%28v=vs.85%29.aspx>
Private Type OSVERSIONINFO
    SizeOfMe As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformID As Long
    ServicePack As String * 128
End Type

'Get the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724451%28v=vs.85%29.aspx>
Private Declare Function api_GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    ByRef VersionInfo As OSVERSIONINFO _
) As BOOL

'--------------------------------------------------------------------------------------

'Get the Locale Identifier (LCID) of this app _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318127(v=vs.85).aspx>
'This is used for Windows XP support as Vista+ use Locale Name strings
Private Declare Function api_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" ( _
) As Long

'Unicode & Locale-aware case conversion (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318700(v=vs.85).aspx>
Private Declare Function api_LCMapString Lib "kernel32" Alias "LCMapStringW" ( _
    ByVal LocaleID As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long _
) As Long

'Unicode & Locale-aware case conversion (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx>
Private Declare Function api_LCMapStringEx Lib "kernel32" Alias "LCMapStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long, _
    ByVal VersionInfoPointer As Long, _
    ByVal ReservedPointer As Long, _
    ByVal SortHandle As Long _
) As Long

Private Enum LCMAP
    LCMAP_LINGUISTIC_CASING = &H1000000 'Manage accents &c. in lower/upper casing
    LCMAP_LOWERCASE = &H100&            'Convert to lower case
    LCMAP_UPPERCASE = &H200&            'Convert to UPPER CASE
    LCMAP_SORTKEY = &H400&              'Generate a sort key
    LCMAP_BYTEREV = &H800&              'Reverse the bytes
    LCMAP_HIRAGANA = &H100000           'Change Japanese Katakana to Hiragana
    LCMAP_KATAKANA = &H200000           'Change Japanese Hiragana to Katakana
    LCMAP_HALFWIDTH = &H400000          'Use standard (half-width) characters
    LCMAP_FULLWIDTH = &H800000          'Use wide-characters (i.e. Latin in Chinese)
    'Change traditional Chinese to simplified Chinese and vice-versa
    LCMAP_SIMPLIFIED_CHINESE = &H2000000
    LCMAP_TRADITIONAL_CHINESE = &H4000000
End Enum

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'Error number!
Private My_Error As OZ80_ERROR

Private WindowsVersion

'During class initialisation we manually create special 64-bit numbers _
 in these values, which cannot be written out natively in VB6
Private INDEFINITE As Double
Private NUMBER_MAX As Double
Private NUMBER_MIN As Double
'The min/max numbers that can be multiplied/divided by 1'000 or 1'024 without overflow
Private NUMBER_MIN_K As Double
Private NUMBER_MAX_K As Double
Private NUMBER_MIN_KB As Double
Private NUMBER_MAX_KB As Double

'--------------------------------------------------------------------------------------

Public Enum OZ80_WARNING
    OZ80_WARNING_NONE                   'Skip "0"
    OZ80_WARNING_ROUND                  'A decimal number had to be round-down
End Enum

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_DUPLICATE_CONSTANT       '- Constant already defined
    OZ80_ERROR_DUPLICATE_LABEL          '- Duplicate Label
    OZ80_ERROR_DUPLICATE_PROC_HELP      '- Duplicate `HELP` parameter
    OZ80_ERROR_DUPLICATE_PROC_INTERRUPT '- Duplicate `INTERRUPT` parameter
    OZ80_ERROR_DUPLICATE_PROC_PARAMS    '- Duplicate `PARAMS` parameter
    OZ80_ERROR_DUPLICATE_PROC_RETURN    '- Duplicate `RETURN` parameter
    OZ80_ERROR_DUPLICATE_PROC_SECTION   '- Duplicate `SECTION` parameter
    OZ80_ERROR_DUPLICATE_SECTION        '- Can't define a Section twice
    OZ80_ERROR_DUPLICATE_START          '- Duplicate `START` Procedure
    OZ80_ERROR_DUPLICATE_TABLE_SECTION  '- Duplicate `SECTION` parameter
    OZ80_ERROR_EXPECTED                 'Incorrect content at the current scope
    OZ80_ERROR_EXPECTED_BRACKET         '- Close bracket ("}","]",")") without open
    OZ80_ERROR_EXPECTED_EXPRESSION      '- Expression required here
    OZ80_ERROR_EXPECTED_PROC_NAME       '- A label name must follow `PROC`
    OZ80_ERROR_EXPECTED_PROC_PARAMS     '- Invalid stuff in the `PARAMS` list
    OZ80_ERROR_EXPECTED_PROC_RETURN     '- Invalid stuff in the `RETURN` list
    OZ80_ERROR_EXPECTED_ROOT            '- Only certain keywords allowed at root
    OZ80_ERROR_EXPECTED_SECTION_NAME    '- A section name must follow `SECTION`
    OZ80_ERROR_EXPECTED_SYSTEM_NAME     '- A system name must follow `SYSTEM`
    OZ80_ERROR_EXPECTED_TABLE_NAME      '- A label name must follow `TABLE`
    OZ80_ERROR_EXPRESSION               'Not a valid expression
    OZ80_ERROR_EXPRESSION_Z80           '- Not a valid Z80 instruction parameter
    OZ80_ERROR_FILE_END                 'Unexpected end of file
    OZ80_ERROR_FILE_NOTFOUND            'Requested file does not exist
    OZ80_ERROR_FILE_READ                'Some kind of problem while file handle open
    OZ80_ERROR_INDEFINITE               'Indefinite value cannot be used here
    OZ80_ERROR_INVALID_INTERRUPT        'Invalid Interrupt address
    OZ80_ERROR_INVALID_NAME             'Invalid label/property/variable name
    OZ80_ERROR_INVALID_NAME_RAM         '- Invalid RAM name, i.e. `$.name`
    OZ80_ERROR_INVALID_NAME_HASH        '- Invalid hash name, i.e. `#hash`
    OZ80_ERROR_INVALID_NUMBER           'Not a valid number
    OZ80_ERROR_INVALID_NUMBER_DEC       '- Invalid decimal number
    OZ80_ERROR_INVALID_NUMBER_HEX       '- Invalid hexadecimal number
    OZ80_ERROR_INVALID_NUMBER_BIN       '- Invalid binary number
    OZ80_ERROR_INVALID_PROC_INTERRUPT   'SECTION & INTERRUPT params cannot co-exist
    OZ80_ERROR_INVALID_SECTION          'Section used, but not defined
    OZ80_ERROR_INVALID_SLOT             'Incorrect use of the Slot parameter
    OZ80_ERROR_INVALID_WORD             'Couldn't parse a word
    OZ80_ERROR_INVALID_Z80PARAMS        'Not the right parameters for a Z80 instruction
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_OVERFLOW_HILO            '- HI & LO functions limited to 16-bit inputs
    OZ80_ERROR_OVERFLOW_LINE            '- Line too long
    OZ80_ERROR_OVERFLOW_FILE            '- File too long / large
    OZ80_ERROR_OVERFLOW_Z80             '- 16-bit number used with an 8-bit instruction
    OZ80_ERROR_TEXT_CHAR                'Character code out of range
    OZ80_ERROR_UNDEFINED                'Named item is used, but undefined
    OZ80_ERROR_UNDEFINED_CONST          '- A Constant has been used before definition
End Enum

'Logging: _
 --------------------------------------------------------------------------------------

Public Enum OZ80_LOG
    OZ80_LOG_ACTION                     'The key important happenings
    OZ80_LOG_INFO                       'Optional information, not actions happening
    OZ80_LOG_STATUS                     'Display variable values &c. when assigned
    OZ80_LOG_DEBUG                      'Internal information for debugging purposes
End Enum

'Formatting of output messages
Private Const LOG_VERSION               As String = "OZ80MANDIAS v{$}.{$},{$}"
Private Const LOG_PHASE1                As String = "Phase 1: Structural Pass"
Private Const LOG_PHASE2                As String = "Phase 2: Layout ROM"
Private Const LOG_PHASE3                As String = "Phase 3: Full Assemble"
Private Const LOG_PHASE4                As String = "Phase 4: Post Processing"
Private Const LOG_ASM_COMPLETE          As String = "Assembly Complete"
Private Const LOG_TIME                  As String = "Time: {$}s"
Private Const LOG_SECTION               As String = "SECTION {$}"
Private Const LOG_LINEDOUBLE            As String = "========================================================================="
Private Const LOG_FILE_LENGTH           As String = "Length: {$} Characters"
Private Const LOG_FILE_TOKENS           As String = "Tokens: {$}"

'For diaplying a disassembly

Private Const LOG_ASM_COLS              As String = "ADDR | PX OP OF CB LO HI | MNEMONIC  PARAM1    PARAM2    PARAM3"
Private Const LOG_ASM_HEAD              As String = "-----|-------------------|-----------------------------------------------"
Private Const LOG_ASM                   As String = "···· | ·· ·· ·· ·· ·· ·· | ·········                                     "

Private Const LOG_DATA                  As String = "{$} | {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} {$} | {$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}{$}"

Private Const FORMAT_QUOTE              As String = """"
Private Const FORMAT_TOKEN              As String = "<tkn> {$} ({$},{$}) [{$}] {$}"
Private Const FORMAT_TOKEN_CHUNK        As String = "   CHUNK"
Private Const FORMAT_TOKEN_CONST        As String = "   CONST"
Private Const FORMAT_TOKEN_EXPR         As String = "    EXPR"
Private Const FORMAT_TOKEN_FLAG         As String = "    FLAG"
Private Const FORMAT_TOKEN_KEYWORD      As String = " KEYWORD"
Private Const FORMAT_TOKEN_LABEL        As String = "   LABEL"
Private Const FORMAT_TOKEN_NUMBER       As String = "  NUMBER"
Private Const FORMAT_TOKEN_OPERATOR     As String = "OPERATOR"
Private Const FORMAT_TOKEN_PREFIX       As String = "  PREFIX"
Private Const FORMAT_TOKEN_PROPERTY     As String = "PROPERTY"
Private Const FORMAT_TOKEN_RAM          As String = "     RAM"
Private Const FORMAT_TOKEN_REGISTER     As String = "REGISTER"
Private Const FORMAT_TOKEN_SECTION      As String = " SECTION"
Private Const FORMAT_TOKEN_SQUARE       As String = "HASH/MEM"
Private Const FORMAT_TOKEN_TEXT         As String = "    TEXT"
Private Const FORMAT_TOKEN_Z80          As String = "     Z80"

'To save having to create and destroy short strings all the time, _
 we'll keep this one around for formatting log messages
Private LogMsg As bluString

'Which type of data the Chunk is recording (Z80 assembly or Data). _
 This is used to produce the disassembly log
Private ChunkMode As CHUNK_MODE
Private Enum CHUNK_MODE
    MODE_UNSET
    MODE_Z80                            'When adding Z80 instructions
    MODE_DATA                           'When adding data
End Enum

'When generating the disassembly log, we don't show data as 1-byte per line, _
 but render it in typical bin-hex format of 16 hex-bytes followed by ASCII. _
 We therefore need to remember up to 16 bytes as we go before we can output this
Private ChunkData(0 To 15) As String
Private ChunkDataIndex As Long
'Log-ready address of where the line of data begins
Private ChunkDataAddr As String * 4

'--------------------------------------------------------------------------------------
'File Tokenising:
'--------------------------------------------------------------------------------------

'A lot of the text parsing will be letter-by-letter, so this enum provides the _
 common character codes we'll be referring to, as well as the particular punctuation _
 symbols used for OZ80 syntax. Note that `UCS2` is defined by 'bluString.bas'
Private Enum OZ80_SYNTAX
    ASC_NUL
    ASC_TAB = UCS2.C0_HT
    ASC_LF = UCS2.C0_LF
    ASC_CR = UCS2.C0_CR
    
    ASC_SPC = UCS2.SPACE_
    ASC_USC = UCS2.LOW_LINE
    
    ASC_A = UCS2.LATIN_CAPITAL_LETTER_A
:   ASC_B:      ASC_C:      ASC_D:      ASC_E:      ASC_F:      ASC_G:      ASC_H
:   ASC_I:      ASC_J:      ASC_K:      ASC_L:      ASC_M:      ASC_N:      ASC_O
:   ASC_P:      ASC_Q:      ASC_R:      ASC_S:      ASC_T:      ASC_U:      ASC_V
:   ASC_W:      ASC_X:      ASC_Y:      ASC_Z

    ASC_0 = UCS2.DIGIT_ZERO
:   ASC_1:      ASC_2:      ASC_3:      ASC_4:      ASC_5:      ASC_6:      ASC_7
:   ASC_8:      ASC_9
    
    ASC_NEG = UCS2.HYPHEN_MINUS
    
    SYNTAX_COMMENT = UCS2.GRAVE_ACCENT              ' ` - Comment marker
    SYNTAX_HINT1 = UCS2.SEMICOLON                   ' ; - Register hint, `a;index`
    SYNTAX_HINT2 = UCS2.APOSTROPHE                  ' ' - Shadow register, `ex af 'af`
    SYNTAX_TEXT = UCS2.QUOTATION_MARK               ' " - String identifier
    SYNTAX_CONSTANT = UCS2.EXCLAMATION_MARK         ' ! - Constant identifier
    SYNTAX_LABEL = UCS2.COLON                       ' : - Label identifier
    SYNTAX_PROPERTY = UCS2.FULL_STOP                ' . - Property identifier
    SYNTAX_HASH = UCS2.NUMBER_SIGN                  ' # - Hash identifier
    SYNTAX_RAM = UCS2.DOLLAR_SIGN                   ' $ - RAM constant, "$.abc"
    SYNTAX_MACRO = UCS2.COMMERCIAL_AT               ' @ - Macro identifier
    SYNTAX_FUNCT = UCS2.QUESTION_MARK               ' ? - Function identifier
    SYNTAX_NUMBER_HEX = UCS2.DOLLAR_SIGN            ' $ - Hexadecimal number, `$FFFF`
    SYNTAX_NUMBER_BIN = UCS2.PERCENT_SIGN           ' % - Binary number, `%10101011`
    SYNTAX_NEXT = UCS2.COMMA                        ' , - Item seperator, optional
    SYNTAX_PARENS_OPEN = UCS2.LEFT_PARENTHESIS      ' ( - Open parentheses
    SYNTAX_PARENS_CLOSE = UCS2.RIGHT_PARENHESIS     ' ) - Close parentheses
    SYNTAX_BRACES_OPEN = UCS2.LEFT_CURLY_BRACKET    ' { - Open brace
    SYNTAX_BRACES_CLOSE = UCS2.RIGHT_CURLY_BRACKET  ' } - Close brace
    SYNTAX_SQUARE_OPEN = UCS2.LEFT_SQUARE_BRACKET   ' [ - Open bracket
    SYNTAX_SQUARE_CLOSE = UCS2.RIGHT_SQUARE_BRACKET ' ] - Close bracket
    SYNTAX_OPERATOR_ADD = UCS2.PLUS_SIGN            ' + - Add
    SYNTAX_OPERATOR_SUB = UCS2.HYPHEN_MINUS         ' - - Subtract
    SYNTAX_OPERATOR_MUL = UCS2.ASTERISK             ' * - Multiply
    SYNTAX_OPERATOR_DIV = UCS2.SOLIDUS              ' / - Divide
    SYNTAX_OPERATOR_POW = UCS2.CIRCUMFLEX_ACCENT    ' ^ - Power
    SYNTAX_OPERATOR_MOD = UCS2.REVERSE_SOLIDUS      ' \ - Modulus
    SYNTAX_OPERATOR_OR = UCS2.VERTICAL_LINE         ' | - Bitwise OR
    SYNTAX_OPERATOR_AND = UCS2.AMPERSAND            ' & - Bitwise AND
    SYNTAX_OPERATOR_XOR = UCS2.TILDE                ' ~ - Bitwise XOR
End Enum

'The width of a tab stop (this affects the column number given)
'TODO: This should be a class property so that the user can customise
Private Const TAB_WIDTH As Long = 8

'During parsing, we will need to do case-insensitive character comparisons, _
 upper-casing the token word is one of the biggest bottlenecks in the process! _
 Therefore we build a lookup table for upper-casing individual characters
Private ToUpperCase(0 To 65535) As Integer

'String constants at class level so that these texts do not have to be temporarily _
 created and thrown away every time we use them

Private Const STR_A                     As String = "A"
Private Const STR_ADC                   As String = "adc"
Private Const STR_ADD                   As String = "add"
Private Const STR_AF                    As String = "AF"
Private Const STR_AND                   As String = "and"
Private Const STR_B                     As String = "B"
Private Const STR_BC                    As String = "BC"
Private Const STR_BIT                   As String = "bit"
Private Const STR_BOOL                  As String = "BOOL"
Private Const STR_BYTE                  As String = "BYTE"
Private Const STR_C                     As String = "C"
Private Const STR_CALL                  As String = "call"
Private Const STR_CCF                   As String = "ccf"
Private Const STR_CP                    As String = "cp"
Private Const STR_CPD                   As String = "cpd"
Private Const STR_CPDR                  As String = "cpdr"
Private Const STR_CPI                   As String = "cpi"
Private Const STR_CPIR                  As String = "cpir"
Private Const STR_CPL                   As String = "cpl"
Private Const STR_DAA                   As String = "daa"
Private Const STR_DEC                   As String = "dec"
Private Const STR_D                     As String = "D"
Private Const STR_DE                    As String = "DE"
Private Const STR_DEF                   As String = "DEF"
Private Const STR_DI                    As String = "di"
Private Const STR_DJNZ                  As String = "djnz"
Private Const STR_E                     As String = "E"
Private Const STR_EI                    As String = "ei"
Private Const STR_EX                    As String = "ex"
Private Const STR_EXX                   As String = "exx"
Private Const STR_F                     As String = "F"
Private Const STR_HLT                   As String = "hlt"
Private Const STR_HALT                  As String = "halt"
Private Const STR_HASH                  As String = "HASH"
Private Const STR_HELP                  As String = "HELP"
Private Const STR_H                     As String = "H"
Private Const STR_HI                    As String = "HI"
Private Const STR_HL                    As String = "HL"
Private Const STR_I                     As String = "I"
Private Const STR_IM                    As String = "im"
Private Const STR_IN                    As String = "in"
Private Const STR_INC                   As String = "inc"
Private Const STR_IND                   As String = "ind"
Private Const STR_INDR                  As String = "indr"
Private Const STR_INI                   As String = "ini"
Private Const STR_INIR                  As String = "inir"
Private Const STR_INCLUDE               As String = "INCLUDE"
Private Const STR_INDEX                 As String = "INDEX"
Private Const STR_INTERRUPT             As String = "INTERRUPT"
Private Const STR_IX                    As String = "IX"
Private Const STR_IXH                   As String = "IXH"
Private Const STR_IXL                   As String = "IXL"
Private Const STR_IY                    As String = "IY"
Private Const STR_IYH                   As String = "IYH"
Private Const STR_IYL                   As String = "IYL"
Private Const STR_JP                    As String = "jp"
Private Const STR_JR                    As String = "jr"
Private Const STR_K                     As String = "K"
Private Const STR_KB                    As String = "KB"
Private Const STR_L                     As String = "L"
Private Const STR_LD                    As String = "ld"
Private Const STR_LDD                   As String = "ldd"
Private Const STR_LDDR                  As String = "lddr"
Private Const STR_LDI                   As String = "ldi"
Private Const STR_LDIR                  As String = "ldir"
Private Const STR_LO                    As String = "LO"
Private Const STR_LONG                  As String = "LONG"
Private Const STR_M                     As String = "M"
Private Const STR_NC                    As String = "NC"
Private Const STR_NZ                    As String = "NZ"
Private Const STR_NEG                   As String = "neg"
Private Const STR_NOP                   As String = "nop"
Private Const STR_NYBL                  As String = "NYBL"
Private Const STR_OR                    As String = "or"
Private Const STR_OUT                   As String = "out"
Private Const STR_OUTD                  As String = "outd"
Private Const STR_OTDR                  As String = "otdr"
Private Const STR_OUTI                  As String = "outi"
Private Const STR_OTIR                  As String = "otir"
Private Const STR_P                     As String = "P"
Private Const STR_PARAMS                As String = "PARAMS"
Private Const STR_PC                    As String = "PC"
Private Const STR_PE                    As String = "PE"
Private Const STR_PO                    As String = "PO"
Private Const STR_POP                   As String = "pop"
Private Const STR_PROC                  As String = "PROC"
Private Const STR_PUSH                  As String = "push"
Private Const STR_R                     As String = "R"
Private Const STR_RAM                   As String = "RAM"
Private Const STR_RES                   As String = "res"
Private Const STR_RET                   As String = "ret"
Private Const STR_RETI                  As String = "reti"
Private Const STR_RETN                  As String = "retn"
Private Const STR_RETURN                As String = "RETURN"
Private Const STR_RL                    As String = "rl"
Private Const STR_RLA                   As String = "rla"
Private Const STR_RLC                   As String = "rlc"
Private Const STR_RLCA                  As String = "rlca"
Private Const STR_RLD                   As String = "rld"
Private Const STR_RR                    As String = "rr"
Private Const STR_RRA                   As String = "rra"
Private Const STR_RRC                   As String = "rrc"
Private Const STR_RRCA                  As String = "rrca"
Private Const STR_RRD                   As String = "rrd"
Private Const STR_RST                   As String = "rst"
Private Const STR_SBC                   As String = "sbc"
Private Const STR_SECTION               As String = "SECTION"
Private Const STR_SCF                   As String = "scf"
Private Const STR_SET                   As String = "set"
Private Const STR_SLA                   As String = "sla"
Private Const STR_SLL                   As String = "sll"
Private Const STR_SLOT                  As String = "SLOT"
Private Const STR_SMS                   As String = "SMS"
Private Const STR_SP                    As String = "SP"
Private Const STR_SRA                   As String = "sra"
Private Const STR_SRL                   As String = "srl"
Private Const STR_START                 As String = "START"
Private Const STR_SUB                   As String = "sub"
Private Const STR_SYSTEM                As String = "SYSTEM"
Private Const STR_TABLE                 As String = "TABLE"
Private Const STR_TRIP                  As String = "TRIP"
Private Const STR_WORD                  As String = "WORD"
Private Const STR_XOR                   As String = "xor"
Private Const STR_Z                     As String = "Z"

'--------------------------------------------------------------------------------------

Private Enum OZ80_TOKEN
    TOKEN_NONE                          'Skip "0"
    
    [_TOKEN_FIRST]
    [_TOKEN_Z80_BEGIN] = [_TOKEN_FIRST]
    
    'These are just the mnemonic tokens -- the assembler checks the
     'parameters and determines which opcode should be used
    TOKEN_Z80_ADC = [_TOKEN_FIRST]      'Add with Carry
    TOKEN_Z80_ADD                       'Add
    TOKEN_Z80_AND                       'Bitwise AND
    TOKEN_Z80_BIT                       'Bit test
    TOKEN_Z80_CALL                      'Call routine
    TOKEN_Z80_CCF                       'Clear Carry Flag
    TOKEN_Z80_CP                        'Compare
    TOKEN_Z80_CPD                       'Compare and Decrement
    TOKEN_Z80_CPDR                      'Compare, Decrement and Repeat
    TOKEN_Z80_CPI                       'Compare and Increment
    TOKEN_Z80_CPIR                      'Compare, Increment and Repeat
    TOKEN_Z80_CPL                       'Complement (bitwise NOT)
    TOKEN_Z80_DAA                       'Decimal Adjust Accumulator
    TOKEN_Z80_DEC                       'Decrement
    TOKEN_Z80_DI                        'Disable Interrupts
    TOKEN_Z80_DJNZ                      'Decrement and Jump if Not Zero
    TOKEN_Z80_EI                        'Enable Inettupts
    TOKEN_Z80_EX                        'Exchange
    TOKEN_Z80_EXX                       'Exchange shadow registers
    TOKEN_Z80_HALT                      'Stop CPU (wait for interrupt)
    TOKEN_Z80_IM                        'Interrupt Mode
    TOKEN_Z80_IN                        'Input from port
    TOKEN_Z80_INC                       'Increment
    TOKEN_Z80_IND                       'Input and Decrement
    TOKEN_Z80_INDR                      'Input, Decrement and Repeat
    TOKEN_Z80_INI                       'Input and Increment
    TOKEN_Z80_INIR                      'Input, Increment and Repeat
    TOKEN_Z80_JP                        'Jump
    TOKEN_Z80_JR                        'Jump Relative
    TOKEN_Z80_LD                        'Load
    TOKEN_Z80_LDD                       'Load and Decrement
    TOKEN_Z80_LDDR                      'Load, Decrement and Repeat
    TOKEN_Z80_LDI                       'Load and Increment
    TOKEN_Z80_LDIR                      'Load, Increment and Repeat
    TOKEN_Z80_NEG                       'Negate (flip the sign)
    TOKEN_Z80_NOP                       'No Operation (do nothing)
    TOKEN_Z80_OR                        'Bitwise OR
    TOKEN_Z80_OUT                       'Output to port
    TOKEN_Z80_OUTD                      'Output and Decrement
    TOKEN_Z80_OTDR                      'Output, Decrement and Repeat
    TOKEN_Z80_OUTI                      'Output and Increment
    TOKEN_Z80_OTIR                      'Output, Increment and Repeat
    TOKEN_Z80_POP                       'Pull from stack
    TOKEN_Z80_PUSH                      'Push onto stack
    TOKEN_Z80_RES                       'Reset bit
    TOKEN_Z80_RET                       'Return from routine
    TOKEN_Z80_RETI                      'Return from Interrupt
    TOKEN_Z80_RETN                      'Return from NMI
    TOKEN_Z80_RLA                       'Rotate Left (Accumulator)
    TOKEN_Z80_RL                        'Rotate Left
    TOKEN_Z80_RLCA                      'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLC                       'Rotate Left Circular
    TOKEN_Z80_RLD                       'Rotate Left 4-bits
    TOKEN_Z80_RRA                       'Rotate Right (Accumulator)
    TOKEN_Z80_RR                        'Rotate Right
    TOKEN_Z80_RRCA                      'Rotate Right Circular (Accumulator)
    TOKEN_Z80_RRC                       'Rotate Right Circular
    TOKEN_Z80_RRD                       'Rotate Right 4-bits
    TOKEN_Z80_RST                       '"Restart" -- Call a page 0 routine
    TOKEN_Z80_SBC                       'Subtract with Carry
    TOKEN_Z80_SCF                       'Set Carry Flag
    TOKEN_Z80_SET                       'Set bit
    TOKEN_Z80_SLA                       'Shift Left Arithmetic
    TOKEN_Z80_SRA                       'Shift Right Arithmetic
    TOKEN_Z80_SLL                       'Shift Left Logical
    TOKEN_Z80_SRL                       'Shift Right Logical
    TOKEN_Z80_SUB                       'Subtract
    TOKEN_Z80_XOR                       'Bitwise XOR
    [_TOKEN_Z80_END] = TOKEN_Z80_XOR
    
    'Z80 Registers & Flags ............................................................
    [_TOKEN_REGS_BEGIN]
    TOKEN_Z80_A = [_TOKEN_REGS_BEGIN]   'Accumulator
    TOKEN_Z80_AF                        'Accumulator and Flags
    TOKEN_Z80_B                         'Register B
    TOKEN_Z80_BC                        'Register pair B & C
    TOKEN_Z80_C                         'Register C or Carry flag
    TOKEN_Z80_D                         'Register D
    TOKEN_Z80_DE                        'Register pair D & E
    TOKEN_Z80_E                         'Register E
    TOKEN_Z80_H                         'Register H
    TOKEN_Z80_HL                        'Register pair H & L
    TOKEN_Z80_I                         'Interrupt - not to be confused with IX & IY
    TOKEN_Z80_IX                        'Register IX
    TOKEN_Z80_IXH                       'Undocumented high-byte of register IX
    TOKEN_Z80_IXL                       'Undocumented low-byte of register IX
    TOKEN_Z80_IY                        'Register IY
    TOKEN_Z80_IYH                       'Undocumented high-byte of register IY
    TOKEN_Z80_IYL                       'Undocumented low-byte of register IY
    TOKEN_Z80_L                         'Register L
    TOKEN_Z80_M                         'Sign is set flag
    TOKEN_Z80_NC                        'Carry unset flag
    TOKEN_Z80_NZ                        'Zero not set flag
    TOKEN_Z80_P                         'Sign is not set flag
    TOKEN_Z80_PC                        'Program Counter
    TOKEN_Z80_PE                        'Parity/Overflow is set flag
    TOKEN_Z80_PO                        'Parity/Overflow is not set flag
    TOKEN_Z80_R                         'Refresh register (pseudo-random)
    TOKEN_Z80_SP                        'Stack Pointer
    TOKEN_Z80_Z                         'Zero set flag
    [_TOKEN_REGS_END] = TOKEN_Z80_Z
    
    'Operators ........................................................................
    [_TOKEN_OPERATORS_BEGIN]
    TOKEN_OPERATOR_ADD                  'Add "+"
    TOKEN_OPERATOR_SUB                  'Subtract "-"
    TOKEN_OPERATOR_MUL                  'Multiply "*"
    TOKEN_OPERATOR_DIV                  'Divide "/"
    TOKEN_OPERATOR_POW                  'Power "^"
    TOKEN_OPERATOR_MOD                  'Modulus "\"
    TOKEN_OPERATOR_REP                  'Repeat "x"
    TOKEN_OPERATOR_OR                   'Bitwise OR "|"
    TOKEN_OPERATOR_AND                  'Bitwise AND "&"
    TOKEN_OPERATOR_NOT                  'Bitwise NOT "!"
    TOKEN_OPERATOR_XOR                  'Bitwise XOR "~"
    [_TOKEN_OPERATORS_END]
    
    'Keywords .........................................................................
    [_TOKEN_KEYWORDS_BEGIN]
    TOKEN_KEYWORD_BOOL                  'Boolean data type (1-bit)
    TOKEN_KEYWORD_BYTE                  'Byte data type
    TOKEN_KEYWORD_DEF                   'Define constant
    TOKEN_KEYWORD_HASH                  'Define a hash-array
    TOKEN_KEYWORD_HELP                  'Documentation marker
    TOKEN_KEYWORD_HI                    '`HI` function -- high byte of 16-bit Value
    TOKEN_KEYWORD_INCLUDE               'Include another file
    TOKEN_KEYWORD_INDEX                 'Create an index number for a Label
    TOKEN_KEYWORD_INTERRUPT             'Interrupt `PROC :<label> INTERRUPT <expr>`
    TOKEN_KEYWORD_LO                    '`LO` function -- low byte of a 16-bit Value
    TOKEN_KEYWORD_LONG                  'Long data type (4-bytes)
    TOKEN_KEYWORD_NYBL                  'Nybble data type (4-bits)
    TOKEN_KEYWORD_PARAMS                'Parameter list `PROC :<label> PARAMS <list>`
    TOKEN_KEYWORD_PROC                  'Procedure Chunk `PROC :<label> { ... }`
    TOKEN_KEYWORD_RAM                   'RAM definition
    TOKEN_KEYWORD_RETURN                'Returns list `PROC :<label> RETURN <list>`
    TOKEN_KEYWORD_SECTION               'Section definition `SECTION ::<section>
    TOKEN_KEYWORD_SLOT                  'Section Slot pattern `SLOT 0, 1, 2`
    TOKEN_KEYWORD_START                 'The starting vector for the System
    TOKEN_KEYWORD_SYSTEM                'System identifier `SYSTEM "SMS"`
    TOKEN_KEYWORD_TABLE                 'Data table
    TOKEN_KEYWORD_TRIP                  'Triple data type (3-bytes)
    TOKEN_KEYWORD_WORD                  'Word data type (2-bytes)
    [_TOKEN_KEYWORDS_END]
    
    TOKEN_NUMBER
    'Number prefixes ("K" & "KB")
    TOKEN_PREFIX_K                      'x1000
    TOKEN_PREFIX_KB                     'x1024
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_BRACES_OPEN                   '"{" Code/data Chunk, `PROC :<label> { ... }`
    TOKEN_BRACES_CLOSE                  '"}"
    TOKEN_SQUARE_OPEN                   '"[" Hash array and
    TOKEN_SQUARE_CLOSE                  '"]" Memory reference `ld a, [hl]`
    TOKEN_PARENS_OPEN                   '"(" Expression nesting,
    TOKEN_PARENS_CLOSE                  '")" e.g. `HI ($8000 + $80)`
    
    TOKEN_CONST                         'e.g. `!CONST`
    TOKEN_HASH                          'e.g. `#hash`
    TOKEN_LABEL                         'e.g. `:label`
    TOKEN_PROPERTY_USE
    TOKEN_PROPERTY_NEW
    TOKEN_RAM                           'e.g. `$.ram`
    TOKEN_SECTION                       'e.g. `::section`
    TOKEN_TEXT                          'e.g. `"..."`
    
    [_TOKEN_LAST]                       'Do not go above 255!
End Enum

'For logging, we will want to get a text representation of any of the Tokens _
 (oh how I wish VB6 supported static constant arrays)
Private TokenNames(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String

'--------------------------------------------------------------------------------------

'A token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source text files again. Below is the _
 structure for a single token, representing one "word" in the source code

Private Type oz80Token
    Kind As OZ80_TOKEN
    Value As Double                     'Associated value, e.g. for number tokens
    Line As Long                        'Line number in the original source text
    Col As Long                         'Column number in the original source text
End Type

'Use a 20'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 20000

'Caches of the current token to make multiple calls to this class a little quicker, _
 it also helps with a lot of internal referencing, mainly with logging
Private TokenKind As OZ80_TOKEN
Private TokenValue As Double
'When parsing source text, we can save having to pass these through function calls _
 hundreds of times by making them global to this class
Private TokenLine As Long
Private TokenCol As Long

'Look up tables to determine what category a Token belongs to
Private TokenIsExpression(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsFlag(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsList(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsInstruction(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsKeyword(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsOperator(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsParameter(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsRegister(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsRegister8Bit(0 To [_TOKEN_LAST]) As Boolean
Private TokenIsRegister16Bit(0 To [_TOKEN_LAST]) As Boolean

'We need to be able to tell the difference between a property being used as a _
 parameter on a Z80 instruction, and a label-property definition (within a chunk) _
 which occurs on a new line. We keep track of the previous line number used and _
 look for a property occuring on a new line
Private PrevLine As Long

'We will capture the word as we read it here. The bluString is both fast and _
 allows us to access the string as bytes or text
Private Word As bluString

'We track the opening and closing of brackets to validate nesting of brackets _
 and to allow the assembler to skip over blocks
Private Nesting(0 To 255) As TokenNesting
Private NestingDepth As Long

Private Type TokenNesting
    Token As OZ80_TOKEN                 'Which type of bracket `{`, `[`, `(`
    Index As Long                       'The token index of the opening bracket
End Type

'--------------------------------------------------------------------------------------

Private Type oz80File
    'Remember where the original source file is as the assembler will need to change
     'to its directory when handling relative references with includes
    Path As String
    
    Tokens() As oz80Token               'The token stream will live here
    TokensBuffer As Long                'The length of the buffer used for the above
    TokensLength As Long                'The current amount of space used in the array
End Type

'Each source code file is broken down into a machine-readable representation _
 ('oz80File') which we store here. When assembling one file and we come _
 across an `INCLUDE` statement, another File gets added here
Private Files() As oz80File
Private FileCount As Long
'The current source file being processed
Private File As Long

'The current position in the token stream (of the current file). _
 Since VB6 is not multi-threaded we don't have to store this per-file
Private TokenIndex As Long

'The portions of Text in the source code file
Private TokenStrings As bluArrayStrings

'--------------------------------------------------------------------------------------

'Which target system is selected, e.g. SEGA Master System
Private My_System As OZ80_SYSTEM

'A list of system targets. Only the SEGA Master System is supported at the moment, _
 but I will consider supporting other Z80 systems in the future
Public Enum OZ80_SYSTEM
    SYSTEM_NONE                         'System not yet defined
    SYSTEM_SMS                          'SEGA Master System
End Enum

'--------------------------------------------------------------------------------------

'The database of named item values. _
 This includes Constants, RAM Names and Hash values
Private Values As Scripting.Dictionary

Private Type oz80Section
    SectionName As String
    Slot As OZ80_SLOT
End Type

'Sections need to be defined first before they can be used
Private SectionSlots As Scripting.Dictionary
'For normalising Label names, we note which Section is in play
Private CurrentSection As String

'The database of Label locations (in ROM) and their properties. _
 These are separate from the other named items since we keep the _
 Label addresses for the second assembly pass
Private Labels As Scripting.Dictionary
'For normalising Property names, we note which Label is in play
Private CurrentLabel As String

'Z80 Parameters: _
 --------------------------------------------------------------------------------------
'In order to compare the hundreds of permutations of parameters for Z80 instructions, _
 we assign each parameter type a single bit. We can thus check very quickly if a given _
 parameter falls within an allowed list of accepted types

Private Enum OZ80_MASK
    
    [_MASK_REGS_BEGIN] = 1
    MASK_REG_B = 2 ^ 0
    MASK_REG_C = 2 ^ 1
    MASK_REG_D = 2 ^ 2
    MASK_REG_E = 2 ^ 3
    MASK_REG_H = 2 ^ 4
    MASK_REG_L = 2 ^ 5
    MASK_MEM_HL = 2 ^ 6
    MASK_REG_A = 2 ^ 7
    
    'The presence of an IX/IY prefix on the opcode changes H/L to IXH/IYH or IXL/IYL
     'respectively, but only on instructions that use single byte opcodes.
     'This is officially undocumented, but obviously fair game for old systems
    MASK_REG_IXH = 2 ^ 8
    MASK_REG_IXL = 2 ^ 9
    MASK_REG_IYH = 2 ^ 10
    MASK_REG_IYL = 2 ^ 11
    
    'The presence of an IX/IY prefix on the opcode changes a memory reference "[HL]"
     'to IX/IY, with an offset value e.g. "[IX+$8]"
    MASK_MEM_IX = 2 ^ 12
    MASK_MEM_IY = 2 ^ 13
    [_MASK_REGS_END] = MASK_MEM_IY
    
    'A couple of undocumented instructions allow for IX/IY memory references,
     'but not the standard "[HL]" reference
    MASK_MEM_IXY = MASK_MEM_IX Or MASK_MEM_IY
    'And this is the common "[HL|IX+$8|IY+$8]" form that is used often throughout
    MASK_MEM_HLIXY = MASK_MEM_HL Or MASK_MEM_IXY
    
    'The main 8-bit registers are a common instruction parameter
    MASK_REGS_ABCDEHL = MASK_REG_A Or MASK_REG_B Or MASK_REG_C Or MASK_REG_D Or MASK_REG_E Or MASK_REG_E Or MASK_REG_H Or MASK_REG_L
    'The Z80 clumps HL/IX & IY memory references together with 8-bit registers when
     'building opcodes, i.e. "A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]"
    MASK_REGS_ABCDEHL_MEM_HLIXY = MASK_REGS_ABCDEHL Or MASK_MEM_HLIXY
    'The use of the IX/IY prefix turns H/L into IXH/IXL/IYH/IYL in many instances
    MASK_REGS_IXHL = MASK_REG_IXH Or MASK_REG_IXL
    MASK_REGS_IYHL = MASK_REG_IYH Or MASK_REG_IYL
    MASK_REGS_IXYHL = MASK_REGS_IXHL Or MASK_REGS_IYHL
    MASK_REGS_ABCDEIXYHL_MEM_HLIXY = MASK_REGS_ABCDEHL_MEM_HLIXY Or MASK_REGS_IXYHL
    
    'Very uncommon 8-bit registers
    MASK_REG_I = 2 ^ 14                 'Interrupt register
    MASK_REG_R = 2 ^ 15                 'Refresh register, pseudo-random
    
    'The 16-bit register pairs
    MASK_REG_AF = 2 ^ 16                'The Accumulator and the processor Flags
    MASK_REG_BC = 2 ^ 17                'Registers B & C
    MASK_REG_DE = 2 ^ 18                'Registers D & E
    MASK_REG_HL = 2 ^ 19                'Registers H & L
    MASK_REG_SP = 2 ^ 20                'Stack Pointer
    
    MASK_REG_IX = 2 ^ 21
    MASK_REG_IY = 2 ^ 22
    
    MASK_REGS_BC_DE_SP = MASK_REG_BC Or MASK_REG_DE Or MASK_REG_SP
    'Some instructions accept BC/DE/HL/SP, but not IX & IY due to existing prefixes
    MASK_REGS_BC_DE_HL_SP = MASK_REGS_BC_DE_SP Or MASK_REG_HL
    
    'HL, IX & IY are synonymous as they use an opcode prefix to determine which
    MASK_REGS_HL_IXY = MASK_REG_HL Or MASK_REG_IX Or MASK_REG_IY
    'PUSH / POP allow AF but not SP
    MASK_REGS_AF_BC_DE_HL_IXY = MASK_REG_AF Or MASK_REG_BC Or MASK_REG_DE Or MASK_REGS_HL_IXY
    'The LD instruction can take most 16-bit registers
    MASK_REGS_BC_DE_HL_SP_IXY = MASK_REGS_BC_DE_HL_SP Or MASK_REG_IX Or MASK_REG_IY
    
    'Any immediate 8/16-bit Value,
     'e.g. `LD A, $FF`
    MASK_VAL = 2 ^ 23
    
    '..................................................................................
    
    'Register C & Flag C cannot be distinguished by the tokeniser (it isn't aware of
     'context) so they are treated as the same thing. Another bit covers NC/Z/NZ so
     'that these are not accidentally taken as Register C elsewhere
    MASK_FLAGS_CZ = MASK_REG_C Or (2 ^ 24)
    MASK_FLAGS_MP = (2 ^ 25)
    
    MASK_FLAGS = MASK_FLAGS_CZ Or MASK_FLAGS_MP
    
    '..................................................................................
    
    'The IN and OUT instructions can use port "C" (which is, in reality, BC)
    MASK_MEM_BC = 2 ^ 26
    MASK_MEM_DE = 2 ^ 27
    MASK_MEM_SP = 2 ^ 28
    
    MASK_MEM_VAL = 2 ^ 29
    
    '..................................................................................
    
    'This is a shorthand to check for any instance of IX/IY so that we can add the
     'relevant opcode prefix with the simplest of tests
    MASK_ANY_IX = MASK_REG_IX Or MASK_REG_IXH Or MASK_REG_IXL Or MASK_MEM_IX
    MASK_ANY_IY = MASK_REG_IY Or MASK_REG_IYH Or MASK_REG_IYL Or MASK_MEM_IY
    MASK_ANY_IXY = MASK_ANY_IX Or MASK_ANY_IY
    
    'This is shorthand to check for any 8/16-bit parameter Value
    MASK_ANY_VAL = MASK_VAL Or MASK_MEM_VAL
End Enum

Private Type oz80Param
    Mask As OZ80_MASK
    Token As OZ80_TOKEN
    Value As Long
    LogTextPtr As Long                  'When logging, which cached text to use
End Type

'--------------------------------------------------------------------------------------

'For maximum logging speed, we're going to build a massive (~2.5 MB) cache of the _
 possible Z80 parameters as text

'First we create a big list of indexes that contain all the possible Z80 parameters, _
 this includes all 256 possible 8-bit values and all 65'536 possible 16-bit values _
 as well as the "[IX/IY+$??]" values and direct memory references "[$????]"

Private Enum PARAMTEXT_OFFSETS
    PARAMTEXT_NONE
    
    'The Z80 instruction mneomnics
    PARAMTEXT_Z80
    [_PARAMTEXT_Z80] = ([_TOKEN_Z80_END] - [_TOKEN_Z80_BEGIN])
    
    'We have 28 distinct names for registers/flags,
     '"A|AF|B|BC|C|D|DE|E|H|HL|I|IX|IXH|IXL|IY|IYH|IYL|L|M|NC|NZ|P|PC|PE|PO|R|SP|Z"
    PARAMTEXT_REGS = PARAMTEXT_Z80 + [_PARAMTEXT_Z80] + 1
    [_PARAMTEXT_REGS] = ([_TOKEN_REGS_END] - [_TOKEN_REGS_BEGIN])
    
    'The 8-bit values "$00" to "$FF"
    PARAMTEXT_VAL8 = PARAMTEXT_REGS + [_PARAMTEXT_REGS] + 1
    [_PARAMTEXT_VAL8] = &HFF&
    
    'The 16-bit values "$0000" to "$FFFF"
    PARAMTEXT_VAL16 = PARAMTEXT_VAL8 + [_PARAMTEXT_VAL8] + 1
    [_PARAMTEXT_VAL16] = &HFFFF&
    
    'The memory references, registers
    PARAMTEXT_MEM_BC = PARAMTEXT_VAL16 + [_PARAMTEXT_VAL16] + 1
    PARAMTEXT_MEM_DE
    PARAMTEXT_MEM_HL
    PARAMTEXT_MEM_SP
    PARAMTEXT_MEM_IX                    '"[IX]" - without 8-bit value
    PARAMTEXT_MEM_IY                    '"[IY]" - without 8-bit value
    
    'The IX memory references, with 8-bit value
    PARAMTEXT_MEM_IX_VAL
    [_PARAMTEXT_MEM_IX_VAL] = &HFF&
    'And the IY counterpart
    PARAMTEXT_MEM_IY_VAL = PARAMTEXT_MEM_IX_VAL + [_PARAMTEXT_MEM_IX_VAL] + 1
    [_PARAMTEXT_MEM_IY_VAL] = &HFF&
    
    'The 16-bit memory references, i.e. "[$FFFF]"
    PARAMTEXT_MEM_VAL16 = PARAMTEXT_MEM_IY_VAL + [_PARAMTEXT_MEM_IY_VAL] + 1
    [_PARAMTEXT_MEM_VAL16] = &HFFFF&
    
    [_PARAMTEXT_TOTAL] = PARAMTEXT_MEM_VAL16 + [_PARAMTEXT_MEM_VAL16]
End Enum

Private ParamText(0 To [_PARAMTEXT_TOTAL]) As String * 10

'Z80 Opcodes: _
 --------------------------------------------------------------------------------------
'Z80 opcodes are made up of various bit patterns that can be broken down into the _
 following layout: (this information from <z80.info/decoding.htm>)

'             .---------------.
' Bit Number: |7|6|5|4|3|2|1|0|
'             '---------------'
'  Component: | X |  Y  |  Z  |
'             '---------------'
'                 | P |Q|
'                 '-----'

'Here we encode these component's bit numbers so we can easily build opcodes. _
 A set of look-up tables translate instructions / registers into component parts

Private Enum OZ80_OPCODE
    X0 = 0:         X1 = 2 ^ 6:     X2 = 2 ^ 7:     X3 = X1 Or X2
    
    Y0 = 0:         Y1 = 2 ^ 3:     Y2 = 2 ^ 4:     Y3 = Y1 Or Y2
    Y4 = 2 ^ 5:     Y5 = Y4 Or Y1:  Y6 = Y4 Or Y2:  Y7 = Y4 Or Y3
    
    P0 = 0:         P1 = 2 ^ 4:     P2 = 2 ^ 5:     P3 = P1 Or P2

    Q0 = 0:         Q1 = Y1
    
    Z0 = 0:         Z1 = 2 ^ 0:     Z2 = 2 ^ 1:     Z3 = Z1 Or Z2
    Z4 = 2 ^ 2:     Z5 = Z4 Or Z1:  Z6 = Z4 Or Z2:  Z7 = Z4 Or Z3
End Enum

'A look-up table to convert a token for a register into the bit mask for it
Private TokenRegister([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long

'Look-up tables for converting certain z80 parameters into opcode-components
Private OpcodeP_RegisterPair([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long
Private OpcodeY_Condition([_TOKEN_REGS_BEGIN] To [_TOKEN_REGS_END]) As Long
Private OpcodeY_RST(0 To &H38) As Long
Private OpcodeY_Bit(0 To 7) As Long
'Sure this is a waste of memory, but it is fast as we don't have to convert _
 bit-powers to linear numbers for many instructions being assembled
Private OpcodeY_Register([_MASK_REGS_BEGIN] To [_MASK_REGS_END]) As Long
Private OpcodeZ_Register([_MASK_REGS_BEGIN] To [_MASK_REGS_END]) As Long

'Our routine for comparing parameters is going to be used hundreds of times, _
 so for clarity, we'll use some class-level references to share the input _
 parameters with the comparison routine
Private Z80Param1 As oz80Param
Private Z80Param2 As oz80Param
Private Z80Param3 As oz80Param
'For quick emptying of the params
Private Z80ParamEmpty As oz80Param

'Whilst in the syntax `SLOT` uses a list (i.e. `SLOT 0, 1, 2`), we convert that into _
 a bit pattern to make it quick and easy to work with instead of iterating an array
Private Enum OZ80_SLOT
    SLOT0 = 2 ^ 0
    SLOT1 = 2 ^ 1
    SLOT2 = 2 ^ 2
End Enum

'--------------------------------------------------------------------------------------
'ROM Layout:
'--------------------------------------------------------------------------------------

Private Type oz80Chunk
    LabelName As String
    Data() As Byte
    Buffer As Long                      'Buffer size for the data (power of 2)
    Size As Long                        'Length of the actual data in bytes
    Section As Long
    Address As Long
    LogASM As bluLog
End Type
'We'll erase a Chunk by setting it to this empty instance
Private EmptyChunk As oz80Chunk


Private Chunks() As oz80Chunk
Private ChunkCount As Long

'The current chunk being assembled into. This is a reference into the Chunks array _
 so that we can speak to the current Chunk in scope consistently
Private Chunk As Long

'These will be references to the special Chunks at fixed positions within the ROM _
 (when laying out the ROM, we'll need to position these first)
Private Chunk_Header As Long
Private Chunk_Interrupt00 As Long
Private Chunk_Interrupt08 As Long
Private Chunk_Interrupt10 As Long
Private Chunk_Interrupt18 As Long
Private Chunk_Interrupt20 As Long
Private Chunk_Interrupt28 As Long
Private Chunk_Interrupt30 As Long
Private Chunk_Interrupt38 As Long
Private Chunk_Interrupt66 As Long
Private Chunk_Start As Long

Private Type oz80Layout
    Chunk As Long
    
    AddrBegin As Long
    AddrSize As Long
    AddrEnd As Long
    
    PrevIndex As Long
    NextIndex As Long
End Type

Private Layout() As oz80Layout
Private LayoutCount As Long

Private ROM() As Byte
Private ROMSize As Long

'--------------------------------------------------------------------------------------
'Cache:
'--------------------------------------------------------------------------------------

'Lookup table of hexadecimal prettyprint, _
 saves slow & repetitive text manipulation when logging
Private Cache_HexStr8(-1 To &HFF&) As String * 2
Private Cache_HexStr16(0 To &HFFFF&) As String * 4

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output. _
 A bluString is provided to make manipulating the message before displaying very easy
Event Message( _
    ByRef LogLevel As OZ80_LOG, _
    ByRef LogText As bluString _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByRef FilePath As String, _
    ByVal Number As OZ80_ERROR, _
    ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Dim i As Long
    
    'If too many possible tokens are created, things will screw up
    Debug.Assert OZ80_TOKEN.[_TOKEN_LAST] < 256
    Debug.Print "[OZ80] No. Tokens Defined: " & (OZ80_TOKEN.[_TOKEN_LAST] - 1)
    Debug.Print "[OZ80] Size of param text: " & [_PARAMTEXT_TOTAL] * 20 & " bytes"
    
    'Register the events we'll be timing with the profiler
    With Profiler
        Call .RegisterEvent(EVENT_INITIALIZE, "Initialize")
        Call .RegisterEvent(EVENT_TOKENISE, "File_Tokenise")
        Call .RegisterEvent(EVENT_TOKENISE_READWORD, "File_Tokenise;ReadWord")
        Call .RegisterEvent(EVENT_TOKENISE_TOKENWORD, "File_Tokenise;TokenWord")
        Call .RegisterEvent(EVENT_FORMATTOKEN, "FormatToken")
        Call .RegisterEvent(EVENT_PROCESSZ80, "Z80_Process")
    End With
    
    '*** TIMER BEGIN ***
    'Begin using our profiler to time all this caching and initialising
    Call Profiler.EventBegin(EVENT_INITIALIZE)
    
    'This is a re-usable temporary String for logging
    Set LogMsg = New bluString
    
    'Manually create a 64-bit Double "Infinite" number, which cannot be written _
     implicitly in VB6. This method is based upon <stackoverflow.com/a/896292>. _
     We use this to mark numbers as "not yet defined" whilst still allowing full _
     64-Bit floating-point calculations
    Call api_PutMem1(VarPtr(INDEFINITE) + 6, 240)
    Call api_PutMem1(VarPtr(INDEFINITE) + 7, 127)
    'Manually create the largest possible 64-bit number
    Call api_PutMem2(VarPtr(NUMBER_MAX), &HFFFFFFFF)
    Call api_PutMem2(VarPtr(NUMBER_MAX) + 4, -1)
    Call api_PutMem1(VarPtr(NUMBER_MAX) + 6, 239)
    Call api_PutMem1(VarPtr(NUMBER_MAX) + 7, 127)
    'Manually create the largest negative 64-bit number
    Call api_PutMem2(VarPtr(NUMBER_MIN), &HFFFFFFFF)
    Call api_PutMem2(VarPtr(NUMBER_MIN) + 4, -1)
    Call api_PutMem1(VarPtr(NUMBER_MIN) + 6, 239)
    Call api_PutMem1(VarPtr(NUMBER_MIN) + 7, 255)
    'Define min/max numbers that can be multiplied by 1000/1024 without overflow
    Let NUMBER_MAX_K = NUMBER_MAX / 1000
    Let NUMBER_MIN_K = -NUMBER_MAX_K
    Let NUMBER_MAX_KB = NUMBER_MAX / 1024
    Let NUMBER_MIN_KB = -NUMBER_MAX_KB
    
    '----------------------------------------------------------------------------------
    'Upper-case look-up table:
    '----------------------------------------------------------------------------------
    'Begin with a straight forward list of all 2-byte UCS-2 code points
    For i = 0 To &HFFFF&
        If i >= 32768 _
            Then Let ToUpperCase(i) = i - 65536 _
            Else Let ToUpperCase(i) = i
    Next i
    
    'Get the Windows version as Windows Vista and above use alternative APIs for _
     locale-dependent methods (such as `Equals`)
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version. Windows 8.1 and above also cause _
     this API to always return as Windows 8.0 unless the app opts-in to correct _
     version info via manifest, but this isn't an issue to bluString as we are _
     only concerned with testing for Windows Vista+
    If api_GetVersionEx(VersionInfo) = API_TRUE Then
        '`Val`, rather than `CSng` is used here due to locale differences; _
         `CSng("6.1")` will fail in locales that use comma as decimal separator
        Let WindowsVersion = Val( _
            VersionInfo.MajorVersion & "." & VersionInfo.MinorVersion _
        )
    End If
    
    'Use the Windows API to convert these to their linguistic upper-case equivilents
    'Check for Windows Vista and above:
    If WindowsVersion >= 6 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call api_LCMapStringEx( _
                   LocaleNamePointer:=0&, _
                            MapFlags:=LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
                 SourceStringPointer:=VarPtr(ToUpperCase(0)), _
                  SourceStringLength:=UBound(ToUpperCase) + 1, _
            DestinationStringPointer:=VarPtr(ToUpperCase(0)), _
             DestinationStringLength:=UBound(ToUpperCase) + 1, _
                  VersionInfoPointer:=0&, ReservedPointer:=0&, SortHandle:=0& _
        )
    Else
        'Use the Windows XP compatible API
        Call api_LCMapString( _
                            LocaleID:=api_GetThreadLocale(), _
                            MapFlags:=LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
                 SourceStringPointer:=VarPtr(ToUpperCase(0)), _
                  SourceStringLength:=UBound(ToUpperCase) + 1, _
            DestinationStringPointer:=VarPtr(ToUpperCase(0)), _
             DestinationStringLength:=UBound(ToUpperCase) + 1 _
        )
    End If
    
    '----------------------------------------------------------------------------------
    
    For i = 0 To &HF&:          Let Cache_HexStr8(i) = "0" & Hex$(i):       Next i
    For i = &H10& To &HFF&:     Let Cache_HexStr8(i) = Hex$(i):             Next i
    Let Cache_HexStr8(-1) = "··"
    
    For i = 0 To &HF&:          Let Cache_HexStr16(i) = "000" & Hex$(i):    Next i
    For i = &H10& To &HFF&:     Let Cache_HexStr16(i) = "00" & Hex$(i):     Next i
    For i = &H100& To &HFFF&:   Let Cache_HexStr16(i) = "0" & Hex$(i):      Next i
    For i = &H1000& To &HFFFF&: Let Cache_HexStr16(i) = Hex$(i):            Next i
    
    'Z80 Instructions .................................................................
    Let TokenNames(TOKEN_Z80_ADC) = STR_ADC
    Let TokenNames(TOKEN_Z80_ADD) = STR_ADD
    Let TokenNames(TOKEN_Z80_AND) = STR_AND
    Let TokenNames(TOKEN_Z80_BIT) = STR_BIT
    Let TokenNames(TOKEN_Z80_CALL) = STR_CALL
    Let TokenNames(TOKEN_Z80_CCF) = STR_CCF
    Let TokenNames(TOKEN_Z80_CP) = STR_CP
    Let TokenNames(TOKEN_Z80_CPD) = STR_CPD
    Let TokenNames(TOKEN_Z80_CPDR) = STR_CPDR
    Let TokenNames(TOKEN_Z80_CPI) = STR_CPI
    Let TokenNames(TOKEN_Z80_CPIR) = STR_CPIR
    Let TokenNames(TOKEN_Z80_CPL) = STR_CPL
    Let TokenNames(TOKEN_Z80_DAA) = STR_DAA
    Let TokenNames(TOKEN_Z80_DEC) = STR_DEC
    Let TokenNames(TOKEN_Z80_DI) = STR_DI
    Let TokenNames(TOKEN_Z80_DJNZ) = STR_DJNZ
    Let TokenNames(TOKEN_Z80_EI) = STR_EI
    Let TokenNames(TOKEN_Z80_EX) = STR_EX
    Let TokenNames(TOKEN_Z80_EXX) = STR_EXX
    Let TokenNames(TOKEN_Z80_HALT) = STR_HALT
    Let TokenNames(TOKEN_Z80_IM) = STR_IM
    Let TokenNames(TOKEN_Z80_IN) = STR_IN
    Let TokenNames(TOKEN_Z80_INC) = STR_INC
    Let TokenNames(TOKEN_Z80_IND) = STR_IND
    Let TokenNames(TOKEN_Z80_INDR) = STR_INDR
    Let TokenNames(TOKEN_Z80_INI) = STR_INI
    Let TokenNames(TOKEN_Z80_INIR) = STR_INIR
    Let TokenNames(TOKEN_Z80_JP) = STR_JP
    Let TokenNames(TOKEN_Z80_JR) = STR_JR
    Let TokenNames(TOKEN_Z80_LD) = STR_LD
    Let TokenNames(TOKEN_Z80_LDD) = STR_LDD
    Let TokenNames(TOKEN_Z80_LDDR) = STR_LDDR
    Let TokenNames(TOKEN_Z80_LDI) = STR_LDI
    Let TokenNames(TOKEN_Z80_LDIR) = STR_LDIR
    Let TokenNames(TOKEN_Z80_NEG) = STR_NEG
    Let TokenNames(TOKEN_Z80_NOP) = STR_NOP
    Let TokenNames(TOKEN_Z80_OR) = STR_OR
    Let TokenNames(TOKEN_Z80_OUT) = STR_OUT
    Let TokenNames(TOKEN_Z80_OUTD) = STR_OUTD
    Let TokenNames(TOKEN_Z80_OTDR) = STR_OTDR
    Let TokenNames(TOKEN_Z80_OUTI) = STR_OUTI
    Let TokenNames(TOKEN_Z80_OTIR) = STR_OTIR
    Let TokenNames(TOKEN_Z80_POP) = STR_POP
    Let TokenNames(TOKEN_Z80_PUSH) = STR_PUSH
    Let TokenNames(TOKEN_Z80_RES) = STR_RES
    Let TokenNames(TOKEN_Z80_RET) = STR_RET
    Let TokenNames(TOKEN_Z80_RETI) = STR_RETI
    Let TokenNames(TOKEN_Z80_RETN) = STR_RETN
    Let TokenNames(TOKEN_Z80_RLA) = STR_RLA
    Let TokenNames(TOKEN_Z80_RL) = STR_RL
    Let TokenNames(TOKEN_Z80_RLC) = STR_RLC
    Let TokenNames(TOKEN_Z80_RLCA) = STR_RLCA
    Let TokenNames(TOKEN_Z80_RLD) = STR_RLD
    Let TokenNames(TOKEN_Z80_RR) = STR_RR
    Let TokenNames(TOKEN_Z80_RRA) = STR_RRA
    Let TokenNames(TOKEN_Z80_RRC) = STR_RRC
    Let TokenNames(TOKEN_Z80_RRCA) = STR_RRCA
    Let TokenNames(TOKEN_Z80_RRD) = STR_RRD
    Let TokenNames(TOKEN_Z80_RST) = STR_RST
    Let TokenNames(TOKEN_Z80_SBC) = STR_SBC
    Let TokenNames(TOKEN_Z80_SCF) = STR_SCF
    Let TokenNames(TOKEN_Z80_SET) = STR_SET
    Let TokenNames(TOKEN_Z80_SLA) = STR_SLA
    Let TokenNames(TOKEN_Z80_SLL) = STR_SLL
    Let TokenNames(TOKEN_Z80_SRA) = STR_SRA
    Let TokenNames(TOKEN_Z80_SRL) = STR_SRL
    Let TokenNames(TOKEN_Z80_SUB) = STR_SUB
    Let TokenNames(TOKEN_Z80_XOR) = STR_XOR
    
    'Z80 Registers / Flags ............................................................
    Let TokenNames(TOKEN_Z80_A) = STR_A
    Let TokenNames(TOKEN_Z80_AF) = STR_AF
    Let TokenNames(TOKEN_Z80_B) = STR_B
    Let TokenNames(TOKEN_Z80_C) = STR_C
    Let TokenNames(TOKEN_Z80_NC) = STR_NC
    Let TokenNames(TOKEN_Z80_BC) = STR_BC
    Let TokenNames(TOKEN_Z80_D) = STR_D
    Let TokenNames(TOKEN_Z80_E) = STR_E
    Let TokenNames(TOKEN_Z80_DE) = STR_DE
    Let TokenNames(TOKEN_Z80_H) = STR_H
    Let TokenNames(TOKEN_Z80_L) = STR_L
    Let TokenNames(TOKEN_Z80_HL) = STR_HL
    Let TokenNames(TOKEN_Z80_I) = STR_I
    Let TokenNames(TOKEN_Z80_IX) = STR_IX
    Let TokenNames(TOKEN_Z80_IXL) = STR_IXL
    Let TokenNames(TOKEN_Z80_IXH) = STR_IXH
    Let TokenNames(TOKEN_Z80_IY) = STR_IY
    Let TokenNames(TOKEN_Z80_IYL) = STR_IYL
    Let TokenNames(TOKEN_Z80_IYH) = STR_IYH
    Let TokenNames(TOKEN_Z80_M) = STR_M
    Let TokenNames(TOKEN_Z80_P) = STR_P
    Let TokenNames(TOKEN_Z80_PC) = STR_PC
    Let TokenNames(TOKEN_Z80_PE) = STR_PE
    Let TokenNames(TOKEN_Z80_PO) = STR_PO
    Let TokenNames(TOKEN_Z80_R) = STR_R
    Let TokenNames(TOKEN_Z80_SP) = STR_SP
    Let TokenNames(TOKEN_Z80_Z) = STR_Z
    Let TokenNames(TOKEN_Z80_NZ) = STR_NZ
    
    'Operators ........................................................................
    Let TokenNames(TOKEN_OPERATOR_ADD) = Chr$(SYNTAX_OPERATOR_ADD)
    Let TokenNames(TOKEN_OPERATOR_SUB) = Chr$(SYNTAX_OPERATOR_SUB)
    Let TokenNames(TOKEN_OPERATOR_MUL) = Chr$(SYNTAX_OPERATOR_MUL)
    Let TokenNames(TOKEN_OPERATOR_DIV) = Chr$(SYNTAX_OPERATOR_DIV)
    Let TokenNames(TOKEN_OPERATOR_POW) = Chr$(SYNTAX_OPERATOR_POW)
    Let TokenNames(TOKEN_OPERATOR_MOD) = Chr$(SYNTAX_OPERATOR_MOD)
    Let TokenNames(TOKEN_OPERATOR_REP) = "x"
    Let TokenNames(TOKEN_OPERATOR_OR) = Chr$(SYNTAX_OPERATOR_OR)
    Let TokenNames(TOKEN_OPERATOR_AND) = Chr$(SYNTAX_OPERATOR_AND)
    Let TokenNames(TOKEN_OPERATOR_XOR) = Chr$(SYNTAX_OPERATOR_XOR)
    
    'Keywords .........................................................................
    Let TokenNames(TOKEN_KEYWORD_BOOL) = STR_BOOL
    Let TokenNames(TOKEN_KEYWORD_BYTE) = STR_BYTE
    Let TokenNames(TOKEN_KEYWORD_DEF) = STR_DEF
    Let TokenNames(TOKEN_KEYWORD_HASH) = STR_HASH
    Let TokenNames(TOKEN_KEYWORD_HELP) = STR_HELP
    Let TokenNames(TOKEN_KEYWORD_HI) = STR_HI
    Let TokenNames(TOKEN_KEYWORD_INCLUDE) = STR_INCLUDE
    Let TokenNames(TOKEN_KEYWORD_INDEX) = STR_INDEX
    Let TokenNames(TOKEN_KEYWORD_INTERRUPT) = STR_INTERRUPT
    Let TokenNames(TOKEN_KEYWORD_LO) = STR_LO
    Let TokenNames(TOKEN_KEYWORD_LONG) = STR_LONG
    Let TokenNames(TOKEN_KEYWORD_NYBL) = STR_NYBL
    Let TokenNames(TOKEN_KEYWORD_PARAMS) = STR_PARAMS
    Let TokenNames(TOKEN_KEYWORD_PROC) = STR_PROC
    Let TokenNames(TOKEN_KEYWORD_RAM) = STR_RAM
    Let TokenNames(TOKEN_KEYWORD_RETURN) = STR_RETURN
    Let TokenNames(TOKEN_KEYWORD_SECTION) = STR_SECTION
    Let TokenNames(TOKEN_KEYWORD_SLOT) = STR_SLOT
    Let TokenNames(TOKEN_KEYWORD_START) = STR_START
    Let TokenNames(TOKEN_KEYWORD_SYSTEM) = STR_SYSTEM
    Let TokenNames(TOKEN_KEYWORD_TABLE) = STR_TABLE
    Let TokenNames(TOKEN_KEYWORD_TRIP) = STR_TRIP
    Let TokenNames(TOKEN_KEYWORD_WORD) = STR_WORD
    
    Let TokenNames(TOKEN_PREFIX_K) = STR_K
    Let TokenNames(TOKEN_PREFIX_KB) = STR_KB
    
    Let TokenNames(TOKEN_BRACES_OPEN) = Chr$(SYNTAX_BRACES_OPEN)
    Let TokenNames(TOKEN_BRACES_CLOSE) = Chr$(SYNTAX_BRACES_CLOSE)
    Let TokenNames(TOKEN_PARENS_OPEN) = Chr$(SYNTAX_PARENS_OPEN)
    Let TokenNames(TOKEN_PARENS_CLOSE) = Chr$(SYNTAX_PARENS_CLOSE)
    Let TokenNames(TOKEN_SQUARE_OPEN) = Chr$(SYNTAX_SQUARE_OPEN)
    Let TokenNames(TOKEN_SQUARE_CLOSE) = Chr$(SYNTAX_SQUARE_CLOSE)
    
    Let TokenNames(TOKEN_TEXT) = Chr$(SYNTAX_TEXT)
    Let TokenNames(TOKEN_LABEL) = Chr$(SYNTAX_LABEL)
    Let TokenNames(TOKEN_PROPERTY_USE) = Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_PROPERTY_NEW) = Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_RAM) = Chr$(SYNTAX_NUMBER_HEX) & Chr$(SYNTAX_PROPERTY)
    Let TokenNames(TOKEN_SECTION) = String(2, Chr$(SYNTAX_LABEL))
    
    '----------------------------------------------------------------------------------
    'Z80 Param Text:
    '----------------------------------------------------------------------------------
    'Index 0 acts as an "unimplemented" marker so we can see when certain cache _
     indexes haven't been implemented yet
    Let ParamText(PARAMTEXT_NONE) = "?????????"
    
    For i = 0 To ([_TOKEN_Z80_END] - [_TOKEN_Z80_BEGIN])
        Let ParamText(PARAMTEXT_Z80 + i) = TokenNames([_TOKEN_Z80_BEGIN] + i)
    Next i
    
    For i = 0 To ([_TOKEN_REGS_END] - [_TOKEN_REGS_BEGIN])
        Let ParamText(PARAMTEXT_REGS + i) = TokenNames([_TOKEN_REGS_BEGIN] + i)
    Next i
    
    For i = 0 To [_PARAMTEXT_VAL8]
        Let ParamText(PARAMTEXT_VAL8 + i) = _
            ChrW$(SYNTAX_NUMBER_HEX) & Cache_HexStr8(i)
    Next i
    
    For i = 0 To [_PARAMTEXT_VAL16]
        Let ParamText(PARAMTEXT_VAL16 + i) = _
            ChrW$(SYNTAX_NUMBER_HEX) & Cache_HexStr16(i)
    Next i
    
    'Register memory references, sans value, _
     .e.g. "[BC]", "[DE]", "[SP]"
    Let ParamText(PARAMTEXT_MEM_BC) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_BC & ChrW$(SYNTAX_SQUARE_CLOSE)
    Let ParamText(PARAMTEXT_MEM_DE) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_DE & ChrW$(SYNTAX_SQUARE_CLOSE)
    Let ParamText(PARAMTEXT_MEM_HL) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_HL & ChrW$(SYNTAX_SQUARE_CLOSE)
    Let ParamText(PARAMTEXT_MEM_IX) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_IX & ChrW$(SYNTAX_SQUARE_CLOSE)
    Let ParamText(PARAMTEXT_MEM_IY) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_IY & ChrW$(SYNTAX_SQUARE_CLOSE)
    Let ParamText(PARAMTEXT_MEM_SP) = _
        ChrW$(SYNTAX_SQUARE_OPEN) & STR_SP & ChrW$(SYNTAX_SQUARE_CLOSE)
    
    'Register IX memory reference, with value -- i.e. "[IX+$FF]"
    For i = 0 To [_PARAMTEXT_MEM_IX_VAL]
        Let ParamText(PARAMTEXT_MEM_IX_VAL + i) = _
            ChrW$(SYNTAX_SQUARE_OPEN) & STR_IX & ChrW$(SYNTAX_OPERATOR_ADD) & _
            Cache_HexStr8(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    'Register IY memory reference, with value -- i.e. "[IY+$FF]"
    For i = 0 To [_PARAMTEXT_MEM_IY_VAL]
        Let ParamText(PARAMTEXT_MEM_IY_VAL + i) = _
            ChrW$(SYNTAX_SQUARE_OPEN) & STR_IY & ChrW$(SYNTAX_OPERATOR_ADD) & _
            Cache_HexStr8(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    
    '16-bit memory references, i.e "[$FFFF]"
    For i = 0 To [_PARAMTEXT_MEM_VAL16]
        Let ParamText(PARAMTEXT_MEM_VAL16 + i) = _
            ChrW$(SYNTAX_SQUARE_OPEN) & ChrW$(SYNTAX_NUMBER_HEX) & _
            Cache_HexStr16(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    
    '----------------------------------------------------------------------------------
    
    'Initialise the token stream string database _
     (the text portions within source code files)
    Set TokenStrings = New bluArrayStrings
    Let TokenStrings.AllowDuplicates = True
    
    'Build the look-up tables for checking which categories each Token falls into
    Dim t As OZ80_TOKEN
    For t = 0 To [_TOKEN_LAST]
        'Check what token category it belongs to
        '..............................................................................
        Let TokenIsInstruction(t) = (t >= [_TOKEN_Z80_BEGIN]) And _
                                  (t <= [_TOKEN_Z80_END])
        Let TokenIsOperator(t) = (t > [_TOKEN_OPERATORS_BEGIN]) And _
                               (t < [_TOKEN_OPERATORS_END])
        Let TokenIsKeyword(t) = (t > [_TOKEN_KEYWORDS_BEGIN]) And _
                              (t < [_TOKEN_KEYWORDS_END])
        
        'Registers and flags are more complex as there's some overlap
        '..............................................................................
        If (t = TOKEN_Z80_C) Then
            'At the token level, we can't differentiate Register C and Flag C; _
             we need context for that and we can't gaurantee the TokenStream is _
             valid until we begin assembling, therefore we treat Register C and _
             Flag C as the same and the assembler will sort out the difference
            Let TokenIsRegister8Bit(t) = True
            Let TokenIsParameter(t) = True
            Let TokenIsRegister(t) = True
            Let TokenIsFlag(t) = True
        
        '8-Bit registers:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_A) Or (t = TOKEN_Z80_B) Or (t = TOKEN_Z80_C) Or _
               (t = TOKEN_Z80_D) Or (t = TOKEN_Z80_E) Or (t = TOKEN_Z80_H) Or _
               (t = TOKEN_Z80_L) Or (t = TOKEN_Z80_I) Or (t = TOKEN_Z80_R) Or _
               (t = TOKEN_Z80_IXL) Or (t = TOKEN_Z80_IXH) Or _
               (t = TOKEN_Z80_IYL) Or (t = TOKEN_Z80_IYH) _
        Then
            Let TokenIsRegister8Bit(t) = True
            Let TokenIsParameter(t) = True
            Let TokenIsRegister(t) = True
        
        '16-Bit registers:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_AF) Or (t = TOKEN_Z80_BC) Or (t = TOKEN_Z80_DE) Or _
               (t = TOKEN_Z80_HL) Or (t = TOKEN_Z80_IX) Or (t = TOKEN_Z80_IY) Or _
               (t = TOKEN_Z80_SP) Or (t = TOKEN_Z80_PC) _
        Then
            Let TokenIsRegister16Bit(t) = True
            Let TokenIsParameter(t) = True
            Let TokenIsRegister(t) = True
        
        'Condition flags:
        '..............................................................................
        ElseIf (t = TOKEN_Z80_NC) Or (t = TOKEN_Z80_M) Or (t = TOKEN_Z80_P) Or _
               (t = TOKEN_Z80_PE) Or (t = TOKEN_Z80_PO) Or (t = TOKEN_Z80_Z) Or _
               (t = TOKEN_Z80_NZ) _
        Then
            Let TokenIsParameter(t) = True
            Let TokenIsFlag(t) = True
        
        'If the token consititutes an Expression (a calculable Value)
        '..............................................................................
        ElseIf (t = TOKEN_PARENS_OPEN) Or _
               (t = TOKEN_CONST) Or _
               (t = TOKEN_HASH) Or _
               (t = TOKEN_KEYWORD_HI) Or _
               (t = TOKEN_LABEL) Or _
               (t = TOKEN_KEYWORD_LO) Or _
               (t = TOKEN_NUMBER) Or _
               (t = TOKEN_OPERATOR_SUB) Or _
               (t = TOKEN_PROPERTY_USE) Or _
               (t = TOKEN_RAM) Or _
               (t = TOKEN_SECTION) _
        Then
            Let TokenIsExpression(t) = True
            Let TokenIsParameter(t) = True
            Let TokenIsList(t) = True
        
        'Text classes as a self-contained List; _
         all other List items are Expressions
        '..............................................................................
        ElseIf (t = TOKEN_TEXT) Then
            Let TokenIsList(t) = True
        
        'Z80 parameters are Expressions, with the addition of _
         Z80 register / memory references, e.g. `[ix+$FF]`
        '..............................................................................
        ElseIf (t = TOKEN_SQUARE_OPEN) Then
            Let TokenIsParameter(t) = True
            
        'The `START` keyword acts as a Label for the System's boot vector
        '..............................................................................
        ElseIf (t = TOKEN_KEYWORD_START) Then
            Let TokenIsExpression(t) = True
            Let TokenIsParameter(t) = True
            Let TokenIsList(t) = True
        End If
    Next t
    
    '----------------------------------------------------------------------------------
    
    'Populate the look-up table for converting register tokens into their bit masks
    Let TokenRegister(TOKEN_Z80_A) = MASK_REG_A
    Let TokenRegister(TOKEN_Z80_B) = MASK_REG_B
    Let TokenRegister(TOKEN_Z80_C) = MASK_REG_C
    Let TokenRegister(TOKEN_Z80_D) = MASK_REG_D
    Let TokenRegister(TOKEN_Z80_E) = MASK_REG_E
    Let TokenRegister(TOKEN_Z80_H) = MASK_REG_H
    Let TokenRegister(TOKEN_Z80_L) = MASK_REG_L
    Let TokenRegister(TOKEN_Z80_I) = MASK_REG_I
    Let TokenRegister(TOKEN_Z80_R) = MASK_REG_R
    Let TokenRegister(TOKEN_Z80_AF) = MASK_REG_AF
    Let TokenRegister(TOKEN_Z80_BC) = MASK_REG_BC
    Let TokenRegister(TOKEN_Z80_DE) = MASK_REG_DE
    Let TokenRegister(TOKEN_Z80_HL) = MASK_REG_HL
    Let TokenRegister(TOKEN_Z80_SP) = MASK_REG_SP
    Let TokenRegister(TOKEN_Z80_IX) = MASK_REG_IX
    Let TokenRegister(TOKEN_Z80_IXH) = MASK_REG_IXH
    Let TokenRegister(TOKEN_Z80_IXL) = MASK_REG_IXL
    Let TokenRegister(TOKEN_Z80_IY) = MASK_REG_IY
    Let TokenRegister(TOKEN_Z80_IYH) = MASK_REG_IYH
    Let TokenRegister(TOKEN_Z80_IYL) = MASK_REG_IYL
    Let TokenRegister(TOKEN_Z80_NC) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_Z) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_NZ) = MASK_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_P) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PE) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PO) = MASK_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_M) = MASK_FLAGS_MP
    
    'Populate the look-up tables for converting z80 parameters into opcodes
    Let OpcodeP_RegisterPair(TOKEN_Z80_BC) = P0
    Let OpcodeP_RegisterPair(TOKEN_Z80_DE) = P1
    Let OpcodeP_RegisterPair(TOKEN_Z80_HL) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_IX) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_IY) = P2
    Let OpcodeP_RegisterPair(TOKEN_Z80_SP) = P3
    Let OpcodeP_RegisterPair(TOKEN_Z80_AF) = P3
    
    Let OpcodeY_Bit(0) = Y0
    Let OpcodeY_Bit(1) = Y1
    Let OpcodeY_Bit(2) = Y2
    Let OpcodeY_Bit(3) = Y3
    Let OpcodeY_Bit(4) = Y4
    Let OpcodeY_Bit(5) = Y5
    Let OpcodeY_Bit(6) = Y6
    Let OpcodeY_Bit(7) = Y7
    
    Let OpcodeY_Condition(TOKEN_Z80_NZ) = Y0
    Let OpcodeY_Condition(TOKEN_Z80_Z) = Y1
    Let OpcodeY_Condition(TOKEN_Z80_NC) = Y2
    Let OpcodeY_Condition(TOKEN_Z80_C) = Y3
    Let OpcodeY_Condition(TOKEN_Z80_PO) = Y4
    Let OpcodeY_Condition(TOKEN_Z80_PE) = Y5
    Let OpcodeY_Condition(TOKEN_Z80_P) = Y6
    Let OpcodeY_Condition(TOKEN_Z80_M) = Y7
    
    Let OpcodeY_Register(MASK_REG_B) = Y0
    Let OpcodeY_Register(MASK_REG_C) = Y1
    Let OpcodeY_Register(MASK_REG_D) = Y2
    Let OpcodeY_Register(MASK_REG_E) = Y3
    Let OpcodeY_Register(MASK_REG_H) = Y4
    Let OpcodeY_Register(MASK_REG_IXH) = Y4
    Let OpcodeY_Register(MASK_REG_IYH) = Y4
    Let OpcodeY_Register(MASK_REG_L) = Y5
    Let OpcodeY_Register(MASK_REG_IXL) = Y5
    Let OpcodeY_Register(MASK_REG_IYL) = Y5
    Let OpcodeY_Register(MASK_MEM_HL) = Y6
    Let OpcodeY_Register(MASK_MEM_IX) = Y6
    Let OpcodeY_Register(MASK_MEM_IY) = Y6
    Let OpcodeY_Register(MASK_REG_A) = Y7
    
    Let OpcodeY_RST(&H0) = Y0
    Let OpcodeY_RST(&H8) = Y1
    Let OpcodeY_RST(&H10) = Y2
    Let OpcodeY_RST(&H18) = Y3
    Let OpcodeY_RST(&H20) = Y4
    Let OpcodeY_RST(&H28) = Y5
    Let OpcodeY_RST(&H30) = Y6
    Let OpcodeY_RST(&H38) = Y7
    
    Let OpcodeZ_Register(MASK_REG_B) = Z0
    Let OpcodeZ_Register(MASK_REG_C) = Z1
    Let OpcodeZ_Register(MASK_REG_D) = Z2
    Let OpcodeZ_Register(MASK_REG_E) = Z3
    Let OpcodeZ_Register(MASK_REG_H) = Z4
    Let OpcodeZ_Register(MASK_REG_IXH) = Z4
    Let OpcodeZ_Register(MASK_REG_IYH) = Z4
    Let OpcodeZ_Register(MASK_REG_L) = Z5
    Let OpcodeZ_Register(MASK_REG_IXL) = Z5
    Let OpcodeZ_Register(MASK_REG_IYL) = Z5
    Let OpcodeZ_Register(MASK_MEM_HL) = Z6
    Let OpcodeZ_Register(MASK_MEM_IX) = Z6
    Let OpcodeZ_Register(MASK_MEM_IY) = Z6
    Let OpcodeZ_Register(MASK_REG_A) = Z7
    
    Call Profiler.EventEnd(EVENT_INITIALIZE)
    '*** TIMER END ***
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Call Reset
        
    'Erase the caches
    Erase ToUpperCase
    Erase Cache_HexStr8, Cache_HexStr16
    Erase TokenNames, ParamText
    
    'Free the look up tables for file tokens
    Erase TokenIsInstruction, TokenIsRegister, TokenIsRegister8Bit, _
          TokenIsRegister16Bit, TokenIsFlag, TokenIsKeyword, TokenIsList, _
          TokenIsOperator, TokenIsExpression, TokenIsParameter
    
    'Clear the token stream Text store
    Call TokenStrings.Clear
    Set TokenStrings = Nothing
    
    'Free the look-up tables for opcode building
    Erase OpcodeP_RegisterPair, OpcodeY_Bit, OpcodeY_Condition, _
          OpcodeY_Register, OpcodeY_RST, OpcodeZ_Register
    Erase TokenRegister
    
    Set LogMsg = Nothing
End Sub

'Reset
'======================================================================================
Private Sub Reset()
    'Dereference any currently attached source code File
    Let File = 0
    'Clear the database of source files
    Erase Files
    
    'Dereference the currently attached Chunk
    Let Chunk = 0
    'Dereference the fixed Chunks so that the array can be emptied without _
     keeping these ones alive
    Let Chunk_Header = 0
    Let Chunk_Interrupt00 = 0
    Let Chunk_Interrupt08 = 0
    Let Chunk_Interrupt10 = 0
    Let Chunk_Interrupt18 = 0
    Let Chunk_Interrupt20 = 0
    Let Chunk_Interrupt28 = 0
    Let Chunk_Interrupt30 = 0
    Let Chunk_Interrupt38 = 0
    Let Chunk_Interrupt66 = 0
    Let Chunk_Start = 0
    'Now erase the list of Chunks
    Erase Chunks: Let ChunkCount = 0
    
    'Reset the value stores
    Set SectionSlots = Nothing: Set SectionSlots = New Scripting.Dictionary
    Set Values = Nothing:       Set Values = New Scripting.Dictionary
    Set Labels = Nothing:       Set Labels = New Scripting.Dictionary
    
    Erase Layout
    Let LayoutCount = 0
    
    Erase ROM
    'TODO: Default ROM size should be per-system / user-selectable?
    Let ROMSize = 8192
End Sub

'Error : Expose the internal error number
'======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property

'Assemble : Compile a Z80 binary from a given starting source code file
'======================================================================================
'FilePath       | Source code file to begin with
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Public Function Assemble( _
    ByVal FilePath As String _
) As OZ80_ERROR
    'This will be a crude overall timer for the assembly
    Dim Begin As Single
    Let Begin = Timer
    
    'Reset the error number
    Let My_Error = OZ80_ERROR_NONE
    
    'Announce ourselves
    Call RaiseMessage( _
        OZ80_LOG_INFO, LOG_VERSION, _
        App.Major, App.Minor, App.Revision _
    )
    
    'Clean up stuff from any previous assemble
    Call Reset
    
    '----------------------------------------------------------------------------------
    'Phase 1: First Pass & Assemble _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE1)
    
    'Start with parsing the given file; the assembler does not work with the original _
     text directly. If there was a problem parsing the file, an error message will _
     already have been raised by an event, we only need to exit
    
    'Now walk the File that's been created, assembling it as we go. _
     If an `INCLUDE` is occurred, it will be tokenised and assembled there and then
    If File_Assemble(FilePath) Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    'Phase 2: Layout the ROM _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE2)
    
    'Now that the size of each Chunk is known, we can arrange the Chunks in the ROM _
     space and determine each Label location
    If LayoutROM() Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    'Phase 3: Second Parse & Assemble (now with ROM Addresses) _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE3)
    
    'TODO
    
    '----------------------------------------------------------------------------------
    'Phase 4: Post Processing, e.g. Checksum _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE4)
    
    'TODO
    
    'Finally, write the ROM to disk
'    Dim FileNumber As Integer: Let FileNumber = FreeFile
'    Open App.Path & "\ROM.sms" For Binary Access Write As #FileNumber
'    Put #FileNumber, , ROM
'    Close #FileNumber
    
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_ASM_COMPLETE)
    Call RaiseMessage(OZ80_LOG_INFO, LOG_TIME, Format$(Timer - Begin, "0.000"))
    
Finish:
    'Clean up
    Call Reset
    'Return the error code (also available in the `Error` property)
    Let Assemble = My_Error
End Function



'/// PARSING //////////////////////////////////////////////////////////////////////////
'/// Read in the source code and walk it

'File_Assemble : Tokenise a source code file and then begin assembling it
'======================================================================================
'FilePath       | File path to the source code file (".oz80")
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function File_Assemble( _
    ByRef FilePath As String _
) As OZ80_ERROR
    'Since we have a "scrrun.dll" reference for Dictionary object, this also gives _
     us the `FileSystemObject` to simplify a lot of mucky folder / file checking
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    '[1] Normalise the file path: _
     ----------------------------------------------------------------------------------
    'The file path handed to us might be just a file name and/or contain a relative _
     path (i.e. "..\"). We need to normalise this given path into an absolute path _
     (e.g. "C:\Folder\File.ext")
    Dim TestPath As bluString
    Set TestPath = New bluString
    
    'If the given path is already absolute, it does not need normalising. _
     NOTE: Upper/lower case will be auto-corrected and could throw a false negative
    If LCase$(FSO.GetAbsolutePathName(FilePath)) = LCase$(FilePath) Then
        'We'll skip ahead to checking the file exists
        Let TestPath.Text = FilePath
    Else
        'Path is relative! If no other file has been parsed yet _
         we will need to use the current directory as a base
        If File = 0 Then
            Let TestPath.Text = FSO.BuildPath(CurDir(), FilePath)
        Else
            'When another file has been parsed before, _
             we work relatively from that file's directory
            Let TestPath.Text = FSO.GetFile(Files(File).Path).ParentFolder.Path
            'TODO: Error?
            Let TestPath.Text = FSO.BuildPath(TestPath.Text, FilePath)
        End If
    End If
    
    '[2] Check file exists: _
     ----------------------------------------------------------------------------------
    If Not FSO.FileExists(TestPath.Text) Then
        'ERROR: File Not Found
        Call RaiseError(OZ80_ERROR_FILE_NOTFOUND)
        GoTo Finish
    End If
    
    '[3] Tokenise the file: _
     ----------------------------------------------------------------------------------
    'Check to see if the file has already been tokenised
    Dim Index As Long, i As Long
    For i = 1 To FileCount
        If Files(i).Path = TestPath.Text Then Let Index = i: Exit For
    Next
    
    'If the file has already been tokenised, re-use it
    If Index <> 0 Then
        'Reference the existing File object
        Let File = Index
    Else
        '*** TIMER BEGIN ***
        Call Profiler.EventBegin(EVENT_TOKENISE)
        
        'Add a File structure to hold the machine representation of the text file; _
         the assembler doesn't work with the original text directly
        Let FileCount = FileCount + 1: Let File = FileCount
        ReDim Preserve Files(1 To FileCount) As oz80File
        
        'Record the file path in the File class for easy reference in the future
        Let Files(File).Path = TestPath.Text
        
        'Explode the source code file into tokens:
        'If the file failed to parse, an error code / message will already have _
         been raised via an event, we only need to exit
        If File_Tokenise(TestPath.Text) Then GoTo Finish
        
        Call Profiler.EventEnd(EVENT_TOKENISE)
        '*** TIMER END ***
    End If
    
    'The user will obviously want to begin at the start of the stream _
     rather than the end
    Call File_Rewind
    
    '[4] Assemble the file: _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, "Assembling ""{$}""", FilePath)
    
    'Now begin assembling the file. If an `INCLUDE` statement is encountered, _
     it will also be tokenised and assembled
    Call ScopeRoot
    
    If My_Error Then GoTo Finish
    Call RaiseMessage(OZ80_LOG_INFO, "Assembly Complete for ""{$}""", FilePath)
    
Finish:
    Set FSO = Nothing
    Set TestPath = Nothing
    
    'If there were any error, an error event and message would already have been _
     raised, we only need to fold upwards and the parent will handle terminating the _
     assembling process early
    Let File_Assemble = My_Error
End Function

'======================================================================================
'======================================================================================
'======================================================================================
'======================================================================================
'======================================================================================
'======================================================================================
'======================================================================================
'======================================================================================

'EOF : End Of File?
'======================================================================================
Private Property Get EOF() As Boolean
    Let EOF = (TokenIndex >= Files(File).TokensLength)
End Property

'TokenText : Gets the string associated with a token (e.g. `TOKEN_TEXT`)
'======================================================================================
Private Property Get TokenText() As String
    'Only valid for tokens that reference text _
     (null-string will be returned otherwise)
    'TODO: Check for existance here
    Let TokenText = TokenStrings.Item( _
        Files(File).Tokens(TokenIndex).Value _
    )
End Property

'NextToken : Move onto the next token
'======================================================================================
'Returns        | False if the token stream hit the end
'======================================================================================
Private Function NextToken() As Boolean
'    'Generate a debug message with the details of the current token
'    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken())
    
    Let TokenIndex = TokenIndex + 1
    Call CacheToken
    
    If TokenIndex >= Files(File).TokensLength Then Exit Function
    Let NextToken = True
End Function

'File_Rewind : Go back to the start of the token stream
'======================================================================================
Private Sub File_Rewind()
    Let TokenIndex = 0
    Call CacheToken
End Sub

'File_Tokenise : Break the source file into words and convert to machine tokens
'======================================================================================
'FilePath       | Path to the source file to tokenise into this class instance
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function File_Tokenise( _
    ByRef FilePath As String _
) As OZ80_ERROR
    'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    
    '----------------------------------------------------------------------------------
    'Read the whole file into a single string
    '----------------------------------------------------------------------------------
    'Check the file even exists first...
    If Not FileExists(FilePath) _
        Then Call RaiseError(OZ80_ERROR_FILE_NOTFOUND): GoTo Finish
    
    'Read the file into our fast string handler (though we won't be doing any string- _
     manipulation, we will be relying on the various methods available to us)
    Dim Source As bluString: Set Source = New bluString
    'Note the string will not be increasing in size, so we can trim the buffer _
     down to the same length as the content
    'TODO: Handle file reading errors
    Call Source.ReadFile(FilePath).ReduceBuffer
    
    'Cache this for speed, we'll be referring to it every loop
    Dim SourceLen As Long: Let SourceLen = Source.Length
    
    Call RaiseMessage(OZ80_LOG_ACTION, _
        "Tokenising ""{$}""", FilePath _
    )
        
    '----------------------------------------------------------------------------------
    'Begin parsing words: _
     ----------------------------------------------------------------------------------
    'Prepare a large buffer to accept the token stream (using `ReDim` for every token _
     added is *incredibly* slow). If the tokens overflow this buffer, we'll increase _
     it by another chunk there and then
    Let Files(File).TokensBuffer = TOKEN_BUFFER
    Let Files(File).TokensLength = 0
    ReDim Files(File).Tokens(0 To TOKEN_BUFFER) As oz80Token
    '(call this after ReDimming, as it will cache the first token info)
    Call File_Rewind
    
    'Set the buffer for capturing the word to 1024 chars to avoid lots of memory _
     re-allocations if larger strings in the source are encountered
    Set Word = New bluString
    Let Word.Buffer = BUFFER_1K
    
    'These will track the "mode" of the text as we go _
     (including single, double, and triple quote)
    Dim IsText As Long, TextMode As Long
    Dim IsComment As Long, IsCommentMulti As Long
    
    'These are used whilst converting words into real numbers
    Dim Number As Double, Multiplier As Long
    
    'Which character in the source code we're parsing. Since the way the routine is _
     written (we increase the position before each word begins) begin with -1 so that _
     we may really begin at 0
    Dim SourcePos As Long
    Let SourcePos = -1
    
    'Line and column number in the source text. Used to record the position in the _
     original file of each word for producing helpful error messages. 1-based
    Dim Line As Long, Col As Long
    Let Line = 1: Let PrevLine = 1
    
    
    
    Dim Token As OZ80_TOKEN
    Dim Value As Double
    
    
    
    'Read a single word:
    '----------------------------------------------------------------------------------
ReadWord:
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_TOKENISE_READWORD)
    
    'Reset the captured word
    Call Word.Clear
SkipChar:
    'Move on to the next character in the source
    Let SourcePos = SourcePos + 1
ReadChar:
    'If the file ends?
    If SourcePos >= SourceLen Then
        'If no word is currently underway then we are all done
        If Word.IsNull Then GoTo Finish
        'Otherwise process this final word
        GoTo TokenWord
    End If
    
    'Read a character:
    Dim Char As OZ80_SYNTAX
    Let Char = Source.Char(SourcePos)
    
    'Keep track of line and column position in the source file _
     (for helpful error messages, even at the tokenising stage)
    '..................................................................................
    
    'New line:
    If Char = C0_LF Then
        Let Line = Line + 1: Let Col = 0
    
    'NOTE: Carriage-Returns are ignored to avoid doubling line number and to _
           maintain compatibility with UNIX line-endings which use Line-Feed only
    ElseIf Char = C0_CR Then
        GoTo SkipChar
    
    'Tab:
    ElseIf Char = C0_HT Then
        'Tab goes to the next tab stop boundary:
        Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH)
    
    ElseIf Source.CharIsVisible(SourcePos) Then
        Let Col = Col + 1
        
    Else
        'Non-printable characters are skipped and do not add to the column number
        GoTo SkipChar
    End If
    
    'Comments: ........................................................................
    If IsComment Then
        'For single-line comments:
        If Not IsCommentMulti Then
            'The end of a line ends a single-line comment:
            If Char = C0_LF Then Let IsComment = False: GoTo ReadWord
            'Is this a multi-line comment? (two backticks)
            'TODO: What is going on here???
            If Not Word.IsNull Then
                If Char = SYNTAX_COMMENT Then Let IsCommentMulti = True
                GoTo ReadWord
            End If
        
        'For multi-line comments, check for back-tick:
        ElseIf Char = SYNTAX_COMMENT Then
            'Check for a second back-tick following
            Let SourcePos = SourcePos + 1: Let Col = Col + 1
            If Source.Char(SourcePos) = SYNTAX_COMMENT Then
                Let IsComment = False: Let IsCommentMulti = False
                GoTo ReadWord
            End If
        End If
        'Skip to the next character in the comment
        GoTo SkipChar
        
    'Parsable Text: ...................................................................
    '(NOTE: Falsey checks are slightly faster)
    ElseIf IsText = 0 Then
        'Special handling for the beginning of a word
        If Word.IsNull Then
            'Skip whitespace before the word begins
            If Source.CharIsSpace(SourcePos) Then GoTo SkipChar
            
            'Punctuation characters that are only one letter long: _
             When we hit one of these we need to end the word immediately, but parsing _
             normally ends a word on the letter after, so we deal with the edge case
            If Char = SYNTAX_NEXT Then GoTo Punctuation
            If Char = SYNTAX_BRACES_OPEN Then GoTo Punctuation
            If Char = SYNTAX_BRACES_CLOSE Then GoTo Punctuation
            If Char = SYNTAX_PARENS_OPEN Then GoTo Punctuation
            If Char = SYNTAX_PARENS_CLOSE Then GoTo Punctuation
            If Char = SYNTAX_SQUARE_OPEN Then GoTo Punctuation
            If Char = SYNTAX_SQUARE_CLOSE Then GoTo Punctuation
            
            If Char = SYNTAX_OPERATOR_ADD Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_SUB Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_MUL Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_DIV Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_POW Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_MOD Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_OR Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_AND Then GoTo Punctuation
            If Char = SYNTAX_OPERATOR_XOR Then GoTo Punctuation
            
            'Is this text? (in which case, ignore wordbreaks until text end). _
             the opening quote is included so that we can recognise it for tokenising
            If Char = SYNTAX_TEXT Then
                Let IsText = True: Let TextMode = 0: GoTo NextChar
                'TODO
            End If
            
            If Char = SYNTAX_COMMENT Then
                Let IsComment = True: GoTo NextChar
            End If
        Else
            'A word is underway, check for end of line
            If Char = ASC_LF Then GoTo TokenWord
            
            'End the word on a space
            If Source.CharIsSpace(SourcePos) Then GoTo TokenWord
            
            'Symbols that can break a word in the middle without spaces to separate
            If Char = SYNTAX_NEXT Then GoTo WordBreak
            If Char = SYNTAX_BRACES_OPEN Then GoTo WordBreak
            If Char = SYNTAX_BRACES_CLOSE Then GoTo WordBreak
            If Char = SYNTAX_PARENS_OPEN Then GoTo WordBreak
            If Char = SYNTAX_PARENS_CLOSE Then GoTo WordBreak
            If Char = SYNTAX_SQUARE_OPEN Then GoTo WordBreak
            If Char = SYNTAX_SQUARE_CLOSE Then GoTo WordBreak
            
            If Char = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_POW Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_OR Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_AND Then GoTo WordBreak
            If Char = SYNTAX_OPERATOR_XOR Then GoTo WordBreak
            
            If Char = SYNTAX_TEXT Then GoTo WordBreak
            
            'If a comment begins at the end of a word without any space to separate, _
             we need to first keep the word before switching to comment parsing
            If Char = SYNTAX_COMMENT Then
                Let IsComment = True: GoTo WordBreak
            End If
        End If
        
    'Text: ............................................................................
    ElseIf TextMode = 0 Then
        'Check for multiple opening quotes
        If Word.Length = 1 And Char = SYNTAX_TEXT Then
            Let TextMode = TextMode + 1: GoTo SkipChar
        End If
        'Skip over the closing quote so that it doesn't get tokenised
        If Char = SYNTAX_TEXT Then Let IsText = False: GoTo TokenWord
    
    ElseIf TextMode = 1 Then
        'Check for multiple opening quotes
        If Word.Length = 1 And Char = SYNTAX_TEXT Then
            Let TextMode = TextMode + 1: GoTo SkipChar
        End If
        'End of line ends the text
        If Char = ASC_LF Then Let IsText = False: GoTo TokenWord
        
    ElseIf TextMode = 2 Then
        'Three quotes ends the text
        If Char = SYNTAX_TEXT Then Let TextMode = TextMode + 1: GoTo SkipChar
        
    ElseIf TextMode < 5 Then
        If Char = SYNTAX_TEXT Then Let TextMode = TextMode + 1: GoTo SkipChar
        Stop
    
    ElseIf TextMode = 5 Then
        'Triple text mode has ended
        Let IsText = False: GoTo TokenWord
    End If
    
NextChar:
    'Is this the first character of the word? _
     If so, record the starting position of the word
    If Word.IsNull Then Let TokenLine = Line: Let TokenCol = Col
    'Since this is a valid letter, capture it
    Call Word.CharPush(Char)
    GoTo SkipChar

    '----------------------------------------------------------------------------------
Punctuation:
    Call Word.Clear.CharPush(Char)
    Let TokenLine = Line: Let TokenCol = Col
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let SourcePos = SourcePos - 1: Let Col = Col - 1
    
    '----------------------------------------------------------------------------------
    'Tokenise the word:
    '----------------------------------------------------------------------------------
TokenWord:
    '*** TIMER END ***
    Call Profiler.EventEnd(EVENT_TOKENISE_READWORD)
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_TOKENISE_TOKENWORD)
    
    'Just a loop counter
    Dim i As Long
    
    'Extract the 1st, 2nd, 3rd & 4th letters of the word as these make up most _
     comparisons and can be done quickly without resorting to Strings
    Dim Asc1 As OZ80_SYNTAX, Asc2 As OZ80_SYNTAX, _
        Asc3 As OZ80_SYNTAX, Asc4 As OZ80_SYNTAX
    Let Asc1 = ToUpperCase(Word.Char(0))
    If Word.Length > 3 Then Let Asc4 = ToUpperCase(Word.Char(3)) Else Let Asc4 = ASC_NUL
    If Word.Length > 2 Then Let Asc3 = ToUpperCase(Word.Char(2)) Else Let Asc3 = ASC_NUL
    If Word.Length > 1 Then Let Asc2 = ToUpperCase(Word.Char(1)) Else Let Asc2 = ASC_NUL
    
    'Comma (very common!) _
     .......................................................................... `,` ...
    If Asc1 = SYNTAX_NEXT Then
        'NOTE: At this point we can only see one word at a time (and not look ahead), _
         so we cannot know when a list has started, and even if the grammer is valid _
         at all (could have a number in completely the wrong place). For that reason, _
         all commas are optional and simply skipped by the tokeniser
        GoTo SkipWord
        
    'A decimal number? _
     ........................................................................ `0-9` ...
    ElseIf Word.CharIsNumeric(0, True) Then
        'Validate the number
        Let Number = 0: Let Multiplier = 1
        For i = Word.Length - 1 To 0 Step -1
            'If not a 0-9 digit then this is not a decimal number!
            If Not Word.CharIsNumeric(i, True) Then
                Debug.Print Word.Text
                Stop
                
                'ERROR: Invalid number!
                Call RaiseError(OZ80_ERROR_INVALID_NUMBER_DEC)
                GoTo Finish
            Else
                'Convert the number from ASCII codes into a value
                On Error GoTo ErrorOverflow
                'I've tried a lookup table here and it's exactly the same speed
                Let Number = Number + (Word.Char(i) - ASC_0) * Multiplier
                'Don't allow the multiplier to overflow when the number doesn't
                If i > 0 Then Let Multiplier = Multiplier * 10
                On Error GoTo 0
            End If
        Next i
        'Add the token with the number value
        Let Token = TOKEN_NUMBER
        Let Value = Number
        GoTo NextWord
    
    'Constant? _
     .......................................................................... `!` ...
    ElseIf Asc1 = SYNTAX_CONSTANT Then
        'ERROR: Invalid hash name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_HASH)
            GoTo Finish
        End If
        
        'TODO: Define constant name
        'Add the index key to the token so that we can retrieve the string later
        Let Token = TOKEN_CONST
        Let Value = TokenStrings.Add(Word.Text)
        
        GoTo NextWord
    
    'Binary number? _
     .......................................................................... `%` ...
    ElseIf Asc1 = SYNTAX_NUMBER_BIN Then
        'ERROR: Cannot be longer than 32-bits
        If Word.Length > 33 Then GoTo ErrorOverflow
        Let Number = 0: Let Multiplier = 0
        For i = Word.Length - 1 To 1 Step -1
            If Word.Char(i) = ASC_1 Then
                Let Number = Number + 2 ^ Multiplier
            ElseIf Word.Char(i) <> ASC_0 Then
                'Error - not a valid binary number
                Stop
            End If
            Let Multiplier = Multiplier + 1
        Next i
        'Tokenise number
        Let Token = TOKEN_NUMBER
        Let Value = Number
        GoTo NextWord
        
    'Hash name? _
     .......................................................................... `#` ...
    ElseIf Asc1 = SYNTAX_HASH Then
        'ERROR: Invalid hash name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_HASH)
            GoTo Finish
        End If
        
        'TODO: Define hash name
        'Add the index key to the token so that we can retrieve the string later
        Let Token = TOKEN_HASH
        Let Value = TokenStrings.Add(Word.Text)
        GoTo NextWord
    
    'RAM name, e.g. `$.gamevar`, `$#RAM_BANK.gamevar` _
     .......................................................................... `$` ...
    ElseIf Asc1 = SYNTAX_RAM _
    And (Asc2 = ASC_NUL Or Asc2 = SYNTAX_PROPERTY Or Asc2 = SYNTAX_HASH) Then
        'ERROR: Invalid RAM name?
        If Not IsValidName() Then
            Call RaiseError(OZ80_ERROR_INVALID_NAME_RAM)
            GoTo Finish
        End If
        
        'TODO: Define RAM constant
        Let Token = TOKEN_RAM
        'Add the index key to the token so that we can retrieve the string later
        Let Value = TokenStrings.Add(Word.Text)
        GoTo NextWord
        
    'Hexadecimal number _
     .......................................................................... `$` ...
    ElseIf Asc1 = SYNTAX_NUMBER_HEX Then
        'Any more than 8-digits will be an overflow (32-bit number maximum)
        If Word.Length > 9 Then GoTo ErrorOverflow
        'Check the digits are 0-9, A-F
        For i = 1 To Word.Length - 1
            'ERROR: Invalid hexadecimal number!
            If Not Word.CharIsHex(i) Then Call RaiseError( _
                OZ80_ERROR_INVALID_NUMBER_HEX _
            ): GoTo Finish
        Next i
        
        'Now cast the string into a real number _
         (Note: The digits are valid, but it could be an overflow)
        On Error GoTo ErrorOverflow
        Let Number = Val("&H" & Mid$(Word.Text, 2) & "&")
        On Error GoTo 0
        
        'Add the token for the number
        Let Token = TOKEN_NUMBER
        Let Value = Number
        GoTo NextWord
        
    '.......................................................................... `:` ...
    ElseIf Asc1 = SYNTAX_LABEL Then
        'A label can be `:` only, without a name (when used within a PROC)
        If Asc2 = ASC_NUL Then
            Let Token = TOKEN_LABEL
            Let Value = TokenStrings.Add(Word.Text)
        
        'A label cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf Not IsValidName() Then
            'TODO: Label specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
        Else
            'Is this a section name (`::section`) or regular label (`:label`)?
            If Asc2 = SYNTAX_LABEL _
                Then Let Token = TOKEN_SECTION _
                Else Let Token = TOKEN_LABEL
            Let Value = TokenStrings.Add(Word.Text)
        End If
        GoTo NextWord
    
    '.......................................................................... `.` ...
    ElseIf Asc1 = SYNTAX_PROPERTY Then
        'At this stage the grammer is not validated, so we have no idea which label _
         this property attaches to. We store the name in an array and add the index _
         to the token stream -- the assembler will handle the context for same-name _
         properties
        
        'A property can be `.` only, without a name
        If Asc2 = ASC_NUL Then
            Let Token = TOKEN_PROPERTY_NEW
            Let Value = TokenStrings.Add(Word.Text)
            GoTo NextWord
        
        'A property cannot begin with a number but must otherwise be A-Z, 0-9 and "_"
        ElseIf Not IsValidName() Then
            'TODO: Property specific error message
            Call RaiseError(OZ80_ERROR_INVALID_NAME)
            GoTo Finish
        Else
            'If the property appears on a new line, it's a label-property definition _
             i.e. it defines a new sub-label of a chunk
            If TokenLine <> PrevLine _
                Then Let Token = TOKEN_PROPERTY_NEW _
                Else Let Token = TOKEN_PROPERTY_USE
            Let Value = TokenStrings.Add(Word.Text)
        End If
        GoTo NextWord
        
    'Text _
     .......................................................................... `"` ...
    ElseIf Asc1 = SYNTAX_TEXT Then
        Let Token = TOKEN_TEXT
        Let Value = TokenStrings.Add(Word.Mid(1).Text)
        GoTo NextWord
        
    ElseIf Asc1 = ASC_A Then '................................................. `a` ...
        'Z80 instructions
        If Word.Equals(STR_ADC) Then Let Token = TOKEN_Z80_ADC: GoTo NextWord
        If Word.Equals(STR_ADD) Then Let Token = TOKEN_Z80_ADD: GoTo NextWord
        If Word.Equals(STR_AND) Then Let Token = TOKEN_Z80_AND: GoTo NextWord
         
    ElseIf Asc1 = ASC_B Then '................................................. `b` ...
        'Z80 instructions
        If Word.Equals(STR_BIT) Then Let Token = TOKEN_Z80_BIT: GoTo NextWord
        'Keywords
        If Word.Equals(STR_BOOL) Then Let Token = TOKEN_KEYWORD_BOOL: GoTo NextWord
        If Word.Equals(STR_BYTE) Then Let Token = TOKEN_KEYWORD_BYTE: GoTo NextWord
        
    ElseIf Asc1 = ASC_C Then '................................................. `c` ...
        'Z80 instructions
        If Word.Equals(STR_CALL) Then Let Token = TOKEN_Z80_CALL: GoTo NextWord
        If Word.Equals(STR_CCF) Then Let Token = TOKEN_Z80_CCF: GoTo NextWord
        If Word.Equals(STR_CP) Then Let Token = TOKEN_Z80_CP: GoTo NextWord
        If Word.Equals(STR_CPD) Then Let Token = TOKEN_Z80_CPD: GoTo NextWord
        If Word.Equals(STR_CPDR) Then Let Token = TOKEN_Z80_CPDR: GoTo NextWord
        If Word.Equals(STR_CPI) Then Let Token = TOKEN_Z80_CPI: GoTo NextWord
        If Word.Equals(STR_CPIR) Then Let Token = TOKEN_Z80_CPIR: GoTo NextWord
        If Word.Equals(STR_CPL) Then Let Token = TOKEN_Z80_CPL: GoTo NextWord
        'Z80 flag condition
        If Asc2 = ASC_NUL Then Let Token = TOKEN_Z80_C: GoTo NextWord
        
    ElseIf Asc1 = ASC_D Then '................................................. `d` ...
        'Z80 instructions
        If Word.Equals(STR_DAA) Then Let Token = TOKEN_Z80_DAA: GoTo NextWord
        If Word.Equals(STR_DEC) Then Let Token = TOKEN_Z80_DEC: GoTo NextWord
        If Word.Equals(STR_DI) Then Let Token = TOKEN_Z80_DI: GoTo NextWord
        If Word.Equals(STR_DJNZ) Then Let Token = TOKEN_Z80_DJNZ: GoTo NextWord
        'Keywords
        If Word.Equals(STR_DEF) Then Let Token = TOKEN_KEYWORD_DEF: GoTo NextWord
        
    ElseIf Asc1 = ASC_E Then '................................................. `e` ...
        'Z80 instructions
        If Word.Equals(STR_EI) Then Let Token = TOKEN_Z80_EI: GoTo NextWord
        If Word.Equals(STR_EX) Then Let Token = TOKEN_Z80_EX: GoTo NextWord
        If Word.Equals(STR_EXX) Then Let Token = TOKEN_Z80_EXX: GoTo NextWord
        
    ElseIf Asc1 = ASC_F Then '................................................. `f` ...
        'You can't refer to the `f` register directly (only `af`)
        GoTo WordError
        
    ElseIf Asc1 = ASC_G Then '................................................. `g` ...
        'There's nothing that begins with G
        GoTo WordError
        
    ElseIf Asc1 = ASC_H Then '................................................. `h` ...
        'Z80 instructions
        If Word.Equals(STR_HLT) Then Let Token = TOKEN_Z80_HALT: GoTo NextWord
        If Word.Equals(STR_HALT) Then Let Token = TOKEN_Z80_HALT: GoTo NextWord
        'Keywords
        If Word.Equals(STR_HASH) Then Let Token = TOKEN_KEYWORD_HASH: GoTo NextWord
        If Word.Equals(STR_HELP) Then Let Token = TOKEN_KEYWORD_HELP: GoTo NextWord
        If Word.Equals(STR_HI) Then Let Token = TOKEN_KEYWORD_HI: GoTo NextWord
        
    ElseIf Asc1 = ASC_I Then '................................................. `i` ...
        'Z80 instructions
        If Word.Equals(STR_IM) Then Let Token = TOKEN_Z80_IM: GoTo NextWord
        If Word.Equals(STR_IN) Then Let Token = TOKEN_Z80_IN: GoTo NextWord
        If Word.Equals(STR_INC) Then Let Token = TOKEN_Z80_INC: GoTo NextWord
        If Word.Equals(STR_IND) Then Let Token = TOKEN_Z80_IND: GoTo NextWord
        If Word.Equals(STR_INDR) Then Let Token = TOKEN_Z80_INDR: GoTo NextWord
        If Word.Equals(STR_INI) Then Let Token = TOKEN_Z80_INI: GoTo NextWord
        If Word.Equals(STR_INIR) Then Let Token = TOKEN_Z80_INIR: GoTo NextWord
        'keywords
        If Word.Equals(STR_INCLUDE) Then Let Token = TOKEN_KEYWORD_INCLUDE: GoTo NextWord
        If Word.Equals(STR_INDEX) Then Let Token = TOKEN_KEYWORD_INDEX: GoTo NextWord
        If Word.Equals(STR_INTERRUPT) Then Let Token = TOKEN_KEYWORD_INTERRUPT: GoTo NextWord
        
    ElseIf Asc1 = ASC_J Then '................................................. `j` ...
        'Z80 instructions
        If Word.Equals(STR_JP) Then Let Token = TOKEN_Z80_JP: GoTo NextWord
        If Word.Equals(STR_JR) Then Let Token = TOKEN_Z80_JR: GoTo NextWord
        'There are no `j` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_K Then '................................................. `k` ...
        'Number suffixes K & KB
        If Asc2 = ASC_NUL Then
            'Just `K`
            Let Token = TOKEN_PREFIX_K: GoTo NextWord
        
        'B must be next for `KB`
        ElseIf Asc2 = ASC_B And Asc3 = ASC_NUL Then
            Let Token = TOKEN_PREFIX_KB: GoTo NextWord
            
        End If
        'There are no `k` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_L Then '................................................. `l` ...
        'Z80 instructions
        If Word.Equals(STR_LD) Then Let Token = TOKEN_Z80_LD: GoTo NextWord
        If Word.Equals(STR_LDD) Then Let Token = TOKEN_Z80_LDD: GoTo NextWord
        If Word.Equals(STR_LDDR) Then Let Token = TOKEN_Z80_LDDR: GoTo NextWord
        If Word.Equals(STR_LDI) Then Let Token = TOKEN_Z80_LDI: GoTo NextWord
        If Word.Equals(STR_LDIR) Then Let Token = TOKEN_Z80_LDIR: GoTo NextWord
        'Keywords
        If Word.Equals(STR_LO) Then Let Token = TOKEN_KEYWORD_LO: GoTo NextWord
        If Word.Equals(STR_LONG) Then Let Token = TOKEN_KEYWORD_LONG: GoTo NextWord
    
    ElseIf Asc1 = ASC_M Then '................................................. `m` ...
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Let Token = TOKEN_Z80_M: GoTo NextWord
        'There is no `m` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_N Then '................................................. `n` ...
        'Z80 instructions
        If Word.Equals(STR_NEG) Then Let Token = TOKEN_Z80_NEG: GoTo NextWord
        If Word.Equals(STR_NOP) Then Let Token = TOKEN_Z80_NOP: GoTo NextWord
        'Keywords
        If Word.Equals(STR_NYBL) Then Let Token = TOKEN_KEYWORD_NYBL: GoTo NextWord
        'Z80 flag conditions
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_C Then Let Token = TOKEN_Z80_NC: GoTo NextWord
            If Asc2 = ASC_Z Then Let Token = TOKEN_Z80_NZ: GoTo NextWord
        End If
        'There are no `n` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_O Then '................................................. `o` ...
        'Z80 instructions
        If Word.Equals(STR_OR) Then Let Token = TOKEN_Z80_OR: GoTo NextWord
        If Word.Equals(STR_OUT) Then Let Token = TOKEN_Z80_OUT: GoTo NextWord
        If Word.Equals(STR_OUTD) Then Let Token = TOKEN_Z80_OUTD: GoTo NextWord
        If Word.Equals(STR_OTDR) Then Let Token = TOKEN_Z80_OTDR: GoTo NextWord
        If Word.Equals(STR_OUTI) Then Let Token = TOKEN_Z80_OUTI: GoTo NextWord
        If Word.Equals(STR_OTIR) Then Let Token = TOKEN_Z80_OTIR: GoTo NextWord
        'There are no `o` registers
        GoTo WordError
        
    ElseIf Asc1 = ASC_P Then '................................................. `p` ...
        'Z80 instructions
        If Word.Equals(STR_POP) Then Let Token = TOKEN_Z80_POP: GoTo NextWord
        If Word.Equals(STR_PUSH) Then Let Token = TOKEN_Z80_PUSH: GoTo NextWord
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Let Token = TOKEN_Z80_P: GoTo NextWord
        If Asc3 = ASC_NUL Then
            If Asc2 = ASC_E Then Let Token = TOKEN_Z80_PE: GoTo NextWord
            If Asc2 = ASC_O Then Let Token = TOKEN_Z80_PO: GoTo NextWord
        End If
        'Keywords
        If Word.Equals(STR_PARAMS) Then Let Token = TOKEN_KEYWORD_PARAMS: GoTo NextWord
        If Word.Equals(STR_PROC) Then Let Token = TOKEN_KEYWORD_PROC: GoTo NextWord
    
    ElseIf Asc1 = ASC_Q Then '................................................. `q` ...
        'There's nothing that begins with Q
        GoTo WordError
        
    ElseIf Asc1 = ASC_R Then '................................................. `r` ...
        'Z80 instructions
        If Word.Equals(STR_RES) Then Let Token = TOKEN_Z80_RES: GoTo NextWord
        If Word.Equals(STR_RET) Then Let Token = TOKEN_Z80_RET: GoTo NextWord
        If Word.Equals(STR_RETI) Then Let Token = TOKEN_Z80_RETI: GoTo NextWord
        If Word.Equals(STR_RETN) Then Let Token = TOKEN_Z80_RETN: GoTo NextWord
        If Word.Equals(STR_RL) Then Let Token = TOKEN_Z80_RL: GoTo NextWord
        If Word.Equals(STR_RLA) Then Let Token = TOKEN_Z80_RLA: GoTo NextWord
        If Word.Equals(STR_RLC) Then Let Token = TOKEN_Z80_RLC: GoTo NextWord
        If Word.Equals(STR_RLCA) Then Let Token = TOKEN_Z80_RLCA: GoTo NextWord
        If Word.Equals(STR_RLD) Then Let Token = TOKEN_Z80_RLD: GoTo NextWord
        If Word.Equals(STR_RR) Then Let Token = TOKEN_Z80_RR: GoTo NextWord
        If Word.Equals(STR_RRA) Then Let Token = TOKEN_Z80_RRA: GoTo NextWord
        If Word.Equals(STR_RRC) Then Let Token = TOKEN_Z80_RRC: GoTo NextWord
        If Word.Equals(STR_RRCA) Then Let Token = TOKEN_Z80_RRCA: GoTo NextWord
        If Word.Equals(STR_RRD) Then Let Token = TOKEN_Z80_RRD: GoTo NextWord
        If Word.Equals(STR_RST) Then Let Token = TOKEN_Z80_RST: GoTo NextWord
        'Keywords
        If Word.Equals(STR_RAM) Then Let Token = TOKEN_KEYWORD_RAM: GoTo NextWord
        If Word.Equals(STR_RETURN) Then Let Token = TOKEN_KEYWORD_RETURN: GoTo NextWord
        
    ElseIf Asc1 = ASC_S Then '................................................. `s` ...
        'Z80 instructions
        If Word.Equals(STR_SBC) Then Let Token = TOKEN_Z80_SBC: GoTo NextWord
        If Word.Equals(STR_SCF) Then Let Token = TOKEN_Z80_SCF: GoTo NextWord
        If Word.Equals(STR_SET) Then Let Token = TOKEN_Z80_SET: GoTo NextWord
        If Word.Equals(STR_SLA) Then Let Token = TOKEN_Z80_SLA: GoTo NextWord
        If Word.Equals(STR_SLL) Then Let Token = TOKEN_Z80_SLL: GoTo NextWord
        If Word.Equals(STR_SRA) Then Let Token = TOKEN_Z80_SRA: GoTo NextWord
        If Word.Equals(STR_SRL) Then Let Token = TOKEN_Z80_SRL: GoTo NextWord
        If Word.Equals(STR_SUB) Then Let Token = TOKEN_Z80_SUB: GoTo NextWord
        'Keywords
        If Word.Equals(STR_SECTION) Then Let Token = TOKEN_KEYWORD_SECTION: GoTo NextWord
        If Word.Equals(STR_SLOT) Then Let Token = TOKEN_KEYWORD_SLOT: GoTo NextWord
        If Word.Equals(STR_START) Then Let Token = TOKEN_KEYWORD_START: GoTo NextWord
        If Word.Equals(STR_SYSTEM) Then Let Token = TOKEN_KEYWORD_SYSTEM: GoTo NextWord
        
    ElseIf Asc1 = ASC_T Then '................................................. `t` ...
        'Keywords
        If Word.Equals(STR_TABLE) Then Let Token = TOKEN_KEYWORD_TABLE: GoTo NextWord
        If Word.Equals(STR_TRIP) Then Let Token = TOKEN_KEYWORD_TRIP: GoTo NextWord
        'There is no `t` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_U Then '................................................. `u` ...
        'There's nothing that begins with U
        GoTo WordError
    
    ElseIf Asc1 = ASC_V Then '................................................. `v` ...
        'There's nothing that begins with V
        GoTo WordError
        
    ElseIf Asc1 = ASC_W Then '................................................. `w` ...
        'Keywords
        If Word.Equals(STR_WORD) Then Let Token = TOKEN_KEYWORD_WORD: GoTo NextWord
        'There is no `w` register
        GoTo WordError
        
    ElseIf Asc1 = ASC_X Then '................................................. `x` ...
        'If only `x`, then it's the repeat operator
        If Asc2 = ASC_NUL Then Let Token = TOKEN_OPERATOR_REP: GoTo NextWord
        'Z80 instructions
        If Word.Equals(STR_XOR) Then Let Token = TOKEN_Z80_XOR: GoTo NextWord
        'There is no `x` register
        GoTo WordError
    
    ElseIf Asc1 = ASC_Y Then '................................................. `y` ...
        'There's nothing that begins with Y
        GoTo WordError
    
    ElseIf Asc1 = ASC_Z Then '................................................. `z` ...
        'Z80 flag conditions
        If Asc2 = ASC_NUL Then Let Token = TOKEN_Z80_Z: GoTo NextWord
        'There is no 'z' register
        GoTo WordError
        
    ElseIf Asc1 = SYNTAX_BRACES_OPEN Then '.................................... `{` ...
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_BRACES_OPEN) Then GoTo Finish
        Let Token = TOKEN_BRACES_OPEN: GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_BRACES_CLOSE Then  '.................................. `}` ...
        'Close the current nesting level, error if incorrect bracket-matching
        Let Token = TOKEN_BRACES_CLOSE: Let Value = Nesting(NestingDepth).Index
        If NestOut(TOKEN_BRACES_OPEN, TOKEN_BRACES_CLOSE) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_PARENS_OPEN Then '.................................... `(` ...
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_PARENS_OPEN) Then GoTo Finish
        Let Token = TOKEN_PARENS_OPEN: GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_PARENS_CLOSE Then '................................... `)` ...
        'Close the current nesting level, error if incorrect bracket-matching
        Let Token = TOKEN_PARENS_CLOSE: Let Value = Nesting(NestingDepth).Index
        If NestOut(TOKEN_PARENS_OPEN, TOKEN_PARENS_CLOSE) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_SQUARE_OPEN Then '.................................... `[` ...
        'Open another nesting level (if there was an error, stop parsing)
        If NestIn(TOKEN_SQUARE_OPEN) Then GoTo Finish
        Let Token = TOKEN_SQUARE_OPEN: GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_SQUARE_CLOSE Then '................................... `]` ...
        'Close the current nesting level, error if incorrect bracket-matching
        Let Token = TOKEN_SQUARE_CLOSE: Let Value = Nesting(NestingDepth).Index
        If NestOut(TOKEN_SQUARE_OPEN, TOKEN_SQUARE_CLOSE) Then GoTo Finish
        GoTo NextWord
        
    ElseIf Asc1 = SYNTAX_OPERATOR_ADD Then '................................... `+` ...
        Let Token = TOKEN_OPERATOR_ADD: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_SUB Then '................................... `-` ...
        Let Token = TOKEN_OPERATOR_SUB: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MUL Then '................................... `*` ...
        Let Token = TOKEN_OPERATOR_MUL: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_DIV Then '................................... `/` ...
        Let Token = TOKEN_OPERATOR_DIV: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_POW Then '................................... `^` ...
        Let Token = TOKEN_OPERATOR_POW: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_MOD Then '................................... `\` ...
        Let Token = TOKEN_OPERATOR_MOD: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_OR Then  '................................... `|` ...
        Let Token = TOKEN_OPERATOR_OR: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_AND Then '................................... `&` ...
        Let Token = TOKEN_OPERATOR_AND: GoTo NextWord
    ElseIf Asc1 = SYNTAX_OPERATOR_XOR Then '................................... `~` ...
        Let Token = TOKEN_OPERATOR_XOR: GoTo NextWord
    End If
    
    'Z80 Registers: _
     ----------------------------------------------------------------------------------
Registers:
    'Ignore anything beginning with a letter before A
    If Asc1 < ASC_A Then GoTo WordError
    'No register has a first letter higher than S
    If Asc1 > ASC_S Then GoTo WordError
    'If there's only 1 letter total, then it's obviously a 1-letter register
    If Asc2 = ASC_NUL Then GoTo OneLetter
    'If a semi-colon occurs at the second letter, then the register must be 1 letter
    If Asc2 = SYNTAX_HINT1 Then GoTo OneLetter
    'Also an apostrophe (shadow register) is the same
    If Asc2 = SYNTAX_HINT2 Then GoTo OneLetter
    
    'If there's only 2 letters in total, it has to be a two-letter register
    If Asc3 = ASC_NUL Then GoTo TwoLetters
    'If a semi-colon or apostrophe occurs at the third letter, _
     it could be a two-letter register
    If Asc3 = SYNTAX_HINT1 Then GoTo TwoLetters
    If Asc3 = SYNTAX_HINT2 Then GoTo TwoLetters
    
    'with all else failing, is it three letters?
    If Asc4 = ASC_NUL Then GoTo ThreeLetters
    If Asc4 = SYNTAX_HINT1 Then GoTo ThreeLetters
    If Asc4 = SYNTAX_HINT2 Then GoTo ThreeLetters
    
    GoTo WordError

OneLetter:
    If Asc1 = ASC_A Then Let Token = TOKEN_Z80_A: GoTo NextWord
    If Asc1 = ASC_B Then Let Token = TOKEN_Z80_B: GoTo NextWord
    If Asc1 = ASC_C Then Let Token = TOKEN_Z80_C: GoTo NextWord
    If Asc1 = ASC_D Then Let Token = TOKEN_Z80_D: GoTo NextWord
    If Asc1 = ASC_E Then Let Token = TOKEN_Z80_E: GoTo NextWord
    If Asc1 = ASC_H Then Let Token = TOKEN_Z80_H: GoTo NextWord
    If Asc1 = ASC_I Then Let Token = TOKEN_Z80_I: GoTo NextWord
    If Asc1 = ASC_L Then Let Token = TOKEN_Z80_L: GoTo NextWord
    If Asc1 = ASC_R Then Let Token = TOKEN_Z80_R: GoTo NextWord
    GoTo WordError

TwoLetters:
    If Asc1 = ASC_A Then                '"af"
        If Asc2 = ASC_F Then Let Token = TOKEN_Z80_AF: GoTo NextWord
    ElseIf Asc1 = ASC_B Then            '"bc"
        If Asc2 = ASC_C Then Let Token = TOKEN_Z80_BC: GoTo NextWord
    ElseIf Asc1 = ASC_D Then            '"de"
        If Asc2 = ASC_E Then Let Token = TOKEN_Z80_DE: GoTo NextWord
    ElseIf Asc1 = ASC_H Then            '"hl"
        If Asc2 = ASC_L Then Let Token = TOKEN_Z80_HL: GoTo NextWord
    ElseIf Asc1 = ASC_I Then            '"ix"/"iy"
        If Asc2 = ASC_X Then Let Token = TOKEN_Z80_IX: GoTo NextWord
        If Asc2 = ASC_Y Then Let Token = TOKEN_Z80_IY: GoTo NextWord
    ElseIf Asc1 = ASC_P Then            '"pc"
        If Asc2 = ASC_C Then Let Token = TOKEN_Z80_PC: GoTo NextWord
    ElseIf Asc1 = ASC_S Then            '"sp"
        If Asc2 = ASC_P Then Let Token = TOKEN_Z80_SP: GoTo NextWord
    End If
    GoTo WordError
    
ThreeLetters:
    If Asc1 = ASC_I Then
        If Asc2 = ASC_X Then
            If Asc3 = ASC_H Then Let Token = TOKEN_Z80_IXH: GoTo NextWord
            If Asc3 = ASC_L Then Let Token = TOKEN_Z80_IXL: GoTo NextWord
        ElseIf Asc2 = ASC_Y Then
            If Asc3 = ASC_H Then Let Token = TOKEN_Z80_IYH: GoTo NextWord
            If Asc3 = ASC_L Then Let Token = TOKEN_Z80_IYL: GoTo NextWord
        End If
    End If
    GoTo WordError
    
    '----------------------------------------------------------------------------------
NextWord:
    With Files(File)
        'Add the information gleaned to the current token and move to the next empty slot
        With .Tokens(TokenIndex)
            Let .Kind = Token
            Let .Line = TokenLine
            Let .Col = TokenCol
            Let .Value = Value
        End With
        
        'Remember the line number used for this token to check for properties occuring _
         on a new line -- this constitutes a property definition in a Chunk
        Let PrevLine = TokenLine
        
'        'TODO: Not necessary here?
'        Call CacheToken
        
        'Move to the next token number
        Let TokenIndex = TokenIndex + 1

        'This will also likely increase the total number of tokens
        Let .TokensLength = TokenIndex

        'Note that the tokens array is given a large size to begin with as ReDimming _
         it with every token added is *incredibly* slow. If we go over this buffer _
         size, we need to increase it by a large chunk. Once parsing is complete the _
         buffer will be trimmed down to the final size
        If .TokensLength > .TokensBuffer Then
            Let .TokensBuffer = .TokensBuffer + TOKEN_BUFFER
            ReDim Preserve .Tokens(0 To .TokensBuffer) As oz80Token
        End If
    End With

'    'Generate a debug message
'    RaiseEvent Message(OZ80_LOG_DEBUG, FormatToken())
        
SkipWord:
    Call Profiler.EventEnd(EVENT_TOKENISE_TOKENWORD)
    '*** TIMER END ***
    GoTo ReadWord

ErrorOverflow: '.......................................................................
    On Error GoTo 0
    Call RaiseError(OZ80_ERROR_OVERFLOW)
    GoTo Finish
    
WordError: '...........................................................................
    Call RaiseError(OZ80_ERROR_INVALID_WORD)
    
    '----------------------------------------------------------------------------------
Finish:
    If My_Error = OZ80_ERROR_NONE Then
        'Reduce the buffer to the final size
        Let Files(File).TokensBuffer = Files(File).TokensLength
        ReDim Preserve Files(File).Tokens(0 To Files(File).TokensLength) As oz80Token
        'Move the cursor back to the start of the stream
        Call File_Rewind
        
        'Log the results of the tokenisation
        Call RaiseMessage(OZ80_LOG_INFO, LOG_FILE_LENGTH, Format$(SourceLen, "#,0"))
        Call RaiseMessage(OZ80_LOG_INFO, LOG_FILE_TOKENS, _
            Format$(Files(File).TokensLength + 1, "#,0") _
        )
    Else
        'clean up!
        'TODO: This leaves the class in a broken state _
         (no tokens array anymore and no call to `cacheToken`)
        Erase Files(File).Tokens: Let Files(File).TokensLength = 0
    End If
    
    'Free up the memory used for the text file _
     (we won't be referring to the original text again)
    Set Source = Nothing
    'Clear the nesting information
    Erase Nesting: Let NestingDepth = 0
    
    'Return the error number to the assembler
    Let File_Tokenise = My_Error
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'CacheToken : Cache the current token info
'======================================================================================
Private Sub CacheToken()
    If TokenIndex > Files(File).TokensLength Then
        Let TokenKind = 0
        Let TokenValue = 0
        Let TokenLine = 0
        Let TokenCol = 0
        Exit Sub
    End If
    With Files(File).Tokens(TokenIndex)
        Let TokenKind = .Kind
        Let TokenValue = .Value
        Let TokenLine = .Line
        Let TokenCol = .Col
    End With
End Sub

'FormatToken : Format the details of the current token for logging
'======================================================================================
'Returns        | A new bluString containing the details of the current token
'======================================================================================
Private Function FormatToken( _
) As bluString
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_FORMATTOKEN)
    
    'Set the template we'll be putting the data into
    Dim Template As New bluString
    Let Template.Text = FORMAT_TOKEN
    
    'Format the Token number
    Dim Format_TokenIndex As New bluString
    Let Format_TokenIndex.Text = Format$(Format$(TokenIndex + 1, "#,#"), "@@@@@@@")
    
    Dim Format_Msg As New bluString
    Dim Format_TokenType As String
    
    If TokenKind = TOKEN_BRACES_OPEN _
    Or TokenKind = TOKEN_BRACES_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_CHUNK
        
    ElseIf TokenKind = TOKEN_PARENS_OPEN _
        Or TokenKind = TOKEN_PARENS_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_EXPR
        
    ElseIf TokenKind = TOKEN_SQUARE_OPEN _
        Or TokenKind = TOKEN_SQUARE_CLOSE Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_SQUARE
        
    ElseIf TokenKind = TOKEN_CONST Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_CONST
        
    ElseIf TokenIsFlag(TokenKind) Then
        '.............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_FLAG
        
    ElseIf TokenKind = TOKEN_HASH Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_SQUARE
        
    ElseIf TokenIsInstruction(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_Z80
        
    ElseIf TokenIsKeyword(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Call Format_Msg.UCase
        Let Format_TokenType = FORMAT_TOKEN_KEYWORD
        
    ElseIf TokenKind = TOKEN_LABEL Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_LABEL
        
    ElseIf TokenKind = TOKEN_NUMBER Then
        '..............................................................................
        'TODO: Don't try and display an indefinite value
        Let Format_Msg.Text = Format$(TokenValue, "#,0")
        Let Format_TokenType = FORMAT_TOKEN_NUMBER
        
    ElseIf TokenIsOperator(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_OPERATOR
        
    ElseIf TokenKind = TOKEN_PREFIX_K _
        Or TokenKind = TOKEN_PREFIX_KB Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_PREFIX
        
    ElseIf TokenKind = TOKEN_PROPERTY_NEW _
        Or TokenKind = TOKEN_PROPERTY_USE Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_PROPERTY
            
    ElseIf TokenKind = TOKEN_RAM Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_RAM
        
    ElseIf TokenIsRegister(TokenKind) Then
        '..............................................................................
        Let Format_Msg.Text = TokenNames(TokenKind)
        Let Format_TokenType = FORMAT_TOKEN_REGISTER
        
    ElseIf TokenKind = TOKEN_SECTION Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_SECTION
        
    ElseIf TokenKind = TOKEN_TEXT Then
        '..............................................................................
        Let Format_Msg.Text = TokenStrings.Item(TokenValue)
        Call Format_Msg.Wrap(FORMAT_QUOTE)
        Let Format_TokenType = FORMAT_TOKEN_TEXT
    
    Else
        Debug.Print TokenNames(TokenKind)
        Stop
    End If
    
    Set FormatToken = Template.Format( _
        Format_TokenIndex, _
        Right$("    " & TokenLine, 5), Left$(TokenCol & "  ", 3), _
        Format_TokenType, _
        Format_Msg _
    )
        
    '*** TIMER END ***
    Call Profiler.EventEnd(EVENT_FORMATTOKEN)
End Function

'IsValidName : Check validity for named items
'======================================================================================
'Returns        | True if the captured word conforms to name item syntax
'======================================================================================
Private Function IsValidName( _
) As Boolean
    'Variable, label and property names can contain A-Z, 0-9 underscore and dot with _
     the exceptions that the first letter cannot be a number or a dot, two dots cannot _
     occur in a row nor can a number follow a dot and the name cannot end in a dot
    Dim WordChar0 As OZ80_SYNTAX
    Let WordChar0 = Word.Char(0)
    
    'The first letter must be the type identifier; _
     `!CONST`, `:label`, `.property`, `#hash`, `$.RAM`, `?function`, `@macro`
    If WordChar0 <> SYNTAX_CONSTANT And _
       WordChar0 <> SYNTAX_LABEL And _
       WordChar0 <> SYNTAX_PROPERTY And _
       WordChar0 <> SYNTAX_HASH And _
       WordChar0 <> SYNTAX_RAM And _
       WordChar0 <> SYNTAX_FUNCT And _
       WordChar0 <> SYNTAX_MACRO _
    Then
        Exit Function
    End If
    
    'Check for special names with two character prefixes, _
     e.g. "::section", "$.RAM" and "$#RAM"
    Dim Begin As Long
    'TODO: Watch out for word-length! Could be a 1-letter label
    'TODO: Could change this function to use bluString's CharPull
    If (WordChar0 = SYNTAX_LABEL And Word.Char(1) = SYNTAX_LABEL) _
    Or (WordChar0 = SYNTAX_RAM And Word.Char(1) = SYNTAX_PROPERTY) _
    Or (WordChar0 = SYNTAX_RAM And Word.Char(1) = SYNTAX_HASH) _
    Then
        Let Begin = 2
    Else
        Let Begin = 1
    End If
    
    'The name cannot begin with a number or a dot
    If Word.Char(Begin) = SYNTAX_PROPERTY Then Exit Function
    If Word.CharIsNumeric(Begin, True) Then Exit Function
    
    'Cannot end in a dot
    Dim Last As Long
    Let Last = Word.Length - 1
    If Word.Char(Last) = SYNTAX_PROPERTY Then Exit Function
    
    Dim i As Long
    For i = Begin To Last
        'Not a dot?
        If Word.Char(i) <> SYNTAX_PROPERTY Then
            'Is letter not alphanumeric?
            If Not Word.CharIsAlphaNumeric(i, True) Then
                'Finally, not an underscore?
                If Word.Char(i) <> ASC_USC Then Exit Function
            End If
        Else
            'If a dot, ensure that the next letter is not a number. _
             (the check for the last letter as a dot has already been done, _
              so we can safely assume there remains at least one letter)
            If Word.CharIsNumeric(i + 1, True) Then Exit Function
            'Two dots in a row not allowed
            If Word.Char(i + 1) = SYNTAX_PROPERTY Then Exit Function
        End If
    Next i
    
    Let IsValidName = True
End Function

'NestIn : Track an additional level of bracket nesting
'======================================================================================
'BracketType    | Token number to represent "{", "(" or "["
'---------------+----------------------------------------------------------------------
'Returns        | An error number, should the nesting be invalid
'======================================================================================
Private Function NestIn( _
    ByRef BracketType As OZ80_TOKEN _
) As OZ80_ERROR
    'Open another nesting level
    'TODO: Error above 255
    Let NestingDepth = NestingDepth + 1
    Let Nesting(NestingDepth).Token = BracketType
    Let Nesting(NestingDepth).Index = TokenIndex
End Function

'NestOut : Finish a level of bracket nesting, folding up to the previous level
'======================================================================================
'BracketOpen    | Token number for the opening bracket "{", "(" or "["
'BracketClose   | Token number for the closing bracket "}", ")" or "]"
'---------------+----------------------------------------------------------------------
'Returns        | An error number, should the nesting be invalid
'======================================================================================
Private Function NestOut( _
    ByRef BracketOpen As OZ80_TOKEN, _
    ByRef BracketClose As OZ80_TOKEN _
) As OZ80_ERROR
    With Nesting(NestingDepth)
        'Check that the nesting is correct; you cannot close a bracket without an _
         open one before, nor can you close a bracket when a different one is open
        If .Token <> BracketOpen Then
            Let NestOut = RaiseError(OZ80_ERROR_EXPECTED_BRACKET)
            Exit Function
        End If
        'Update the opening brace to reference the closing brace
        Let Files(File).Tokens(.Index).Value = TokenIndex
    End With
    'Reduce the nesting level
    Let NestingDepth = NestingDepth - 1
End Function

'/// Z80 //////////////////////////////////////////////////////////////////////////////

'This is a structured map of the OZ80 syntax: _
 The first column on the left is the name of a scope. For each scope, the columns to _
 the right represent the heirarchy of allowable items in that scope. Whatever is in _
 the third column is only allowed to follow what is in the second column. Items below _
 one another within a column represent more than one item being allowed within a scope.
'Another scope name within a column represents a nested scope. Items within brackets _
 show optional items

'NOTES: *   You can take `{text}` to mean any UCS-2 unicode point _
        *   `{operator}` can be any one of `+ - * / ^ \ & | << >>`

'Scope    : 1           2           3           4         ' Comment
'---------|-----------------------------------------------'----------------------------
'<value>  | {number}                                      '
'         | !{name}                                       ' "Constant"
'         | :{name}                                       ' "Label"
'         | .{name}                                       ' "Property"
'         | #{name}                                       ' "Object"
'         | $.{name}                                      ' "RAM Name"
'         | ::{name}                                      ' "Section"
'         | HI          <expr>                            ' High-8-bits of <expr>
'         | LO          <expr>                            ' Low-8-bits of <expr>
'         |                                               '
'<expr>   | <value>    [{operator}  <expr>]               '
'         | {           <expr>      }                     '
'         |                                               '
'<text>   | "           {text}      "                     '
'         |                                               '
'<list>   | <expr>     [,]         [<list>]               '
'         | <text>     [,]         [<list>]               '
'         |                                               '
'<params> | <param>    [,]         [<param>]              '
'         |                                               '
'<param>  | {register} [HELP        <text>]               '
'         | $.{name}   [HELP        <text>]               '
'         | {number}   [HELP        <text>]               '
'         |                                               '
'.........|...............................................'............................
'         |                                               '
'<root>   | DEF         <def>                             ' Define ...
'         |                                               '
'         | ENUM        !{name}    [<expr>]     <hash>    ' Enumerate Constants
'         |                                               '
'         | <if>                                          '
'         |                                               '
'         | INCLUDE     <text>                            '
'         |                                               '
'         | HASH        #{name}     <hash>                '
'         |                                               '
'         | PROC        :{name}    [HELP        <text>]   '
'         |                        [PARAMS      <params>] '
'         |                        [RETURN      <params>] '
'         |                        [INTERRUPT   <expr>]   '
'         |                         <chunk>               '
'         |                                               '
'         | TABLE       :{name}    [HELP        <text>]   '
'         |                         <chunk>               '
'         |                                               '
'<def>    | !{name}     [<expr>]                          ' Define Constant
'         | SECTION     ::{name}   [SLOT        <list>]   ' Define Section
'         | SYSTEM      <text>                            ' Define System
'         | RAM         $#{name}    <expr>      <expr>    ' Define RAM Bank
'         |                                               '
'<chunk>  | {           <code>      }                     '
'         |                                               '
'<code>   | .{name}                                       '
'         | BYTE        <list>                            '
'         | WORD        <list>                            '
'         | TRIP        <list>                            '
'         | LONG        <list>                            '
'         | <if>                                          '
'         | <asm>                                         '
'         |                                               '
'<if>     | IF          <expr>      <chunk>               '
'         |                        [ELSE        <if>]     '

'ScopeRoot : When we're not within the scope of anything else
'======================================================================================
Private Sub ScopeRoot()
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
Continue:
    'Define Constant / RAM Bank / Section / System:
    '----------------------------------------------------------------------------------
    If TokenKind = TOKEN_KEYWORD_DEF Then
        'Pass over the `DEF` keyword
        'ERROR: The file ended before the first parameter
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'What kind of definition is this?
        If TokenKind = TOKEN_CONST Then
            'Defining a constant
            Call DefineConst
        
        'Defining a RAM Bank
        ElseIf TokenKind = TOKEN_KEYWORD_RAM Then
            'Hand off to the function that will process this
            Call DefineRAMBank
            
        'Defining the target System
        ElseIf TokenKind = TOKEN_KEYWORD_SYSTEM Then
            'Hand off to the function that will process this
            Call DefineSystem
        
        'Defining a code Section
        ElseIf TokenKind = TOKEN_KEYWORD_SECTION Then
            'Hand off to the function that will process this
            Call DefineSection
            
        Else
            'ERROR: Unexpected definition
            'TODO: More specific error number/description
            Call RaiseError(OZ80_ERROR_EXPECTED_ROOT)
        End If
    
    'Define a Hash:
    '----------------------------------------------------------------------------------
    '`HASH #<hash> [ ... ]`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_HASH Then
        'Hand off to the function that will process this
        Call ProcessHash
    
    'Inlcude another source file:
    '----------------------------------------------------------------------------------
    '`INCLUDE <text>`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_INCLUDE Then
        'Hand off to the function that will process this
        Call ProcessInclude
    
    'Define a Procedure:
    '----------------------------------------------------------------------------------
    '`PROC :<label> { [...] }`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_PROC Then
        'Hand off to the function that will process this
        Call DefineProc
    
    'Create some RAM variables:
    '----------------------------------------------------------------------------------
    '`RAM [<expr>] [ ... ]`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_RAM Then
        'Hand off to the function that will process this
        Call DefineRAMHash
    
    'Define the starting Procedure:
    '----------------------------------------------------------------------------------
    '`START {...}`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_START Then
        'Hand off to the function that will process this
        Call DefineStart
        
    'Defing a data table:
    '----------------------------------------------------------------------------------
    '`TABLE :<name> {...}`
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_TABLE Then
        'Hand off to the function that will process this
        Call DefineTable
    
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
        Call RaiseError(OZ80_ERROR_EXPECTED_ROOT)
    End If
    
    If My_Error Then Exit Sub
    'As we loop, other functions will move the Token pointer forward (such as when _
     evaluating expressions). Check if we've hit the end of the file
    If EOF Then Exit Sub
    'As long as there is content left to process in the file, keep going
    GoTo Continue
End Sub

'DefineConst : Define a Constant value
'======================================================================================
'DEF        !<name>     [<expr>]
'======================================================================================
Private Sub DefineConst()
    '`ScopeRoot` has already verified that the current token is a Constant name; _
     check if the Constant has already been defined
    If Values.Exists(TokenText) Then
        'ERROR: You can't redefine Constants
        Call RaiseError(OZ80_ERROR_DUPLICATE_CONSTANT)
        Exit Sub
    End If
    
    'Remember the Constant's name before moving past it
    Dim ConstName As String
    Let ConstName = TokenText
    
    'The constant value is optional (it will default to 0), so the file can end here _
     (it's up to the parent to check if there's a unclosed scope problem)
    If Not NextToken() Then
        'Define the Constant as zero
        Call Values.Add(ConstName, 0)
        Exit Sub
    End If
    
    'Check that this is an Expression
    If Not TokenIsExpression(TokenKind) Then
        'ERROR: Expected an Expression to define the Constant
        Call RaiseError(OZ80_ERROR_EXPECTED_EXPRESSION)
        Exit Sub
    End If
    
    'Calculate the Expression's Value:
    '(If there was an error in the Expression, the error event _
      will already have been raised, we only need to fold upwards)
    Dim Value As Double
    If ParseExpression(Value) Then Exit Sub
    
    'Define the Constant:
    Call Values.Add(ConstName, Value)
End Sub

'ProcessHash : Process a Hash block
'======================================================================================
'HASH       #<name>     [ ... ]
'======================================================================================
Private Sub ProcessHash()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    Call NextToken
    
    If TokenKind = TOKEN_SQUARE_OPEN Then
        Let TokenIndex = TokenValue
        Call NextToken
    End If
End Sub

'ProcessInclude : Process an Include statement
'======================================================================================
'INCLUDE    <text>
'======================================================================================
Private Sub ProcessInclude()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Fetch the file path _
     (if there's an error doing this, fold upwards)
    Dim FilePath As bluString
    If GetText(FilePath) Then Exit Sub
    
    'Remember the current file being parsed
    Dim ReturnFile As Long, ReturnToken As Long
    Let ReturnFile = File
    Let ReturnToken = TokenIndex
    
    'Tokenise & assemble the included file
    If File_Assemble(FilePath.Text) Then Exit Sub
    
    'Reattach the previous file
    Let File = ReturnFile
    Let TokenIndex = ReturnToken
    Call CacheToken
    
    Call RaiseMessage( _
        OZ80_LOG_ACTION, "Continuing Assembly of ""{$}""", _
        Files(File).Path _
    )
End Sub

'DefineProc : Process a Procedure Chunk
'======================================================================================
'PROC   :<label>
'       [HELP       <text>]
'       [SECTION    ::<section>]    |   [INTERRUPT  <expr>]
'       [PARAMS     <params>]
'       [RETURN     <params>]
'{ [...] }
'======================================================================================
Private Sub DefineProc()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'ERROR: No Label name for Procedure
    If TokenKind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_PROC_NAME _
    ): GoTo ErrOut
    
    'Remember the String of the Label name, _
     we can't create the Label until the parameters are read
    Dim LabelName As String
    Let LabelName = TokenText
    
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'The `INTERRUPT` parameter is optional, but can be zero, so use 'undefined'
    Dim Interrupt As Double
    Let Interrupt = INDEFINITE
    
ReadParams:
    'Keep track of which parameters have been included or not
    Dim IsInterruptDefined As Boolean
    Dim IsSectionDefined As Boolean
    Dim IsParamsDefined As Boolean
    Dim IsReturnDefined As Boolean
    Dim IsHelpDefined As Boolean
    
    '----------------------------------------------------------------------------------
    '`SECTION` Parameter : Set which Section the Procedure belongs to _
     ----------------------------------------------------------------------------------
    If TokenKind = TOKEN_KEYWORD_SECTION Then
        'ERROR: The Section parameter has already been specified
        If IsSectionDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_SECTION _
        ): GoTo ErrOut
        
        'ERROR: Cannot co-exist with the `INTERRUPT` parameter
        If IsInterruptDefined Then Call RaiseError(OZ80_ERROR_INVALID_PROC_INTERRUPT): GoTo ErrOut
        
        'ERROR: File ended before the Section name?
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'ERROR: Not a Section name following the Section keyword
        If TokenKind <> TOKEN_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_SECTION_NAME _
        ): GoTo ErrOut
        
        'ERROR: Section has not been defined
        If Not Labels.Exists(TokenText) Then Call RaiseError( _
            OZ80_ERROR_INVALID_SECTION _
        ): GoTo ErrOut
        
        'Record the Section name
        Dim SectionName As String
        Let SectionName = TokenText
        Let IsSectionDefined = True
        
        'Pass over the Section name
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        GoTo ReadParams
    
    '----------------------------------------------------------------------------------
    '`HELP` Parameter : Provide a description for the procedure
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_HELP Then
        'ERROR: The `HELP` parameter has already been specified
        If IsParamsDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_HELP _
        ): GoTo ErrOut
         
        'Pass over the `HELP` keyword, the file cannot end here
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Fetch the text description (ignored for now)
        'For now, just read and discard only text tokens
        Do
            'If not a text token, then continue
            If TokenKind <> TOKEN_TEXT Then Exit Do
            'Move over the token, the file ending is an error
            If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        Loop
        
        Let IsHelpDefined = True
        GoTo ReadParams
        
    '----------------------------------------------------------------------------------
    '`PARAMS` Parameter : Document which Z80 parameters the Procedure uses _
     ----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_PARAMS Then
        'The `PARAMS` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM Names, _
         and optional `HELP` descriptions for each
         
        'ERROR: The `PARAMS` parameter has already been specified
        If IsParamsDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_PARAMS _
        ): GoTo ErrOut
         
        'Pass over the `PARAMS` keyword, the file cannot end here
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        Do
            'Check for a valid `PARAMS` item
            If TokenIsRegister(TokenKind) _
            Or TokenKind = TOKEN_RAM _
            Then
                'Skip over the list item, these are not recorded just yet _
                 but will be in the future (for auto-documentation)
                'ERROR: File cannot end in the midst of the List
                If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Check for a `HELP` description for the list item
                If TokenKind = TOKEN_KEYWORD_HELP Then
                    'Pass over the `HELP` keyword, the file cannot end here
                    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    'Fetch the text description (ignored for now)
                    'For now, just read and discard only text tokens
                    Do
                        'If not a text token, then continue
                        If TokenKind <> TOKEN_TEXT Then Exit Do
                        'Move over the token, the file ending is an error
                        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    Loop
                End If
                
            'The list ends when we hit one of the other `PROC` parameters or the _
             opening brace. Anything else is considered a bad `PARAMS` list item
            ElseIf TokenKind = TOKEN_KEYWORD_HELP _
            Or TokenKind = TOKEN_KEYWORD_RETURN _
            Or TokenKind = TOKEN_KEYWORD_SECTION _
            Or TokenKind = TOKEN_BRACES_OPEN _
            Then
                'The list has come to the end, _
                 look for more parameters
                Let IsParamsDefined = True
                GoTo ReadParams
            Else
                'ERROR: Only Registers and RAM Names are allowed at the moment
                Call RaiseError(OZ80_ERROR_EXPECTED_PROC_PARAMS)
                GoTo ErrOut
            End If
        Loop
        
    '----------------------------------------------------------------------------------
    '`RETURN` Parameter : Document what the Procedure sets / returns _
     ----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_RETURN Then
        'The `RETURN` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM Names
         
        'ERROR: The `RETURN` parameter has already been specified
        If IsReturnDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_PARAMS _
        ): GoTo ErrOut
         
        'Pass over the `RETURN` keyword, the file cannot end here
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        Do
            'Check for a valid `RETURN` item
            If TokenIsRegister(TokenKind) _
            Or TokenKind = TOKEN_RAM _
            Then
                'Skip over the list item, these are not recorded just yet _
                 but will be in the future (for auto-documentation)
                'ERROR: File cannot end in the midst of the List
                If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Check for a `HELP` description for the list item
                If TokenKind = TOKEN_KEYWORD_HELP Then
                    'Pass over the `HELP` keyword, the file cannot end here
                    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    'Fetch the text description (ignored for now)
                    'For now, just read and discard only text tokens
                    Do
                        'If not a text token, then continue
                        If TokenKind <> TOKEN_TEXT Then Exit Do
                        'Move over the token, the file ending is an error
                        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                    Loop
                End If
                
            'The list ends when we hit one of the other `PROC` parameters or the _
             opening brace. Anything else is considered a bad `RETURN` list item
            ElseIf TokenKind = TOKEN_KEYWORD_HELP _
            Or TokenKind = TOKEN_KEYWORD_PARAMS _
            Or TokenKind = TOKEN_KEYWORD_SECTION _
            Or TokenKind = TOKEN_BRACES_OPEN _
            Then
                'The list has come to the end, _
                 look for more parameters
                Let IsReturnDefined = True
                GoTo ReadParams
            Else
                'ERROR: Only Registers and RAM Names are allowed at the moment
                Call RaiseError(OZ80_ERROR_EXPECTED_PROC_RETURN)
                GoTo ErrOut
            End If
        Loop
        
    '----------------------------------------------------------------------------------
    '`INTERRUPT` Parameter : Define a fixed-location Interrupt Procedure _
     ----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_INTERRUPT Then
        'The `INTERRUPT` parameter fixes a Procedure to a specific location in the _
         ROM intended for hardware interrupts. This parameter cannot co-exist with _
         the `SECTION` parameter
        
        'ERROR: The `INTERRUPT` parameter has already been specified
        If IsInterruptDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_INTERRUPT _
        ): GoTo ErrOut
        
        'ERROR: Cannot co-exist with the `SECTION` parameter
        If IsSectionDefined _
            Then Call RaiseError(OZ80_ERROR_INVALID_PROC_INTERRUPT): GoTo ErrOut
        
        'ERROR: File cannot end before the parameter Expression!
        If Not NextToken() _
            Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Is this an Expression at all?
        If Not TokenIsExpression(TokenKind) _
            Then Call RaiseError(OZ80_ERROR_EXPECTED): GoTo ErrOut
        
        'Calculate the Expression:
        '(If there was an error in the Expression, the error event _
          will already have been raised, we only need to fold upwards)
        If ParseExpression(Interrupt) Then GoTo ErrOut
        
        'An indefinite Value cannot be used for an Interrupt address, we will only be _
         able to define the indefinite Values (i.e. Label locations in ROM) once _
         the whole code is parsed and we use the Sections to layout the ROM
        If Interrupt = INDEFINITE Then Call RaiseError(OZ80_ERROR_INDEFINITE): GoTo ErrOut
        
        'Validate the Value, only certain Interrupt addresses are allowed
        If Not ((Interrupt = &H0) Or (Interrupt = &H8) Or (Interrupt = &H10) _
            Or (Interrupt = &H18) Or (Interrupt = &H20) Or (Interrupt = &H28) _
            Or (Interrupt = &H30) Or (Interrupt = &H38) Or (Interrupt = &H66)) _
        Then Call RaiseError(OZ80_ERROR_INVALID_INTERRUPT): GoTo ErrOut
        
        Let IsInterruptDefined = True
        GoTo ReadParams
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the Chunk
    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): GoTo ErrOut
    
    'The chunk cannot be left open
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'When we assemble the Procedure, we need to be able to normalise _
     the Property names to their Section+Label
    Let CurrentSection = SectionName
    Let CurrentLabel = LabelName
    
    'Now define the Label name; if a Section was given, it needs to be prepended
    Dim FullLabelName As String
    Let FullLabelName = CurrentSection & CurrentLabel
    
    'You can't name two Chunks the same thing
    If ChunkExists(FullLabelName) _
        Then Call RaiseError(OZ80_ERROR_DUPLICATE_LABEL): GoTo ErrOut
    
    'Create the Chunk
    If Interrupt = INDEFINITE Then
        'A standard Procedure without fixed (interrupt) address
        Call Chunk_Begin(FullLabelName)
    Else
        'A special fixed-Procedure for hardware interrupts
        Call Chunk_Begin(FullLabelName, Interrupt)
        'We keep direct references to Interrupt Procedures to make it easier to _
         lay out the ROM contents
        If Interrupt = &H0& Then
            Let Chunk_Interrupt00 = Chunk
        ElseIf Interrupt = &H8& Then Let Chunk_Interrupt08 = Chunk
        ElseIf Interrupt = &H10& Then Let Chunk_Interrupt10 = Chunk
        ElseIf Interrupt = &H18& Then Let Chunk_Interrupt18 = Chunk
        ElseIf Interrupt = &H20& Then Let Chunk_Interrupt20 = Chunk
        ElseIf Interrupt = &H28& Then Let Chunk_Interrupt28 = Chunk
        ElseIf Interrupt = &H30& Then Let Chunk_Interrupt30 = Chunk
        ElseIf Interrupt = &H38& Then Let Chunk_Interrupt38 = Chunk
        ElseIf Interrupt = &H66& Then Let Chunk_Interrupt66 = Chunk
        End If
    End If
    
'    Dim LogMsg As String
'    Let LogMsg = "PROC " & LabelName
'    If Interrupt <> INDEFINITE _
'        Then Let LogMsg = LogMsg & " INTERRUPT $" & HexStr8(CLng(Interrupt))
'    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    
    'Log a header for the disassembly
    Call Chunks(Chunk).LogASM.LogFormat("PROC {$}", LabelName)
    Call Chunks(Chunk).LogASM.Log(LOG_LINEDOUBLE)
    Let ChunkMode = MODE_UNSET
    
    'Enter the Procedure Chunk to process its innards
    Call ScopeChunk
    
    'Finalise the disassembly log
    If ChunkMode = MODE_DATA Then
        Call Chunks(Chunk).LogASM.LogFormat( _
            LOG_DATA, _
            ChunkDataAddr, ChunkData, _
            "·", "·", "·", "·", "·", "·", "·", "·", _
            "·", "·", "·", "·", "·", "·", "·", "·" _
        )
    End If
    Call Chunks(Chunk).LogASM.Log(LOG_LINEDOUBLE)
    Call RaiseMessage(OZ80_LOG_DEBUG, Chunks(Chunk).LogASM.Concatenate())
    
ErrOut:
    'Leave this scope
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'DefineRAMBank : Define a RAM Block
'======================================================================================
'DEF RAM    $#{name}    <expr>      <expr>
'======================================================================================
Private Sub DefineRAMBank()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'TODO: Implementation
    Call NextToken
    Call NextToken
    Call NextToken
End Sub

'DefineRAMHash : Process a RAM block
'======================================================================================
'RAM       [<expr>]     [ ... ]
'======================================================================================
Private Sub DefineRAMHash()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'TODO: Implementation
    Call NextToken
    If TokenKind = TOKEN_SQUARE_OPEN Then
        Let TokenIndex = TokenValue
        Call NextToken
    End If
End Sub

'DefineSection : Define a Section
'======================================================================================
'DEF SECTION    ::<section>    [SLOT   <list>]
'======================================================================================
Private Sub DefineSection()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'ERROR: Not a section name following the section keyword
    If TokenKind <> TOKEN_SECTION Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_SECTION_NAME _
    ): Exit Sub
    
    'Keep ahold of the Section identifier
    Dim SectionName As String
    Let SectionName = TokenText
    
    'Has this Section already been defined?
    If Labels.Exists(TokenText) Then Call RaiseError( _
        OZ80_ERROR_DUPLICATE_SECTION _
    ): Exit Sub
    
    Call Labels.Add(SectionName, INDEFINITE)
    Call SectionSlots.Add(SectionName, 0)
    
    '----------------------------------------------------------------------------------
    
    'If the file ends here it's valid, as the params are optional
    If Not NextToken() Then GoTo Finish
    
    'The only allowed Slot lists are: 0 / 1 / 2 / 0, 1 / 0, 1, 2 / 1, 2. _
     These are stored as a bit-pattern allowing packing into one value, _
     which when zero infers the default "let the assembler choose for me"
    Dim Slot As OZ80_SLOT
    
    'The only parameter is the Slot List
    If TokenKind = TOKEN_KEYWORD_SLOT Then
        'ERROR: File ending before we have the value
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'ERROR: There needs to be at least one item in the List
        If Not TokenIsExpression(TokenKind) _
            Then Call RaiseError(OZ80_ERROR_EXPECTED): Exit Sub
        
        'Get the list of Slots assigned to the Section
        Dim List As Scripting.Dictionary
        Set List = GetListAsDictionary()
        
        'Can't be more than three -- "0, 1, 2"
        If List.Count > 3 Then Call RaiseError(OZ80_ERROR_INVALID_SLOT): Exit Sub
        
        'An indefinite value cannot be used for the Slot
        If List.Item(0) = INDEFINITE Then GoTo ErrIndefinite
        
        'Check for one item: 0, 1 or 2
        If List.Count = 1 Then
            If List.Item(0) = 0 Then Let Slot = SLOT0: GoTo Finish
            If List.Item(0) = 1 Then Let Slot = SLOT1: GoTo Finish
            If List.Item(0) = 2 Then Let Slot = SLOT2: GoTo Finish
            'An invalid value
            GoTo ErrInvalid
        End If
        
        'Check for two items: 0, 1 or 1, 2
        If List.Count = 2 Then
            'An indefinite value cannot be used for the Slot
            If List.Item(1) = INDEFINITE Then GoTo ErrIndefinite
            If List.Item(0) = 0 And List.Item(1) = 1 Then Let Slot = SLOT0 Or SLOT1: GoTo Finish
            If List.Item(0) = 1 And List.Item(1) = 2 Then Let Slot = SLOT1 Or SLOT2: GoTo Finish
            'An invalid value
            GoTo ErrInvalid
        End If
        
        'Check for 0, 1, 2
        If List.Item(0) <> 0 Then GoTo ErrInvalid
        If List.Item(1) <> 1 Then GoTo ErrInvalid
        If List.Item(2) <> 2 Then GoTo ErrInvalid
        Let Slot = SLOT0 Or SLOT1 Or SLOT2
        
    End If

Finish:
    'Record the result:
    Let SectionSlots(SectionName) = Slot
    Call RaiseMessage( _
        OZ80_LOG_STATUS, LOG_SECTION, _
        SectionName _
    )
'    RaiseEvent Message(OZ80_LOG_STATUS, _
'        "SECTION " & TokenStrings(SectionKey) & _
'        " SLOT " & IIf( _
'            Slot = 0, "?", _
'                IIf(Slot And OZ80_SLOT.SLOT0, "0 ", "") & _
'                IIf(Slot And OZ80_SLOT.SLOT1, "1 ", "") & _
'                IIf(Slot And OZ80_SLOT.SLOT2, "2 ", "") _
'            ) _
'    )
    Exit Sub

    '----------------------------------------------------------------------------------
ErrInvalid:
    'Only a specific Slot pattern can be used
    Call RaiseError(OZ80_ERROR_INVALID_SLOT)
    Exit Sub

ErrIndefinite:
    'An indefinite Value cannot be used for a Slot number, we will only be _
     able to define the indefinite Values (i.e. Label locations in ROM) once _
     the whole code is parsed and we use the Sections to layout the ROM
    Call RaiseError(OZ80_ERROR_INDEFINITE)
End Sub

'DefineStart : Define the System's boot Procedure
'======================================================================================
'START      { ... }
'======================================================================================
Private Sub DefineStart()
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'The only allowed token is the opening of the Chunk
    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): Exit Sub
    
    'The chunk cannot be left open
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'You can't define the START procedure twice
    If ChunkExists(STR_START) Then _
        Call RaiseError(OZ80_ERROR_DUPLICATE_START): Exit Sub
    
    'Create the Chunk
    'TODO: Check that the System has been defined
    Call Chunk_Begin(STR_START, 0)
    
    'Keep a reference to this for laying out the ROM
    Let Chunk_Start = Chunk
    
    Let CurrentSection = vbNullString
    Let CurrentLabel = STR_START
    
    'Enter the Procedure Chunk to process its innards
    Call ScopeChunk
    
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'DefineSystem : Define the target system (e.g. SEGA Master System)
'======================================================================================
'SYSTEM     <text>
'======================================================================================
Private Sub DefineSystem()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Check that the parameter is a list (or at least the start of one)
    If Not TokenIsList(TokenKind) _
        Then Call RaiseError(OZ80_ERROR_EXPECTED_SYSTEM_NAME): Exit Sub
    
    'Fetch the System name string from the proceeding list _
     (if there was any error during that, fold upwards)
    Dim SystemName As bluString
    If GetText(SystemName) Then Exit Sub
    
    'Check if it's any of the allowed System names
    If SystemName.Equals(STR_SMS) Then
        'Configure the assembler for the SEGA Master System
        Let My_System = SYSTEM_SMS
    Else
        'Not a valid system name
        Call RaiseError(OZ80_ERROR_EXPECTED_SYSTEM_NAME)
    End If
End Sub

'DefineTable : Process a data table Chunk
'======================================================================================
'TABLE  :<label>
'       [HELP       <text>]
'       [SECTION    ::<section>]
'{ [...] }
'======================================================================================
Private Sub DefineTable()
    'ERROR: The file ended before the first parameter
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'ERROR: No Label name for Table
    If TokenKind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_TABLE_NAME _
    ): GoTo ErrOut
    
    'Remember the String of the Label name, _
     we can't create the Label until the parameters are read
    Dim LabelName As String
    Let LabelName = TokenText
    
    'ERROR: The file cannot end until the Chunk has been opened and closed
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
ReadParams:
    'Keep track of which params have been included or not
    Dim IsSectionDefined As Boolean
    
    '----------------------------------------------------------------------------------
    '`SECTION` Parameter : Set which Section the Table belongs to _
     ----------------------------------------------------------------------------------
    If TokenKind = TOKEN_KEYWORD_SECTION Then
        'ERROR: The Section parameter has already been specified
        If IsSectionDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_TABLE_SECTION _
        ): GoTo ErrOut
        
        'ERROR: File ended before the Section name?
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'ERROR: Not a Section name following the Section keyword
        If TokenKind <> TOKEN_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_SECTION_NAME _
        ): GoTo ErrOut
        
        'ERROR: Section has not been defined
        If Not Labels.Exists(TokenText) Then Call RaiseError( _
            OZ80_ERROR_INVALID_SECTION _
        ): GoTo ErrOut
        
        'Record the Section name
        Dim SectionName As String
        Let SectionName = TokenText
        Let IsSectionDefined = True
        
        'Pass over the Section name
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        GoTo ReadParams
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the Chunk
    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): GoTo ErrOut
    
    'The chunk cannot be left open
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
    
    'When we assemble the Table, we need to be able to normalise _
     the Property names to their Section+Label
    Let CurrentSection = SectionName
    Let CurrentLabel = LabelName
    
    'Now define the Label name; if a Section was given, it needs to be prepended
    Dim FullLabelName As String
    Let FullLabelName = CurrentSection & CurrentLabel
    
    'You can't name two Chunks the same thing
    If ChunkExists(FullLabelName) Then _
        Call RaiseError(OZ80_ERROR_DUPLICATE_LABEL): GoTo ErrOut
    
    'Create the Chunk
    Call Chunk_Begin(FullLabelName)
    
'    Dim LogMsg As String
'    Let LogMsg = "PROC " & LabelName
'    If Interrupt <> INDEFINITE _
'        Then Let LogMsg = LogMsg & " INTERRUPT $" & HexStr8(CLng(Interrupt))
'    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    
    'Log a header for the disassembly
    Call Chunks(Chunk).LogASM.LogFormat("TABLE {$}", LabelName)
    Call Chunks(Chunk).LogASM.Log(LOG_LINEDOUBLE)
    Let ChunkMode = MODE_UNSET
    
    'Enter the Table Chunk to process its innards
    Call ScopeChunk
    
    'Finalise the disassembly log
    If ChunkMode = MODE_DATA Then
        Call Chunks(Chunk).LogASM.LogFormat( _
            LOG_DATA, _
            ChunkDataAddr, ChunkData, _
            "·", "·", "·", "·", "·", "·", "·", "·", _
            "·", "·", "·", "·", "·", "·", "·", "·" _
        )
    End If
    Call Chunks(Chunk).LogASM.Log(LOG_LINEDOUBLE)
    Call RaiseMessage(OZ80_LOG_DEBUG, Chunks(Chunk).LogASM.Concatenate())
    
ErrOut:
    'Leave this scope
    Let CurrentSection = vbNullString
    Let CurrentLabel = vbNullString
End Sub

'ScopeChunk: Within a code/data Chunk, i.e. `PROC`, `TABLE`
'======================================================================================
Private Sub ScopeChunk()
    Dim Value As Double
    'Within a code/data chunk we can have local labels, data statements and Z80 code
    
Continue:
    'Z80 code?
    If TokenIsInstruction(TokenKind) Then
        '*** TIMER BEGIN ***
        Call Profiler.EventBegin(EVENT_PROCESSZ80)
        Call Z80_Process
        Call Profiler.EventEnd(EVENT_PROCESSZ80)
        '*** TIMER END ***
        
        If My_Error Then Exit Sub
        GoTo Continue
    
    '----------------------------------------------------------------------------------
    'Byte Data
    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_BYTE Then
'        'Move past the `BYTE` keyword
'        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
'
'        'There MUST be at least one list item after the keyword
'        'TODO: Add error here
'        If Not TokenIsList(TokenKind) Then Stop
'
'        'Start walking the list
'        Do
'            If TokenIsExpression(TokenKind) Then
'                If ParseExpression(Value) Then Stop
'
'                If ChunkMode <> MODE_DATA Then
'                    Dim i As Long
'                    For i = 0 To 15
'                        Let ChunkData(i) = HexStr8(-1)
'                    Next
'                    Let ChunkDataIndex = Chunk.Size Mod 16
'                    Let ChunkDataAddr = HexStr16(Chunk.Size)
'
'                    If ChunkMode <> MODE_UNSET Then Call Chunk.LogASM.Add( _
'                        "-----+-------------------+-----------------------------+-----------------" _
'                    )
'                    Call Chunk.LogASM.Add( _
'                        "ADDR | 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F | 0123456789ABCDEF" _
'                    )
'                    Call Chunk.LogASM.Add( _
'                        "-----|-------------------------------------------------|-----------------" _
'                    )
'                    Let ChunkMode = MODE_DATA
'                End If
'
'                If Value = INDEFINITE Then
'                    Call Chunk.AddByte(0)
'                    Let ChunkData(ChunkDataIndex) = HexStr8(-1)
'                Else
'                    Call Chunk.AddByte(CLng(Value))
'                    Let ChunkData(ChunkDataIndex) = HexStr8(CLng(Value))
'                End If
'
'                Let ChunkDataIndex = ChunkDataIndex + 1
'                If ChunkDataIndex = 16 Then
'                    Call Chunk.LogASM.Add( _
'                        LOG_DATA, _
'                        ChunkDataAddr, ChunkData, _
'                        "·", "·", "·", "·", "·", "·", "·", "·", _
'                        "·", "·", "·", "·", "·", "·", "·", "·" _
'                    )
'                    Erase ChunkData
'                    Let ChunkDataIndex = 0
'                End If
'            ElseIf Not TokenIsList(TokenKind) Then
'                Exit Do
'            End If
'
'            Call NextToken
'        Loop
    End If
    
    'The chunk might be empty `{}`, in which case we don't want to walk it
    If TokenKind = TOKEN_BRACES_CLOSE Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    
    'Move to the next token. If the file ends unexpectedly, the chunk has not been _
     closed properly (closing brace is expected)
    If Not NextToken() Then
        Call RaiseError(OZ80_ERROR_FILE_END)
    
    'When the closing brace is encountered, we'll drop out of the scope
    ElseIf TokenKind <> TOKEN_BRACES_CLOSE Then
        GoTo Continue
    End If
    
'    'Finalise the disassembly log
'    If ChunkMode = MODE_DATA Then
'        Call RaiseMessage( _
'            OZ80_LOG_DEBUG, LOG_DATA, _
'            ChunkDataAddr, ChunkData, _
'            "·", "·", "·", "·", "·", "·", "·", "·", _
'            "·", "·", "·", "·", "·", "·", "·", "·" _
'        )
'    End If
'    Call RaiseMessage(OZ80_LOG_DEBUG, LOG_LINEDOUBLE)
'    Let ChunkMode = MODE_UNSET
    
Finish:
    'Move onto the next token ready for the root scope to process. _
     The file could end just after the chunk, so no error is necessary
    Call NextToken
End Sub

'Z80_GetParameter : Fetch the type / value of a Z80 instruction parameter
'======================================================================================
'ReturnParam    | An `oz80_Param` structure to populate
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function Z80_GetParameter( _
    ByRef ReturnParam As oz80Param _
) As OZ80_ERROR
    Dim Value As Double
    
    'If the parameter doesn't contain a Value, default to -1. This is so that when _
     building the opcode we can immediately tell when to include the offset byte
    Let ReturnParam.Value = -1
    
    '----------------------------------------------------------------------------------
    'A memory reference, such as `[IX+$FF]` _
     ----------------------------------------------------------------------------------
    If TokenKind = TOKEN_SQUARE_OPEN Then
        'ERROR: File ends just as the memory expression is opened
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'The first item in the memory expression can be a register _
         ..............................................................................
        If TokenIsRegister(TokenKind) Then
            'Copy the register type to the parameter being built
            Let ReturnParam.Token = TokenKind
            
            'The `IN` & `OUT` instructions accept only the C register, _
             though in reality this is register pair BC
            If TokenKind = TOKEN_Z80_C Then
                Let ReturnParam.Mask = MASK_MEM_BC
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_BC
            
            '`LD` accepts registers BC & DE
            ElseIf TokenKind = TOKEN_Z80_BC Then
                Let ReturnParam.Mask = MASK_MEM_BC
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_BC
            ElseIf TokenKind = TOKEN_Z80_DE Then
                Let ReturnParam.Mask = MASK_MEM_DE
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_DE
            
            'HL / IX & IY
            ElseIf TokenKind = TOKEN_Z80_HL Then
                Let ReturnParam.Mask = MASK_MEM_HL
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_HL
                
            'IX/IY can have an additional expression, e.g. `ld A, [IX+$FF]`
            ElseIf TokenKind = TOKEN_Z80_IX Then
                Let ReturnParam.Mask = MASK_MEM_IX
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_IX
                GoTo AddExpr
            ElseIf TokenKind = TOKEN_Z80_IY Then
                Let ReturnParam.Mask = MASK_MEM_IY
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_IY
                GoTo AddExpr
                
            ElseIf TokenKind = TOKEN_Z80_SP Then
                Let ReturnParam.Mask = MASK_MEM_SP
                Let ReturnParam.LogTextPtr = PARAMTEXT_MEM_SP
            Else
                'ERROR: register not allowed here
                Stop
            End If
            
            'ERROR: File ends before the memory expression is closed
            If Not NextToken() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
            
            GoTo EndScope
            
AddExpr:    'Look for the `+` following the register name
            If Not NextToken() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
            
            'If present, skip over it and check for an expression
            If TokenKind = TOKEN_OPERATOR_ADD Then
                'ERROR: Something must follow the "+"!
                If Not NextToken() _
                    Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
                
                'Try and calculate the Expression
                '(If there was an error in the Expression, the error event _
                  will already have been raised, we only need to fold upwards)
                If ParseExpression(Value) Then GoTo ErrOut
                
                'TODO: Handle indefinite value
                'TODO: Error if number is over 16-bit
                'TODO: Issue warning for rounding
                
                If Value = INDEFINITE _
                    Then Let ReturnParam.Value = 0 _
                    Else Let ReturnParam.Value = Value
                
                'If this is a value for an IX/IY memory reference, _
                 select the appropriate cached text for logging
                'TODO: Validate maximum 8-bit? Where?
                'TODO: Special text for indefinite?
                
                If ReturnParam.Mask = MASK_MEM_IX Then
                    Let ReturnParam.LogTextPtr = _
                        PARAMTEXT_MEM_IX_VAL + ReturnParam.Value
                ElseIf ReturnParam.Mask = MASK_MEM_IY Then
                    Let ReturnParam.LogTextPtr = _
                        PARAMTEXT_MEM_IY_VAL + ReturnParam.Value
                Else
                    Let ReturnParam.LogTextPtr = _
                        PARAMTEXT_MEM_VAL16 + ReturnParam.Value
                End If
                
                If TokenKind = TOKEN_SQUARE_CLOSE Then
                    Call NextToken
                    GoTo ErrOut
                End If
                'ERROR: Unexpected token in memory expression
                Stop
            End If
        
        '..............................................................................
        Else
            'Is it an expression at all?
            'TODO: ERROR: Not a valid expression
            If Not TokenIsExpression(TokenKind) Then Stop
        
            'Besides registers, memory addresses are allowed
            Let ReturnParam.Token = TOKEN_NUMBER
            
            'Try and calculate the expression
            '(If there was an error in the Expression, the error event _
              will already have been raised, we only need to fold upwards)
            If ParseExpression(Value) Then GoTo ErrOut
            
            'TODO: Handle indefinite value
            'TODO: Error if number is over 16-bit
            'TODO: Issue warning for rounding
            
            If Value = INDEFINITE _
                Then Let ReturnParam.Value = 0 _
                Else Let ReturnParam.Value = Value
            
            'Any number will class as 16-bit since it's the maximum allowed
            Let ReturnParam.Mask = MASK_MEM_VAL
            Let ReturnParam.LogTextPtr = _
                PARAMTEXT_MEM_VAL16 + ReturnParam.Value
        End If
    
EndScope:
        '..............................................................................
        If TokenKind = TOKEN_SQUARE_CLOSE Then
            Call NextToken
            GoTo ErrOut
        End If
        'ERROR: Unexpected token in memory expression
        Stop
    
    '----------------------------------------------------------------------------------
    'A Register _
     ----------------------------------------------------------------------------------
    ElseIf TokenIsRegister(TokenKind) _
        Or TokenIsFlag(TokenKind) _
    Then
        'Copy the register type to the parameter being built
        Let ReturnParam.Token = TokenKind
        
        'Set the index for the cached text prepared for the register name
        Let ReturnParam.LogTextPtr = _
            PARAMTEXT_REGS + (TokenKind - [_TOKEN_REGS_BEGIN])
        
        'Map the register token to the mask bit; token numbers are sequential _
         whilst parameter masks are bit combinations allowing rapid comparisons _
         of parameters with allowed options
        Let ReturnParam.Mask = TokenRegister(TokenKind)
        
        'Now move beyond the register token
        Call NextToken
    
    '----------------------------------------------------------------------------------
    'Anything else classes as a calculable expression _
     ----------------------------------------------------------------------------------
    ElseIf TokenIsExpression(TokenKind) Then
        'We manually build a number token, since we don't want to use the first _
         token of the expression -- it might be a label for instance and we can _
         only send registers or numbers to be assembled into opcodes
        Let ReturnParam.Token = TOKEN_NUMBER
        Let ReturnParam.Mask = MASK_VAL

GetValue:
        'Try and calculate the expression
        '(If there was an error in the Expression, the error event _
          will already have been raised, we only need to fold upwards)
        If ParseExpression(Value) Then GoTo ErrOut
        
        'TODO: Handle indefinite value
        'TODO: Error if number is over 16-bit
        'TODO: Issue warning for rounding
        
        If Value = INDEFINITE _
            Then Let ReturnParam.Value = 0 _
            Else Let ReturnParam.Value = Value
    
        Let ReturnParam.LogTextPtr = PARAMTEXT_VAL16 + ReturnParam.Value
        
    '----------------------------------------------------------------------------------
    Else
        'ERROR: Not a Z80 parameter
        'TODO: parameter specific error number
        Call RaiseError(OZ80_ERROR_EXPRESSION_Z80)
    End If

ErrOut:
    'Return the error number
    Let Z80_GetParameter = My_Error
End Function

'Z80_Process : Process a Z80 instruction (and parameters)
'======================================================================================
'{asm} [<param>] [<param>] [<param>]
'======================================================================================
Private Sub Z80_Process()
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN
    Let Instruction = TokenKind
    'Skip over the Z80 instruction token now that we know it. _
     If the file ends it's an error because the chunk has not been closed
    If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    '----------------------------------------------------------------------------------
    
    'A built in macro allows `PUSH` and `POP` Z80 instructions to handle multiple _
     parameters; each parameter is broken into its own `PUSH`/`POP` instruction
    If (Instruction = TOKEN_Z80_PUSH) _
    Or (Instruction = TOKEN_Z80_POP) _
    Then
        'Is there any parameters to begin with?
        If Not TokenIsParameter(TokenKind) _
            Then Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS): Exit Sub
        
        'Clear any 2nd/3rd parameters set from the last time
        LSet Z80Param2 = Z80ParamEmpty
        LSet Z80Param3 = Z80ParamEmpty
        
        Do
            'Fetch a parameter
            '(If there was an error, the error number will already have been set, _
              we only need to fold upwards to end the assembling
            If Z80_GetParameter(Z80Param1) Then Exit Sub
            
            'If it's not a register, then it's invalid
            If (Z80Param1.Mask And MASK_REGS_AF_BC_DE_HL_IXY) = 0 Then
                Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
                Exit Sub
            End If
            
            'Assemble this as a single `PUSH`/`POP` instruction
            Call Z80_Assemble(Instruction)
            
        'Look for the next register in the list
        Loop While TokenIsRegister(TokenKind)
    
    'Up to three parameters are supported to allow for the undocumented instructions, _
     e.g. `SET 0 [IX+$8] A` which sets bit 0 of the address and copies the result to A
    '----------------------------------------------------------------------------------
    ElseIf TokenIsParameter(TokenKind) Then
        'Parameter Expressions might return an unknown Value _
         (such as when a Label, which has not yet been addressed, is used)
        
        'Get the first parameter
        '(If there was an error, the error number will already have been set, _
          we only need to fold upwards to end the assembling
        If Z80_GetParameter(Z80Param1) Then Exit Sub
        
        'Check if a second parameter exists
        If TokenIsParameter(TokenKind) Then
            'Fetch the second parameter
            If Z80_GetParameter(Z80Param2) Then Exit Sub
            
            'Check if a third parameter exists
            If TokenIsParameter(TokenKind) Then
                'Fetch the third parameter (very uncommon)
                'NOTE: Parameter three is only used for some undocumented instructions _
                 and will not be IX/IY (though IX/IY can be in previous parameters)
                If Z80_GetParameter(Z80Param3) Then Exit Sub
            Else
                'There is no third parameter, clear it from last time
                LSet Z80Param3 = Z80ParamEmpty
            End If
        Else
            'There are no second or third parameters, clear them from last time
            LSet Z80Param2 = Z80ParamEmpty
            LSet Z80Param3 = Z80ParamEmpty
        End If
        
        Call Z80_Assemble(Instruction)
    '----------------------------------------------------------------------------------
    Else
        'There are no parameters, clear them from the last time
        LSet Z80Param1 = Z80ParamEmpty
        LSet Z80Param2 = Z80ParamEmpty
        LSet Z80Param3 = Z80ParamEmpty
        
        Call Z80_Assemble(Instruction)
    End If
End Sub

'Z80_Assemble : Assemble bytes from a Z80 instruction and parameters
'======================================================================================
'Instruction    | Token number of the Z80 instruction
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function Z80_Assemble( _
    ByRef Instruction As OZ80_TOKEN _
) As OZ80_ERROR
    Dim LogStr As String
    Let LogStr = LOG_ASM
    Dim LogPtr As Long
    Let LogPtr = StrPtr(LogStr)
    
    'We already know the instruction, so we can add that to the disassembly now
    Call api_RtlMoveMemory( _
        LogPtr + 54, _
        StrPtr(ParamText(PARAMTEXT_Z80 + (Instruction - [_TOKEN_Z80_BEGIN]))), _
        20 _
    )
    'The parameters have also been prepared
    If Z80Param1.Mask <> 0 Then Call api_RtlMoveMemory( _
        LogPtr + 74, StrPtr(ParamText(Z80Param1.LogTextPtr)), 20 _
    )
    If Z80Param2.Mask <> 0 Then Call api_RtlMoveMemory( _
        LogPtr + 94, StrPtr(ParamText(Z80Param2.LogTextPtr)), 20 _
    )
    If Z80Param3.Mask <> 0 Then Call api_RtlMoveMemory( _
        LogPtr + 114, StrPtr(ParamText(Z80Param3.LogTextPtr)), 20 _
    )
    
    'Initialse variables to hold each piece of the assembly
    Dim Opcode As Long
    Dim Prefix As Long, Offset As Long, Value As Long
    Let Prefix = -1: Let Offset = -1: Let Value = -1
    
    Dim ByteCount As Long
    Dim Bytes(0 To 6) As Byte
    
    '----------------------------------------------------------------------------------
    
    'Is Param 1 an 8/16-bit parameter Value?
    If (Z80Param1.Mask And MASK_ANY_VAL) <> 0 Then
        Let Value = Z80Param1.Value
    'Is Param 2 an 8/16-bit parameter Value?
    ElseIf (Z80Param2.Mask And MASK_ANY_VAL) <> 0 Then
        Let Value = Z80Param2.Value
    End If
    
    'Use of IX / IY requires a prefix byte ($DD/$FD) before the opcode. _
     We need to record which parameter is the IX/IY reference, along with its value
    Dim ParamIXY As oz80Param
    
    'Is Param 1 an IX/IY parameter?
    If (Z80Param1.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Z80Param1
    'Is Param 2 an IX/IY parameter?
    ElseIf (Z80Param2.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Z80Param2
    End If
    
    'If there is any IX/Y parameter, setup the opcode prefix
    If ParamIXY.Mask <> 0 Then
        'Is it IX or IY?
        If (ParamIXY.Mask And MASK_ANY_IX) <> 0 _
            Then Let Prefix = &HDD& _
            Else Let Prefix = &HFD&
        'We can already add this to the bytes to output
        Let Bytes(ByteCount) = Prefix
        Let ByteCount = ByteCount + 1
        'And include it in the disassembly log
        Call api_RtlMoveMemory(LogPtr + 14, StrPtr(Cache_HexStr8(Prefix)), 4)
        'Does it include an offset byte? i.e. "[IX+$FF]"
        If (ParamIXY.Mask And MASK_MEM_IXY) <> 0 Then
            Let Offset = ParamIXY.Value
            'Include this in the disassembly log
            Call api_RtlMoveMemory(LogPtr + 26, StrPtr(Cache_HexStr8(Offset)), 4)
        End If
    End If
     
    'See, BASIC does have a good use!
    On (Instruction - [_TOKEN_Z80_BEGIN] + 1) GoTo _
        iADC, iADD, iAND, iBIT, iCALL, iCCF, iCP, iCPD, iCPDR, iCPI, iCPIR, iCPL, _
        iDAA, iDEC, iDI, iDJNZ, iEI, iEX, iEXX, iHALT, iIM, iIN, iINC, iIND, iINDR, _
        iINI, iINIR, iJP, iJR, iLD, iLDD, iLDDR, iLDI, iLDIR, iNEG, iNOP, iOR, iOUT, _
        iOUTD, iOTDR, iOUTI, iOTIR, iPOP, iPUSH, iRES, iRET, iRETI, iRETN, iRLA, iRL, _
        iRLCA, iRLC, iRLD, iRRA, iRR, iRRCA, iRRC, iRRD, iRST, iSBC, iSCF, iSET, _
        iSLA, iSRA, iSLL, iSRL, iSUB, iXOR
    
    'TODO: ERROR: completely invalid token number!!
    Stop
    
iADC:
    'Add with Carry _
     ..................................................................................
    '`ADC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y1 Or OpcodeZ_Register(Z80Param2.Mask)
    '`ADC A, $8`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y1 Or Z6
    '`ADC HL, BC|DE|HL|SP`
    ElseIf Z80_TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = X1 Or Q1 Or Z2 Or OpcodeP_RegisterPair(Z80Param2.Token)
        GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iADD:
    'Add: _
     ..................................................................................
    '`ADD HL, BC|DE|HL|SP`
    If Z80_TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD IX, BC|DE|IX|SP`
    ElseIf Z80_TestParams(MASK_REG_IX, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IX Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IX"), _
         we still need to validate the params to ensure HL & IX are not mixed
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD IY, BC|DE|IY|SP`
    ElseIf Z80_TestParams(MASK_REG_IY, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IY Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IY"), _
         we still need to validate the params to ensure HL & IY are not mixed
        Let Opcode = OpcodeP_RegisterPair(Z80Param2.Token) Or Q1 Or Z1
    '`ADD A, A|B|C|D|E|H|IXH|IYH|L|IYH|IYL|[HL|IX+$8|IY+$8]`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeZ_Register(Z80Param2.Mask)
    '`ADD A, $8`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iAND:
    'Bitwise AND: _
     ..................................................................................
    '`AND A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y4 Or OpcodeZ_Register(Z80Param1.Mask)
    '`AND $8`
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y4 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iBIT:
    'Test a bit: _
     ..................................................................................
    '`BIT 0-7, A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X1 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`BIT 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - Undocumented set bit & copy to register
    ElseIf Z80_TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X1 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iCALL:
    'Call Procedure: _
     ..................................................................................
    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
    If Z80_TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token) Or Z4: GoTo Output16
    '`CALL $16` - absolute call
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Q1 Or Z5: GoTo Output16
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iCCF:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Y7 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCP:
    '..................................................................................
    '`CP A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y7 Or OpcodeZ_Register(Z80Param1.Mask)
    '`CP $8`
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y7 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iCPD:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y5 Or Z1: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPDR:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y7 Or Z1: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iCPI:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y4 Or Z1: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPIR:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y6 Or Z1: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPL:
    'Compliment: Flip the bits of the Accumulator _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = Y5 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDAA:
    'Decimal Adjust Accumulator: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = Y4 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iDEC:
    'Decrement: _
     ..................................................................................
    '`DEC BC|DE|HL|SP|IX|IY`
    If Z80_TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Q1 Or Z3
    '`DEC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    ElseIf Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z5
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iDI:
    'Disable Interrupts: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X3 Or Y6 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDJNZ:
    'Decrement, Jump if Not Zero: _
     ..................................................................................
    If Z80_TestParams(MASK_VAL) Then: Let Opcode = Y2: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
     
iEI:
    'Enable Interrupts: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X3 Or Y7 Or Z3: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iEX:
    'Exchange: _
     ..................................................................................
    '`EX AF AF'`
    If Z80_TestParams(MASK_REG_AF, MASK_REG_AF) Then
        Let Opcode = Y1
    '`EX (SP), HL|IX|IY`
    ElseIf Z80_TestParams(MASK_MEM_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or Y4 Or Z3
    '`EX DE, HL`
    ElseIf Z80_TestParams(MASK_REG_DE, MASK_REG_HL) Then
        Let Opcode = X3 Or Y5 Or Z3
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iEXX:
    'Exchange All: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X3 Or Q1 Or P1 Or Z1: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iHALT:
    'Halt: Stop the processor and wait for interrupts _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Y6 Or Z6: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iIM:
    'Interrupt Mode: _
     ..................................................................................
    '`IM 0|1|2`
    If Z80_TestParams(MASK_VAL) Then
        If Z80Param1.Value = 0 Then
            Let Opcode = X1 Or Z6
        ElseIf Z80Param1.Value = 1 Then
            Let Opcode = X1 Or Y2 Or Z6
        ElseIf Z80Param1.Value = 2 Then
            Let Opcode = X1 Or Y3 Or Z6
        Else
            'ERROR: Wrong parameters for Z80 instruction
            GoTo InvalidParams
        End If
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iIN:
    'Read from port: _
     ..................................................................................
    '`IN A, ($8)`
    If Z80_TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = X3 Or Y3 Or Z3
    '`IN A|B|C|D|E|H|L, [C]`
    ElseIf Z80_TestParams(MASK_REGS_ABCDEHL, MASK_MEM_BC) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask): GoTo OutputED
    '`IN (C)`
    ElseIf Z80_TestParams(MASK_MEM_BC) Then
        Let Opcode = X1 Or Y6: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iINC:
    'Increment: _
     ..................................................................................
    '`INC BC|DE|HL|SP|IX|IY`
    If Z80_TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Z3
    '`INC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    ElseIf Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z4
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iIND:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y5 Or Z2: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINDR:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y7 Or Z2: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iINI:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y4 Or Z2: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINIR:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y6 Or Z2: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iJP:
    'Jump: _
     ..................................................................................
    '`JP [HL|IX|IY]` or `JP HL|IX|IY` - indirect absolute jump
    If Z80_TestParams(MASK_MEM_HLIXY Or MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P2 Or Q1 Or Z1
    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
    ElseIf Z80_TestParams(MASK_FLAGS, MASK_VAL) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token) Or Z2: GoTo Output16
    '`JP $16` - absolute jump
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Z3: GoTo Output16
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iJR:
    'Jump Relative: _
     ..................................................................................
    '`JR $8` - relative immediate jump
    If Z80_TestParams(MASK_VAL) Then
        Let Opcode = Y3
    '`JR NZ|Z|NC|C, $8` - conditional relative immediate jump
    ElseIf Z80_TestParams(MASK_FLAGS_CZ, MASK_VAL) Then
        Let Opcode = OpcodeY_Condition(Z80Param1.Token) + Y4
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iLD:
    'The Mother-Load: _
     ..................................................................................
    '`LD BC|DE|HL|SP|IX|IY, $16`
    If Z80_TestParams(MASK_REGS_BC_DE_HL_SP_IXY, MASK_VAL) Then
        Let Opcode = OpcodeP_RegisterPair(Z80Param1.Token) Or Z1: GoTo Output16
    '`LD [BC], A`
    ElseIf Z80_TestParams(MASK_MEM_BC, MASK_REG_A) Then
        Let Opcode = Z2
    '`LD [DE], A`
    ElseIf Z80_TestParams(MASK_MEM_DE, MASK_REG_A) Then
        Let Opcode = P1 Or Z2
    '`LD A, [BC]`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_MEM_BC) Then
        Let Opcode = Q1 Or Z2
    '`LD A, [DE]`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_MEM_DE) Then
        Let Opcode = P1 Or Q1 Or Z2
    '`LD [$16], HL|IX|IY`
    ElseIf Z80_TestParams(MASK_MEM_VAL, MASK_REGS_HL_IXY) Then
        Let Opcode = P2 Or Z2: GoTo Output16
    '`LD [$16], A`
    ElseIf Z80_TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = P3 Or Z2: GoTo Output16
    '`LD HL|IX|IY, [$16]`
    ElseIf Z80_TestParams(MASK_REGS_HL_IXY, MASK_MEM_VAL) Then
        Let Opcode = P2 Or Q1 Or Z2: GoTo Output16
    '`LD A, [$16]`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let Opcode = P3 Or Q1 Or Z2: GoTo Output16
    '`LD A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8], $8`
    ElseIf Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY, MASK_VAL) Then
        Let Opcode = OpcodeY_Register(Z80Param1.Mask) Or Z6
    'You can't combine IX and IY use in the same opcode, there's only one prefix, _
     so these two following checks ensure that IX & IY are not used together
    '`LD A|B|C|D|E|H|IXH|L|IXL[HL|IX+$8], A|B|C|D|E|H|IXH|L|IXL`
    ElseIf Z80_TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD A|B|C|D|E|H|IYH|L|IYL[HL|IY+$8], A|B|C|D|E|H|IYH|L|IYL`
    ElseIf Z80_TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    'You can't have `LD [IX+$8], [IX+$8]!
    '`LD A|B|C|D|E|H|IXH|L|IXL, A|B|C|D|E|H|IXH|L|IXL|[HL|IX+$8]`
    ElseIf Z80_TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD A|B|C|D|E|H|IYH|L|IYL, A|B|C|D|E|H|IYH|L|IYL|[HL|IY+$8]`
    ElseIf Z80_TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY _
    ) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param1.Mask) Or OpcodeZ_Register(Z80Param2.Mask)
    '`LD SP, HL|IX|IY`
    ElseIf Z80_TestParams(MASK_REG_SP, MASK_REGS_HL_IXY) Then
        Let Opcode = X3 Or P3 Or Q1 Or Z1
    '`LD [$16], BC|DE|SP`
    ElseIf Z80_TestParams(MASK_MEM_VAL, MASK_REGS_BC_DE_SP) Then
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param2.Token) Or Z3
        GoTo OutputED16
    '`LD BC|DE|SP, ($16)`
    ElseIf Z80_TestParams(MASK_REGS_BC_DE_SP, MASK_MEM_VAL) Then
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Q1 Or Z3
        GoTo OutputED16
    '`LD I, A`
    ElseIf Z80_TestParams(MASK_REG_I, MASK_REG_A) Then
        Let Opcode = X1 Or Z7: GoTo OutputED
    '`LD R, A`
    ElseIf Z80_TestParams(MASK_REG_R, MASK_REG_A) Then
        Let Opcode = X1 Or Y1 Or Z7: GoTo OutputED
    '`LD A, I`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_REG_I) Then
        Let Opcode = X1 Or Y2 Or Z7: GoTo OutputED
    '`LD A, R`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_REG_R) Then
        Let Opcode = X1 Or Y3 Or Z7: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iLDD:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y5: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iLDDR:
    '..................................................................................
    If Z80_TestParams() Then: Let Opcode = X2 Or Y7: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iLDI:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y4: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iLDIR:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y6: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iNEG:
    'Negate the Accumulator: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Z4: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iNOP:
    'No Operation: _
     ..................................................................................
    '`NOP` - Opcode is 0
    If Z80_TestParams() Then: Let Opcode = 0: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOR:
    'Bitwise OR: _
     ..................................................................................
    '`OR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y6 Or OpcodeZ_Register(Z80Param1.Mask)
    '`OR $8`
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y6 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iOUT:
    'Write to port: _
     ..................................................................................
    '`OUT [$8], A`
    If Z80_TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let Opcode = X3 Or Y2 Or Z3
    '`OUT [C], A|B|C|D|E|H|L`
    ElseIf Z80_TestParams(MASK_MEM_BC, MASK_REGS_ABCDEHL) Then
        Let Opcode = X1 Or OpcodeY_Register(Z80Param2.Mask) Or Z1: GoTo OutputED
    '`OUT [C], 0`
    ElseIf Z80_TestParams(MASK_MEM_BC, MASK_VAL) Then
        Let Opcode = X1 Or Y6 Or Z1: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iOUTD:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y5 Or Z3: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTDR:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y7 Or Z3: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iOUTI:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y4 Or Z3: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTIR:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X2 Or Y6 Or Z3: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iPOP:
    'Fetch from stack: _
     ..................................................................................
    '`POP AF|BC|DE|HL|IX|IY`
    If Z80_TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Z1
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iPUSH:
    'Push a register on to the stack: _
     ..................................................................................
    '`PUSH AF|BC|DE|HL|IX|IY`
    If Z80_TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let Opcode = X3 Or OpcodeP_RegisterPair(Z80Param1.Token) Or Z5
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iRES:
    'Reset a bit: _
     ..................................................................................
    '`RES 0-7 A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X2 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`RES 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RES memory & register
    ElseIf Z80_TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X2 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRET:
    'Return: _
     ..................................................................................
    '`RET NZ|Z|NC|C|PO|PE|P|M` - conditional return
    If Z80_TestParams(MASK_FLAGS) Then
        Let Opcode = X3 Or OpcodeY_Condition(Z80Param1.Token)
    '`RET`
    ElseIf Z80_TestParams() Then
        Let Opcode = X3 Or Q1 Or Z1
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iRETI:
    'Return from Interrupt: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Y1 Or Z5: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRETN:
    'Return from NMI: _
     ..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Z5: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLA:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Y2 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRL:
    '..................................................................................
    '`RL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y2 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RL memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y2 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iRLCA:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLC:
    '..................................................................................
    '`RLC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = OpcodeZ_Register(Z80Param1.Mask)
    '`RLC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RLC memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRLD:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Y5 Or Z7: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRRA:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Y3 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRR:
    '..................................................................................
    '`RR A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y3 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RR [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RR memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y3 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iRRCA:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Y1 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRRC:
    '..................................................................................
    '`RRC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y1 Or OpcodeZ_Register(Z80Param1.Mask)
    '`RRC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RRC memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y1 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iRRD:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = X1 Or Y4 Or Z7: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRST:
    '"Reset" - Fire specific interrupt vector: _
     ..................................................................................
    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
    If Z80_TestParams(MASK_VAL) Then
        'TODO: Need to validate the number is in range at some point in the chain
        Let Opcode = X3 Or OpcodeY_RST(Z80Param1.Value) Or Z7
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iSBC:
    '..................................................................................
    '`SBC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y3 Or OpcodeZ_Register(Z80Param2.Mask)
    '`SBC A, $8`
    ElseIf Z80_TestParams(MASK_REG_A, MASK_VAL) Then
        Let Opcode = X3 Or Y3 Or Z6
    '`SBC HL, BC|DE|HL|SP`
    ElseIf Z80_TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let Opcode = X1 Or OpcodeP_RegisterPair(Z80Param2.Token) Or Z2: GoTo OutputED
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output

iSCF:
    '..................................................................................
    If Z80_TestParams() Then Let Opcode = Y6 Or Z7: GoTo Output
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iSET:
    'Set a bit: _
     ..................................................................................
    '`SET 0-7 A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = X3 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param2.Mask)
    '`SET 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SET memory & register
    ElseIf Z80_TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = X3 Or OpcodeY_Bit(Z80Param1.Value) Or OpcodeZ_Register(Z80Param3.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSLA:
    '..................................................................................
    '`SLA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y4 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SLA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLA memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y4 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iSRA:
    '..................................................................................
    '`SRA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y5 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SRA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRA memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y5 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSLL:
    '..................................................................................
    '`SLL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y6 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SLL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLL memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y6 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB
    
iSRL:
    '..................................................................................
    '`SRL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let Opcode = Y7 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SRL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRL memory and copy to register
    ElseIf Z80_TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let Opcode = Y7 Or OpcodeZ_Register(Z80Param2.Mask)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    'This is a CB-prefixed double-opcode
    GoTo OutputCB

iSUB:
    'Subtract: _
     ..................................................................................
    '`SUB A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y2 Or OpcodeZ_Register(Z80Param1.Mask)
    '`SUB $8`
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y2 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo Output
    
iXOR:
    'Bitwise XOR: _
     ..................................................................................
    '`XOR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
    If Z80_TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let Opcode = X2 Or Y5 Or OpcodeZ_Register(Z80Param1.Mask)
    '`XOR $8`
    ElseIf Z80_TestParams(MASK_VAL) Then
        Let Opcode = X3 Or Y5 Or Z6
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If

    '----------------------------------------------------------------------------------
    'Write the assembled bytes to the Chunk
    '----------------------------------------------------------------------------------
Output:
    '..................................................................................
    'TODO: Validate 8-bit value?
    
    Let Bytes(ByteCount) = Opcode: Let ByteCount = ByteCount + 1
    Call api_RtlMoveMemory(LogPtr + 20, StrPtr(Cache_HexStr8(Opcode)), 4)
    
    If Offset <> -1 Then Let Bytes(ByteCount) = Offset: Let ByteCount = ByteCount + 1
    
    If Value <> -1 Then
        Let Bytes(ByteCount) = Value: Let ByteCount = ByteCount + 1
        Call api_RtlMoveMemory(LogPtr + 38, StrPtr(Cache_HexStr8(Value)), 4)
    End If
    
    GoTo AddBytes
    
Output16:
    '..................................................................................
    'TODO: Validate 16-bit Value
    
    Let Bytes(ByteCount) = Opcode: Let ByteCount = ByteCount + 1
    Call api_RtlMoveMemory(LogPtr + 20, StrPtr(Cache_HexStr8(Opcode)), 4)
    
    If Value <> -1 Then
        Let Bytes(ByteCount) = (Value And &HFF&): Let ByteCount = ByteCount + 1
        Call api_RtlMoveMemory(LogPtr + 38, StrPtr(Cache_HexStr8(Value And &HFF&)), 4)
        Let Bytes(ByteCount) = ((Value \ &H100&) And &HFF&): Let ByteCount = ByteCount + 1
        Call api_RtlMoveMemory(LogPtr + 44, StrPtr(Cache_HexStr8((Value \ &H100&) And &HFF&)), 4)
    End If
    
    GoTo AddBytes
    
OutputED:
    '..................................................................................
    'Push the implied prefix on the list of bytes to output
    Let Bytes(ByteCount) = &HED&: Let ByteCount = ByteCount + 1
    Call api_RtlMoveMemory(LogPtr + 14, StrPtr(Cache_HexStr8(&HED&)), 4)
    
    Let Bytes(ByteCount) = Opcode: Let ByteCount = ByteCount + 1
    Call api_RtlMoveMemory(LogPtr + 20, StrPtr(Cache_HexStr8(Opcode)), 4)
    
    If Offset <> -1 Then Let Bytes(ByteCount) = Offset: Let ByteCount = ByteCount + 1
    
    If Value <> -1 Then
        Let Bytes(ByteCount) = Value: Let ByteCount = ByteCount + 1
        Call api_RtlMoveMemory(LogPtr + 38, StrPtr(Cache_HexStr8(Value)), 4)
    End If
    
    GoTo AddBytes
    
OutputED16:
    '..................................................................................
    'Push the implied prefix on the list of bytes to output
    Let Bytes(ByteCount) = &HED&: Let ByteCount = ByteCount + 1
    Let Bytes(ByteCount) = Opcode: Let ByteCount = ByteCount + 1
    
    If Offset <> -1 Then Let Bytes(ByteCount) = Offset: Let ByteCount = ByteCount + 1
    
    If Value <> -1 Then
        Let Bytes(ByteCount) = (Value And &HFF&): Let ByteCount = ByteCount + 1
        Let Bytes(ByteCount) = ((Value \ &H100&) And &HFF&): Let ByteCount = ByteCount + 1
    End If
    
    Call api_RtlMoveMemory(LogPtr + 14, StrPtr(Cache_HexStr8(&HED&)), 4)
    Call api_RtlMoveMemory(LogPtr + 20, StrPtr(Cache_HexStr8(Opcode)), 4)
    
    GoTo AddBytes
    
OutputCB:
    '..................................................................................
    'The double-opcodes use $CB as the primary opcode and then a second opcode after _
     the offset byte. It's fastest to handle and log this special case specifically
    Let Bytes(ByteCount) = &HCB&
    Call api_RtlMoveMemory(LogPtr + 20, StrPtr(Cache_HexStr8(&HCB&)), 4)
    
    If Offset <> -1 Then Let Bytes(ByteCount) = Offset: Let ByteCount = ByteCount + 1
    
    Let Bytes(ByteCount) = Opcode: Let ByteCount = ByteCount + 1
    Call api_RtlMoveMemory(LogPtr + 32, StrPtr(Cache_HexStr8(Opcode)), 4)
    
AddBytes:
    '----------------------------------------------------------------------------------
    'Now the assembled bytes have been queued, add them to the Chunk
    With Chunks(Chunk)
        If ChunkMode <> MODE_Z80 Then
            If ChunkMode = MODE_DATA Then Call .LogASM.LogFormat( _
                LOG_DATA, _
                ChunkDataAddr, ChunkData, _
                "·", "·", "·", "·", "·", "·", "·", "·", _
                "·", "·", "·", "·", "·", "·", "·", "·" _
            )
    
            If ChunkMode <> MODE_UNSET Then Call .LogASM.Log( _
                "-----+-------------------+-----------------------------+-----------------" _
            )
            Call .LogASM.Log(LOG_ASM_COLS)
            Call .LogASM.Log(LOG_ASM_HEAD)
            Let ChunkMode = MODE_Z80
        End If
                
        Dim Index As Long
        Let Index = .Size
        
        'Write the current Chunk address on the disassembly log
        Call api_RtlMoveMemory(LogPtr, StrPtr(Cache_HexStr16(Index)), 8)
        
        'Increase the Chunk accordingly
        Let .Size = Index + ByteCount
        
        'If the Chunk's buffer is full then double the buffer size to compensate. _
         This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, ...
        If .Size >= .Buffer Then
            Let .Buffer = .Buffer * 2
            ReDim Preserve .Data(0 To .Buffer - 1) As Byte
        End If
        
        'Copy the bytes into the Chunk
        Call api_RtlMoveMemory( _
            DestinationPointer:=VarPtr(.Data(Index)), _
                 SourcePointer:=VarPtr(Bytes(0)), _
                        Length:=ByteCount _
        )
        
        'Add the disassembly to the log
        Call .LogASM.Log(LogStr)
    End With
    
    GoTo Finish
    
InvalidParams:
    '----------------------------------------------------------------------------------
    'ERROR: Wrong parameters for Z80 instruction
    Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
    
Finish:
    'Return the error number
    Let Z80_Assemble = My_Error
End Function

'Z80_TestParams : Check if parameters are of the right type desired
'======================================================================================
'[Test1]        | A bit-pattern of allowed parameters to check parameter 1 against
'[Test2]        | A bit-pattern of allowed parameters to check parameter 2 against
'[Test3]        | A bit-pattern of allowed parameters to check parameter 3 against
'---------------+----------------------------------------------------------------------
'Returns        | True if all given parameters fulfill the tests, False otherwise
'======================================================================================
Private Function Z80_TestParams( _
    Optional ByRef Test1 As OZ80_MASK, _
    Optional ByRef Test2 As OZ80_MASK, _
    Optional ByRef Test3 As OZ80_MASK _
) As Boolean
    'Why call the function with no parameters? It's a consistent way to check when _
     there *should* be no parameters for an instruction!
    
    'Are we testing a first parameter?
    If (Test1 = 0) Then
        'If a parameter exists, but there shouldn't be, return False
        If (Z80Param1.Mask <> 0) Then Exit Function
        'If there is no first parameter, pass True
        If (Z80Param1.Mask = 0) Then GoTo Finish
    End If
    
    '1st Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param1.Mask And Test1) = 0 Then Exit Function
    
    'Are we testing a second parameter?
    If (Test2 = 0) Then
        'If there shouldn't be a second parameter, return False
        If (Z80Param2.Mask <> 0) Then Exit Function
        'If there is no second parameter, the test has passed
        If (Z80Param2.Mask = 0) Then GoTo Finish
    End If
    
    '2nd Parameter Test: _
     ----------------------------------------------------------------------------------
        
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param2.Mask And Test2) = 0 Then Exit Function
    
    'If there shouldn't be a third parameter, return False
    If (Test3 = 0) And (Z80Param3.Mask <> 0) Then Exit Function
    'If there is no third parameter, the test has passed
    If (Test3 = 0) And (Z80Param3.Mask = 0) Then GoTo Finish
    
    '3rd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Z80Param3.Mask And Test3) = 0 Then Exit Function
    
Finish:
    Let Z80_TestParams = True
End Function

'GetListAsDictionary : Reads a whole list into a Dictionary
'======================================================================================
Private Function GetListAsDictionary() As Scripting.Dictionary
    'Instantiate the return object
    Set GetListAsDictionary = New Scripting.Dictionary
    
    'As long as the list continues, keep fetching Values
    Do While TokenIsExpression(TokenKind)
        'Parse the Expression; if it's indefinite, _
         leave that up to the caller to handle
        Dim Value As Double
        'If there was an error in the expression, the error event _
         will already have been raised, we only need to fold upwards
        If ParseExpression(Value) Then Exit Function
        
        Call GetListAsDictionary.Add( _
            GetListAsDictionary.Count, Value _
        )
    Loop
End Function

'GetText : Reads a list, interpretting Values as Unicode points
'======================================================================================
'ReturnText     | A bluString object that will be populated with the text
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function GetText( _
    ByRef ReturnText As bluString _
) As OZ80_ERROR
    'If the bluString is not instantiated, do so now
    If ReturnText Is Nothing Then Set ReturnText = New bluString
    
    'If the current token is not a list item, then the list has come to an end
    Do While TokenIsList(TokenKind)
    
        'If the current token is text, append it to what we have
        If TokenKind = TOKEN_TEXT Then
            Call ReturnText.Append(TokenText)
        
        'Otherwise it'll be a calculable Expression, _
         hopefully providing an ASCII value
        ElseIf TokenIsExpression(TokenKind) Then
            'Get the Expression Value
            Dim Value As Double
            'If there were an error in parsing the Expression, fold up
            If ParseExpression(Value) Then GoTo ErrOut
            
            'TODO: Indefnite values? Allowed or not in Text? _
                   If we allow for text mapping with different sized strings then _
                   the string size cannot be determinate
            
            'Any UCS-2 value is allowed, it's up to the caller _
             to do text-mapping and validation
            If (Value < 0) Or (Value > 65535) Then
                Call RaiseError(OZ80_ERROR_TEXT_CHAR)
                GoTo ErrOut
            End If
        End If
    
        'Continue reading the list. Note that when the file ends, so does the list _
         and this is not an error as far as this function is concerned. If the parent _
         is within a scope and expecting the closing bracket, then it's up to them _
         to raise the error
        Call NextToken
    Loop
    
    'The resultant string won't be changing any more
    Call ReturnText.ReduceBuffer
ErrOut:
    'Return the error number
    Let GetText = My_Error
End Function

'GetListItem : Read in an Expression from a list
'======================================================================================
Private Function GetListItem( _
             ByRef Value As Double, _
    Optional ByRef Length As Long = 0 _
) As Boolean
    'If the current token is not an expression, the list has ended; return False
    If Not TokenIsExpression(TokenKind) Then Exit Function
    
    'Parse the Expression; if it's indefinite, _
     leave that up to the caller to handle
    '(If there was an error in the Expression, the error event _
      will already have been raised, we only need to fold upwards)
    If ParseExpression(Value) Then Exit Function
    
    'Increases the number of items read from the list thus far
    Let Length = Length + 1
    
    'Return True as the list has not ended yet
    Let GetListItem = True
End Function

'ParseExpression : Evaluate a calculation into a Value
'======================================================================================
'ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
'               | Note that this is a 64-bit floating-pointer number, so it's up to
'               | the caller to validate the number according to their needs
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function ParseExpression( _
    ByRef ReturnValue As Double _
) As OZ80_ERROR
    Dim Operand1 As Double              'Left-hand value in a calculation
    Dim Operand2 As Double              'Right-hand value in a calculation
    Dim Operator As OZ80_TOKEN          'Operator to use
    
    'Begin with fetching the first operand
    If ParseValue(Operand1) Then GoTo ErrOut
    'For the moment the Expression's calculated Value will be this unless we find _
     an Operator to extend the Expression
    Let ReturnValue = Operand1
    
    'The file could legitimately end with an Expression, so don't error out if the _
     token stream comes to an end since the Expression is valid. The parent scopes _
     will handle an unexpected end error
    If EOF Then Exit Function
    
    'If next token is not an Operator assume the Expression has finished, _
      the parent will handle whatever the next token is used for
    If Not TokenIsOperator(TokenKind) Then Exit Function
    
Continue:
    'Remember the Operator for after we have operand 2 to work with
    Let Operator = TokenKind
    
    'Get the second operand
    'TODO: Error - unexpected end
    If Not NextToken() Then Stop
    'Cannot have an operator hanging there without an operand following
    If ParseValue(Operand2) Then Stop
        
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If (Operand1 = INDEFINITE) _
    Or (Operand2 = INDEFINITE) Then
        'Mark expression result as indefinite and skip applying the operator
        Let ReturnValue = INDEFINITE
    
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
        
    ElseIf Operator = TOKEN_OPERATOR_XOR Then
        Let ReturnValue = Operand1 Xor Operand2
        
    Else
        Stop
    End If
    
    'Does the expression continue?
    If TokenIsOperator(TokenKind) Then
        'Set the current result as the first operand and go around again
        Let Operand1 = ReturnValue
        GoTo Continue
    End If
    
ErrOut:
    'Return the error number
    Let ParseExpression = My_Error
End Function

'ParseValue : Get a single Value as part of an Expression
'======================================================================================
'ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
'               | Note that this is a 64-bit floating-pointer number, so it's up to
'               | the caller to validate the number according to their needs
'---------------+----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Private Function ParseValue( _
    ByRef ReturnValue As Double _
) As OZ80_ERROR
    '----------------------------------------------------------------------------------
    'A nested Expression?
    '----------------------------------------------------------------------------------
    If TokenKind = TOKEN_PARENS_OPEN Then
        'Move over the open parentheses. If the file ends here, it's an error
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Parse the nested Expression and return its Value as our own
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        
        'The nested Expression has to be closed
        If TokenKind <> TOKEN_PARENS_CLOSE _
            Then Call RaiseError(OZ80_ERROR_EXPECTED_BRACKET): GoTo ErrOut
        
        'Move over the closing parentehses. The Expression can end here, so it's not _
         necessarily an error if the file suddenly ends. The parent can check for _
         any scopes remaining open
        Call NextToken
        
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'The `HI` function -- returns the high-byte of a 16-bit Value
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_HI Then
        'Move over the `HI` keyword. If the file ends here, it's an error
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Fetch the parameter to the `HI` function
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        'If that is indefinite, so will our result
        If ReturnValue = INDEFINITE Then Exit Function
        
        'Check that this is no more than a 16-bit number
        If ReturnValue > 65535 Then Call RaiseError(OZ80_ERROR_OVERFLOW_HILO): GoTo ErrOut
        'Extract the high-byte
        Let ReturnValue = (ReturnValue And &HFF00&) \ &H100&
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'The `LO` function -- returns the low-byte of a 16-bit Value
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_LO Then
        'Move over the `LO` keyword. If the file ends here, it's an error
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        
        'Fetch the parameter to the `LO` function
        If ParseExpression(ReturnValue) Then GoTo ErrOut
        'If that is indefinite, so will be our result
        If ReturnValue = INDEFINITE Then Exit Function
        
        'Check that this is no more than a 16-bit number _
         (despite this not being an issue to the low 8-bits, we prevent numbers _
          higher than 16-bit to protect against unintended behaviour)
        If ReturnValue > 65535 Then Call RaiseError(OZ80_ERROR_OVERFLOW_HILO): GoTo ErrOut
        'Extract the low-byte
        Let ReturnValue = ReturnValue And &HFF&
        'The token pointer will already be at the next token to read, _
         so we exit without moving it forward any more
        Exit Function
        
    '----------------------------------------------------------------------------------
    'Negative Numbers
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_OPERATOR_SUB Then
        'If the negative sign preceeds, remember this until the full number has _
         been calculated. ONLY decimal numbers can be negative, you cannot negate _
         hexadecimal/binary numbers or named-items (Constants, RAM names, Labels etc.)
        Dim IsNegative As Boolean: Let IsNegative = True
        'Move forward past the negative sign. If the file ends here, it's an error
        If Not NextToken() Then Call RaiseError(OZ80_ERROR_FILE_END): GoTo ErrOut
        'Continue parsing the number...
        GoTo GetNumber
        
    '----------------------------------------------------------------------------------
    'A decimal number:
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_NUMBER Then
GetNumber:
        'Begin with getting the given value
        Let ReturnValue = TokenValue
        'Was there a negative sign preceeding this number?
        If IsNegative Then Let ReturnValue = -ReturnValue
        
        'Move over the number. If the file ends here, the Value is still valid, _
         so no error. The parent can check if any scope has been left open. We skip _
         checking for a number prefix, but we do need to apply any negative sign
        If NextToken() Then
            'Check for a number prefix, e.g. `KB`
            '..........................................................................
            If TokenKind = TOKEN_PREFIX_K Then
                'Check for min/max number you can multiply by 1'000 without overflowing
                If (ReturnValue < NUMBER_MIN_K) Or (ReturnValue > NUMBER_MAX_K) Then
                    Call RaiseError(OZ80_ERROR_OVERFLOW)
                    GoTo ErrOut
                End If
                'Multiply by 1'000
                Let ReturnValue = ReturnValue * 1000
                'Move over the prefix
                Call NextToken
                
            '..........................................................................
            ElseIf TokenKind = TOKEN_PREFIX_KB Then
                'Check for min/max number you can multiply by 1'024 without overflowing
                If (ReturnValue < NUMBER_MIN_KB) Or (ReturnValue > NUMBER_MAX_KB) Then
                    Call RaiseError(OZ80_ERROR_OVERFLOW)
                    GoTo ErrOut
                End If
                'Multiply by 1'024
                Let ReturnValue = ReturnValue * 1024
            Else
                'If the token is not a prefix, exit this function without moving the _
                 token pointer forward (as we've already done that)
                GoTo ErrOut
            End If
        End If
    
    '----------------------------------------------------------------------------------
    'Constants
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_CONST Then
        'Has the constant been defined?
        If Not Values.Exists(TokenText) Then
            'ERROR: Constant not yet defined. Constants cannot be indefinite
            Call RaiseError(OZ80_ERROR_UNDEFINED_CONST)
            GoTo ErrOut
        End If
        'Retrieve the Constant value
        Let ReturnValue = Values.Item(TokenText)
        
    '----------------------------------------------------------------------------------
    'RAM Names
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_RAM _
        Or TokenKind = TOKEN_HASH _
    Then
        'Has the RAM Name / Hash value been defined?
        If Not Values.Exists(TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the RAM address / Has Value
            Let ReturnValue = Values.Item(TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    'The `START` keyword refers to the System's boot vector
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_KEYWORD_START Then
        'TODO: This is determined by the System, so we can forgo having to reference it
        If Not Labels.Exists(STR_START) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the START address
            Let ReturnValue = Labels.Item(STR_START)
        End If
    
    '----------------------------------------------------------------------------------
    'Labels
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_LABEL Then
        'Label names have to be expanded to their fully-qualified form
        'TODO: If a Label is not found in this Section, check the global Section
        If Not Labels.Exists(CurrentSection & TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the RAM address
            Let ReturnValue = Labels.Item(CurrentSection & TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    'Properties
    '----------------------------------------------------------------------------------
    ElseIf TokenKind = TOKEN_PROPERTY_USE Then
        'Property names have to be expanded to their fully-qualified form
        If Not Labels.Exists(CurrentSection & CurrentLabel & TokenText) Then
            'Return Indefinite
            Let ReturnValue = INDEFINITE
        Else
            'Retrieve the sub-routine address
            Let ReturnValue = Labels.Item(CurrentSection & CurrentLabel & TokenText)
        End If
        
    '----------------------------------------------------------------------------------
    Else
        Debug.Print TokenNames(TokenKind)
        Stop
    End If
    
    'Move over the processed token, the parent expects the token pointer _
     to be on the next unread token
    Call NextToken
ErrOut:
    'Return the error number
    Let ParseValue = My_Error
End Function

'Chunk_Begin : Begin a new Chunk to hold assembled Z80 code / data
'======================================================================================
Private Sub Chunk_Begin( _
             ByRef LabelName As String, _
    Optional ByRef Address As Double = -1 _
)
    'Register the Label name in the database
    Call AddUpdateLabel(LabelName, Address)
    'Does the Chunk already exist? On the first pass we want to ensure that the user _
     doesn't try and define two Chunks with the same name, and this is handled by the _
     `PROC` handler. On the second pass the Chunks will already have been defined
    If ChunkExists(LabelName) Then
        'Attach the existing Chunk
        Let Chunk = GetChunk(LabelName)
    Else
        Let ChunkCount = ChunkCount + 1
        ReDim Preserve Chunks(1 To ChunkCount) As oz80Chunk
        Let Chunk = ChunkCount
        Let Chunks(Chunk).LabelName = LabelName
        Let Chunks(Chunk).Address = CLng(Address)
        Let Chunks(Chunk).Buffer = 1024
        ReDim Chunks(Chunk).Data(0 To 1024 - 1) As Byte
        
        Set Chunks(Chunk).LogASM = New bluLog
    End If
End Sub

'AddUpdateLabel: Blindly adds to or updates the Label database
'======================================================================================
'LabelName      | The fully-qualified Label name, i.e. including the Section name
'[Value]        | The new address to set for the Label
'======================================================================================
Private Sub AddUpdateLabel( _
             ByRef LabelName As String, _
    Optional ByRef Value As Double = 0 _
)
    'Add a new label or update an existing label value?
    If Not Labels.Exists(LabelName) _
        Then Call Labels.Add(LabelName, Value) _
        Else Let Labels.Item(LabelName) = Value
End Sub

'LayoutROM : Calculate the placement of the code/data Chunks in the output ROM
'======================================================================================
'Returns        | The error number
'======================================================================================
Private Function LayoutROM( _
) As OZ80_ERROR
    '----------------------------------------------------------------------------------
    '[1]: Place the Interrupt Procedures
    '----------------------------------------------------------------------------------
    If Chunk_Interrupt00 <> 0 Then
    End If
    
    '----------------------------------------------------------------------------------
    '[2]: Place the ROM Header
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    '[3]: Place the START Procedure
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    '[4]: Place the 1 KB SEGA Slot
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    '[5]: Place un-paged Sections
    '----------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    '[6]: Place paged Sections
    '----------------------------------------------------------------------------------
End Function

Private Function LayoutAdd( _
    ByRef ChunkAddr As Long, _
    ByRef ChunkRef As Long _
) As OZ80_ERROR
    'If there are no layout items yet, start the linked list
    If LayoutCount = 0 Then
        Let LayoutCount = 1
        ReDim Layout(1 To 1) As oz80Layout
        'Apply the fields
        'TODO: Update the database value for the Label Name
        Let Chunks(ChunkRef).Address = ChunkAddr
        Let Layout(1).Chunk = ChunkRef
        
        Let Layout(1).AddrBegin = ChunkAddr
        Let Layout(1).AddrSize = Chunks(ChunkRef).Size
        Let Layout(1).AddrEnd = ChunkAddr + Layout(1).AddrSize
        
        Exit Function
    End If
    
    'Search through the nodes and find the point where this one goes
    Dim i As Long
    Let i = 1
    Do
        'If any current Chunk is in the same place as our desired location, error
        If Layout(i).AddrBegin = ChunkAddr Then
            'TODO: Error when two Chunks want the same address
            Stop
            
        'If a Chunk comes before our desired location, but its size precludes our _
         desired location, then error
        ElseIf Layout(i).AddrBegin < ChunkAddr And Layout(i).AddrEnd > ChunkAddr Then
            'TODO: Error when a Chunk is trying to be placed within another
            Stop
            
        'Check if the current node comes after our Chunk's desired location
        ElseIf Layout(i).AddrBegin > ChunkAddr Then
            
            
        End If
        
        'Follow the chain to the next item in the list
        Let i = Layout(i).NextIndex
        'If this is the end of the list, do not continue
        If i = 0 Then Exit Do
    Loop
End Function

'/// GENERIC //////////////////////////////////////////////////////////////////////////
'/// Functions used anywhere throughout the class

'ChunkExists : Test if a Chunk is in the array by the Label name
'======================================================================================
'LabelName      | Fully-qualified Label Name for the Chunk, e.g. "::Section:Label"
'---------------+----------------------------------------------------------------------
'Returns        | True if the Chunk exists, False otherwise
'======================================================================================
Private Property Get ChunkExists( _
    ByRef LabelName As String _
) As Boolean
    'If there are no Chunks, just return False
    If ChunkCount = 0 Then Exit Property
    'Go over each Chunk and check the name
    Dim i As Long
    For i = 1 To ChunkCount
        If Chunks(i).LabelName = LabelName Then
            Let ChunkExists = True
            Exit Property
        End If
    Next i
End Property

'FileExists : See if a file exists or not
'======================================================================================
'FilePath       | Path to the file to check for
'---------------+----------------------------------------------------------------------
'Returns        | True if the file exists, false if not
'======================================================================================
Private Function FileExists( _
    ByRef FilePath As String _
) As Boolean
    '<cuinl.tripod.com/Tips/fileexist.htm>
    Let FileExists = CBool(Dir$(FilePath) <> vbNullString)
End Function

'GetChunk : Retrieves a Chunk in the array by its Label Name
'======================================================================================
Private Function GetChunk( _
    ByRef LabelName As String _
) As Long
    'Go over each Chunk and check the name
    Dim i As Long
    For i = 1 To ChunkCount
        If Chunks(i).LabelName = LabelName Then
            Let GetChunk = i
            Exit Function
        End If
    Next i
End Function

'RaiseMessage : Provide a progress message
'======================================================================================
'LogLevel       | The type of log message, for filtering out lower-level messages
'FormatText     | The message to log, which can include "{$}" markers to be replaced
'               | with the values provided below
'FormatValues...| A list of Variant values (Strings, Longs &c.) to fill in
'======================================================================================
Private Sub RaiseMessage( _
    ByRef LogLevel As OZ80_LOG, _
    ByRef FormatText As String, _
    ParamArray FormatValues() As Variant _
)
    Let LogMsg.Text = FormatText
    
    'Are there any values to format?
    If UBound(FormatValues) < 0 Then
        'No, raise just the text given
        RaiseEvent Message(LogLevel, LogMsg)
    Else
        'Format the given log message before sending it up
        RaiseEvent Message(LogLevel, LogMsg.Format(FormatValues))
    End If
End Sub

'RiaseError : Provide error details through an event
'======================================================================================
'Number         | Error number to set on the Class and raise via Event
'Message        | A context-specific error message
'[FormatValues] | A list of values to replace string markers "{$}"in the error message
'---------------+----------------------------------------------------------------------
'Returns        | The same error number, for simple tail recursion
'======================================================================================
Private Function RaiseError( _
         ByRef Number As OZ80_ERROR _
) As OZ80_ERROR
    'Set the error number on this class
    Let My_Error = Number
    
    'Fetch the detailed error text
    Dim Title As String, Description As String
    Call GetOZ80Error(Number, Title, Description)
    
    'Pass the error up to the controller
    'TODO: Error before parsing the File?
    RaiseEvent Error( _
        Files(File).Path, _
        Number, Title, Description, _
        TokenLine, TokenCol _
    )
    
    'We return the error number, purely for nicer looking tail-recursion
    Let RaiseError = Number
End Function

'GetOZ80Error : Return an error description for a given error number
'======================================================================================
Private Sub GetOZ80Error( _
    ByRef ErrorNumber As OZ80_ERROR, _
    ByRef ReturnTitle As String, _
    ByRef ReturnDescription As String _
)
    Select Case ErrorNumber
    
    Case OZ80_ERROR_DUPLICATE
        '..............................................................................
        Let ReturnTitle = "Duplicate Definition"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_DUPLICATE_SECTION
        '..............................................................................
        Let ReturnTitle = "Duplicate Definition"
        Let ReturnDescription = _
            "You cannot define a section name twice. There should be only one " & _
            "`SECTION` statement for each section in use."
        
    Case OZ80_ERROR_FILE_END
        '..............................................................................
        Let ReturnTitle = "Unexpected End of File"
        'TODO
        Let ReturnDescription = "The file ended "
        
    Case OZ80_ERROR_EXPECTED
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        'TODO
        Let ReturnDescription = _
            ""
        
    Case OZ80_ERROR_EXPECTED_PROC_NAME
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        Let ReturnDescription = _
            "A label name must follow the `PROC` statement. " & _
            "E.g. `PROC :myProcedure`"
    
    Case OZ80_ERROR_EXPECTED_ROOT
        '..............................................................................
        Let ReturnTitle = "Keyword expected"
        Let ReturnDescription = _
            "Expected `DEF`, `IF`, `INCLUDE`, `PROC, `SECTION`, `SYSTEM` or `TABLE` " & _
            "keywords at this scope. Have you correctly closed any brackets that were " & _
            "open?"
    
    Case OZ80_ERROR_EXPECTED_SECTION_NAME
        '..............................................................................
        Let ReturnTitle = "Unexpected Content"
        Let ReturnDescription = _
            "A section name must follow the `SECTION` statement. " & _
            "E.g. `SECTION ::graphics`"
    
    Case OZ80_ERROR_EXPECTED_SYSTEM_NAME
        '..............................................................................
        Let ReturnTitle = "Expected Valid System Name"
        Let ReturnDescription = _
            "A valid System name must follow the `SYSTEM` statement. " & _
            "E.g. `SYSTEM ""SMS""`."
        
    Case OZ80_ERROR_EXPRESSION
        '..............................................................................
        Let ReturnTitle = "Invalid Expression"
        Let ReturnDescription = _
            "An expression can be any Number, Label, Property, RAM Name or " & _
            "calculation (via operators) of these."
        
    Case OZ80_ERROR_EXPRESSION_Z80
        '..............................................................................
        Let ReturnTitle = "Invalid Z80 Instruction Parameter"
        Let ReturnDescription = _
            "Parameters following a Z80 instruction must be either a Z80 Register " & _
            "(`a`, `b`, `c` etc.), a Z80 memory expression `[ix+$FF]` or a valid " & _
            "numerical expression, i.e. a calculation, a label name or RAM name."
            
    Case OZ80_ERROR_FILE_NOTFOUND
        '..............................................................................
        Let ReturnTitle = "File Not Found"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_FILE_READ
        '..............................................................................
        Let ReturnTitle = "Cannot Read File"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INDEFINITE
        '..............................................................................
        Let ReturnTitle = "Cannot Use Indefinite Value"
        Let ReturnDescription = _
            "A variable cannot be defined with an indefinite value, that is, " & _
            "an expression containing a yet-unknown value, such as a label. " & _
            "label addresses are not set until after assembly."
    
    Case OZ80_ERROR_INVALID_NAME
        '..............................................................................
        Let ReturnTitle = "Invalid Name"
        Let ReturnDescription = _
            "Variable, label and property names can contain A-Z, 0-9 underscore " & _
            "and dot with the following exceptions: " & _
            "1. the first letter cannot be a number or a dot, " & _
            "2. two dots cannot occur in a row " & _
            "3. a number cannot follow a dot, and " & _
            "4. the name cannot end in a dot" _
    
    Case OZ80_ERROR_INVALID_NAME_RAM
        '..............................................................................
        Let ReturnTitle = "Invalid Name"
        Let ReturnDescription = _
            "RAM names must begin with '$.' and follow standard naming rules " & _
            "beyond that, i.e." & _
            "1. the first letter cannot be a number or a dot " & _
              "(this does not include the dot that follows the dollar sign)" & _
            "2. two dots cannot occur in a row " & _
            "3. a number cannot follow a dot, and " & _
            "4. the name cannot end in a dot" _
        
    Case OZ80_ERROR_INVALID_NUMBER
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_INVALID_NUMBER_DEC
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_INVALID_NUMBER_HEX
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        Let ReturnDescription = _
            "Hexadecimal numbers must begin with '$' and must contain 0-9 & A-F " & _
            "letters only. E.g. `$1234ABCD`"
    
    Case OZ80_ERROR_INVALID_NUMBER_BIN
        '..............................................................................
        Let ReturnTitle = "Invalid Number"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INVALID_WORD
        '..............................................................................
        Let ReturnTitle = "Invalid Word"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_INVALID_Z80PARAMS
        '..............................................................................
        Let ReturnTitle = "Invalid Parameters For Z80 Instruction"
        'TODO
        Let ReturnDescription = ""
    
    Case OZ80_ERROR_OVERFLOW
        '..............................................................................
        Let ReturnTitle = "Overflow"
        'TODO
        Let ReturnDescription = ""
        
    Case OZ80_ERROR_UNDEFINED_CONST
        '..............................................................................
        Let ReturnTitle = "Undefined Constant Used"
        Let ReturnDescription = _
            "You've used a Constant name which has not been defined yet. " & _
            "Ensure that early on in your source code you define the Constant " & _
            "Value: " & vbCrLf & vbCrLf & "DEF !CONSTANT 123"
    Case Else
'        Stop
    End Select
End Sub
