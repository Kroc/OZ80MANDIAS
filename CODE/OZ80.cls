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

'We'll be doing some basic timing throughout
Public Profiler As New bluProfiler

Public Enum PROFILER_EVENTS
    EVENT_INITIALIZE                    'Class initialisation -- caching
    EVENT_FILELOAD_API
    EVENT_TOKENISE                      'Convert source file into Token Stream
    EVENT_PARSE_WORD                    '- Parse out a single word
    EVENT_BUILDTREE                     'Build the Syntax Tree from a Token Stream
    EVENT_FORMATTOKEN                   '- Format token data for logging
    EVENT_ASSEMBLE                      'Assembling a file (after parsing)
    EVENT_Z80ASSEMBLE                   'Assembling a Z80 instruction
    EVENT_ROM_LAYOUT
    EVENT_ROM_ADDCHUNK
End Enum

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'Errors:
'--------------------------------------------------------------------------------------
'Error number! (for the Class)
Private Error As OZ80_ERROR

Public Enum OZ80_ERROR
    OZ80_ERROR_NONE                     'Assembly completed successfully
    OZ80_ERROR_ABORT                    'The user cancelled the assembly
    OZ80_ERROR_INTERNAL                 'I made a mistake
    
    OZ80_ERROR_DUPLICATE                'A name has been defined twice
    OZ80_ERROR_EXPECTED                 'Incorrect content at the current scope
    OZ80_ERROR_FILE_END                 'Unexpected end of file
    OZ80_ERROR_FILE_NOTFOUND            'Requested file does not exist
    OZ80_ERROR_FILE_READ                'Some kind of problem while file handle open
    OZ80_ERROR_INDEFINITE               'Indefinite value cannot be used here
    OZ80_ERROR_INVALID
    OZ80_ERROR_LAYOUT                   'A problem with ROM layout (Chunks overlap &c.)
    OZ80_ERROR_OVERFLOW                 'A number overflowed the maximum
    OZ80_ERROR_UNDEFINED                'Named item is used, but undefined
End Enum

'Private Const STR_ERR_OVERFLOW As String = _
'    "Overflow."

Public Enum OZ80_WARNING
    OZ80_WARNING_NONE                   'Skip "0"
    OZ80_WARNING_ROUND                  'A decimal number had to be round-down
End Enum

'--------------------------------------------------------------------------------------

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
'File Parsing:
'--------------------------------------------------------------------------------------
'OZ80 converts the source code files into "Tokens", byte-code representations of the _
 source "words" which are then used for parsing. It would be too slow to re-interpret _
 the text every time we needed to cross-reference

'A lot of the text parsing will be letter-by-letter, so this enum provides the _
 common character codes we'll be referring to, as well as the particular punctuation _
 symbols used for OZ80 syntax. Note that `UCS2` is defined by 'blu_Strings.cls'

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
    
    SYNTAX_COMMENT = UCS2.SEMICOLON                 ' ; - Comment marker
    SYNTAX_HINT = UCS2.GRAVE_ACCENT                 ' ` - Register hint, `A`index`
    SYNTAX_SHADOW = UCS2.APOSTROPHE                 ' ' - Shadow register, `ex AF AF'`
    SYNTAX_TEXT = UCS2.QUOTATION_MARK               ' " - String identifier
    SYNTAX_DOC = UCS2.GRAVE_ACCENT                  ' ` - Documentation text "``..."
    SYNTAX_CONSTANT = UCS2.EXCLAMATION_MARK         ' ! - Constant identifier
    SYNTAX_LABEL = UCS2.colon                       ' : - Label identifier
    SYNTAX_PROPERTY = UCS2.FULL_STOP                ' . - Property identifier
    SYNTAX_TYPE = UCS2.NUMBER_SIGN                  ' # - Type identifier
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
End Enum

'The width of a tab stop (this affects the column number given)
'TODO: This should be a class property so that the user can customise
Private Const TAB_WIDTH As Long = 8

'Look up table for quickly converting ASCII hexadecimal digits to decimal
Private Hex2Dec(0 To 255) As Long

'Tokenisation:
'--------------------------------------------------------------------------------------

'When we parse a source file we extract the words and convert each _
 into a simple numerical representation (a token)

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
    TOKEN_Z80_RL                        'Rotate Left
    TOKEN_Z80_RLA                       'Rotate Left (Accumulator)
    TOKEN_Z80_RLC                       'Rotate Left Circular
    TOKEN_Z80_RLCA                      'Rotate Left Circular (Accumulator)
    TOKEN_Z80_RLD                       'Rotate Left 4-bits
    TOKEN_Z80_RR                        'Rotate Right
    TOKEN_Z80_RRA                       'Rotate Right (Accumulator)
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
    TOKEN_Z80_PE                        'Parity/Overflow is set flag
    TOKEN_Z80_PO                        'Parity/Overflow is not set flag
    TOKEN_Z80_R                         'Refresh register (pseudo-random)
    TOKEN_Z80_SP                        'Stack Pointer
    TOKEN_Z80_Z                         'Zero set flag
    [_TOKEN_REGS_END] = TOKEN_Z80_Z
    
    'Operators ........................................................................
    [_TOKEN_OPERATORS_BEGIN]
    TOKEN_OPERATOR_ADD = [_TOKEN_OPERATORS_BEGIN]
    TOKEN_OPERATOR_SUB                  'Subtract "-"
    TOKEN_OPERATOR_MUL                  'Multiply "*"
    TOKEN_OPERATOR_DIV                  'Divide "/"
    TOKEN_OPERATOR_POW                  'Power "^"
    TOKEN_OPERATOR_MOD                  'Modulus "\"
    TOKEN_OPERATOR_REP                  'Repeat "x"
    TOKEN_OPERATOR_OR                   'Bitwise OR "|"
    TOKEN_OPERATOR_AND                  'Bitwise AND "&"
    [_TOKEN_OPERATORS_END] = TOKEN_OPERATOR_AND
    
    'Keywords .........................................................................
    [_TOKEN_KEYWORDS_BEGIN]
    'Boolean data type (1-bit)
    TOKEN_KEYWORD_BOOL = [_TOKEN_KEYWORDS_BEGIN]
    TOKEN_KEYWORD_BYTE                  'Byte data type
    TOKEN_KEYWORD_CONST                 'Define constant
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
    TOKEN_KEYWORD_RAM                   'RAM block
    TOKEN_KEYWORD_RAMPAGE               'RAM Page definition
    TOKEN_KEYWORD_RETURN                'Returns list `PROC :<label> RETURN <list>`
    TOKEN_KEYWORD_SECTION               'Section definition `SECTION ::<section>
    TOKEN_KEYWORD_SLOT                  'Section Slot pattern `SLOT 0, 1, 2`
    TOKEN_KEYWORD_START                 'The starting vector for the System
    TOKEN_KEYWORD_SYSTEM                'System identifier `SYSTEM "SMS"`
    TOKEN_KEYWORD_TABLE                 'Data table
    TOKEN_KEYWORD_TRIP                  'Triple data type (3-bytes)
    TOKEN_KEYWORD_TYPE                  'Define a custom data Type
    TOKEN_KEYWORD_WORD                  'Word data type (2-bytes)
    [_TOKEN_KEYWORDS_END] = TOKEN_KEYWORD_WORD
    
    TOKEN_NUMBER
    'Number suffixes ("K" & "KB")
    TOKEN_SUFFIX_K                      'x1000
    TOKEN_SUFFIX_KB                     'x1024
    
    'Grouping: (i.e. parenthesis, braces)
    TOKEN_BRACES_OPEN                   '"{" Code/data Chunk, `PROC :<label> { ... }`
    TOKEN_BRACES_CLOSE                  '"}"
    TOKEN_SQUARE_OPEN                   '"[" Data Type and
    TOKEN_SQUARE_CLOSE                  '"]" Memory reference `ld a, [hl]`
    TOKEN_PARENS_OPEN                   '"(" Expression nesting,
    TOKEN_PARENS_CLOSE                  '")" e.g. `HI ($8000 + $80)`
    
    'A defined Name; i.e. Constant, Section, Label &c.
     'The particular kind of Name is given by the token attributes (below this enum)
    TOKEN_NAME
    TOKEN_PROPERTY_NEW
    TOKEN_PROPERTY_ANON                 'Anonymous Property, i.e. `.`
    
    TOKEN_TEXT                          'e.g. `"..."`
    TOKEN_DOC                           'Documentation text, i.e. "``..."
    
    [_TOKEN_LAST]                       'Do not go above 255!
End Enum

'A look-up table of tokens (above) to token attributes (below)
Private TokenAttributes(0 To [_TOKEN_LAST]) As OZ80_ATTR

'The tokens fit into a variety of overlapping categories. _
 This bit-mask pattern allows us to compare multiple tokens very quickly
Private Enum OZ80_ATTR
    ATTR_NONE
    
    ATTR_KEYWORD = 2 ^ 0                'A keyword, such as `PROC`
    ATTR_UNIT = 2 ^ 1                   'A data unit, i.e. BYTE/WORD/TRIP/LONG &c.
    ATTR_LIST = 2 ^ 2                   'A list item; Text or Expression
    ATTR_EXPR = 2 ^ 3                   'An Expression (calculable value)
    ATTR_OPERTATOR = 2 ^ 4
    ATTR_PARAM = 2 ^ 5
    
    ATTR_Z80_MNEMONIC = 2 ^ 6           'Z80 instruction, e.g. `djnz`
    ATTR_Z80_REGISTER = 2 ^ 7           'Z80 register, e.g. `A`
    ATTR_Z80_FLAG = 2 ^ 8               'Z80 flag, e.g. `nz`
    
    'For defined names (such as Sections, Labels & RAM) we want to be able to
    ' identify composite names (e.g. "::Section:Label.property") or the lack there-of
    ' (for example, when only a Section name is allowed, not a fully-qualified Label)
    
    ATTR_NAME_CONST = 2 ^ 9             'A Constant, e.g. "!BLUE"
    ATTR_NAME_LABEL = 2 ^ 10            'A Label name, e.g. ":Label", incl. ":"
    ATTR_NAME_PROPERTY = 2 ^ 11         'A Property name, e.g. ".property"
    ATTR_NAME_RAM = 2 ^ 12              'A RAM page name, e.g. "$_SRAM", incl. "$"
    ATTR_NAME_SECTION = 2 ^ 13          'A Section name, e.g. "::section"
    ATTR_NAME_TYPE = 2 ^ 14             'A Type name, e.g. "#type"
    
    ATTR_NAMES = ATTR_NAME_CONST Or ATTR_NAME_LABEL Or ATTR_NAME_PROPERTY Or ATTR_NAME_RAM Or ATTR_NAME_SECTION Or ATTR_NAME_TYPE
    
    'Lists can contains Lists and Expressions
    ATTR_LISTEXPR = ATTR_LIST Or ATTR_EXPR
    
    'During tokenisation we often can't tell the difference between Z80 Registers and
    ' Flags (no context with which to work with), so often they are lumped together
    ATTR_Z80_REGFLAG = ATTR_Z80_REGISTER Or ATTR_Z80_FLAG
    'Z80 Parameters can be Expressions or Registers
    ATTR_Z80_PARAM = ATTR_EXPR Or ATTR_Z80_REGFLAG
End Enum

'--------------------------------------------------------------------------------------

'Each tokenised word in the source will also retain metadata on where it was _
 originally in the source code, for error messages. Below is the structure for a _
 single token, representing one "word" in the source code

Private Type oz80Token
    Kind As OZ80_TOKEN                  'Token identifier
    Attr As OZ80_ATTR                   'Token attributes (Expression, List, Keyword)
    Value As Double                     'Associated value, e.g. for number tokens
    Line As Long                        'Line number in the original source code
    Col As Long                         'Column number in the original source code
End Type

'A token stream is machine-readable representation of the original source code that _
 does away for the need to refer to the source code file again

Private Type oz80TokenStream
    Tokens() As oz80Token               'The token stream itself
    Buffer As Long                      'The pre-allocated length of the stream
    Length As Long                      'The current length of stream used
    
    Strings() As String
    StringCount As Long
End Type

'Use a 20'000 item initial buffer for the token stream
Private Const TOKEN_BUFFER As Long = 20000

'--------------------------------------------------------------------------------------

Private Type oz80File
    'Remember where the original source file is as the assembler will need to change
     'to its directory when handling relative references with includes
    Path As String
    'The token stream representation of the source file will live here
    Stream As oz80TokenStream
End Type

'Each source code file is broken down into a machine-readable representation _
 ('oz80File') which we store here. When assembling one file and we come _
 across an `INCLUDE` statement, another File gets added here
Private Files() As oz80File
Private FileCount As Long
'The current source file being processed
Private File As Long

'--------------------------------------------------------------------------------------
'Abstract Syntax Tree:
'--------------------------------------------------------------------------------------

Private Type oz80TreeNode
    File As Long                        'File number
    Token As oz80Token
    
    NodeParent As Long                  'Always the parent node, even for siblings
    NodeParam As Long                   'Branch for the parameters of the current node
    NodeChild As Long                   'Index of the ASTNode for inside the scope
    NodeNext As Long                    'Index of the ASTNode for the sibling
End Type

Private Type oz80Tree
    Nodes() As oz80TreeNode
    Strings() As String
    StringCount As Long
    
    Index As Long
    
    Count As Long
    Buffer As Long
End Type

Private Tree As oz80Tree
Private Node As oz80TreeNode

'--------------------------------------------------------------------------------------
'Target System:
'--------------------------------------------------------------------------------------
'OZ80 does not aim to be overly generic, but rather target specific systems so as to _
 avoid requiring the developer to manually implement ROM checksums etc.

'Which target system is selected, e.g. SEGA Master System
Private System As OZ80_SYSTEM

'A list of system targets. Only the SEGA Master System is supported at the moment, _
 but I will consider supporting other Z80 systems in the future
Private Enum OZ80_SYSTEM
    SYSTEM_NONE                         'System not yet defined
    SYSTEM_SMS                          'SEGA Master System
End Enum

'--------------------------------------------------------------------------------------
'Sections:
'--------------------------------------------------------------------------------------
'A "Section" is a group of code/data Chunks that will be placed into the ROM together _
 according to special rules for the Slot pattern (system-specific paging scheme)

Private Type oz80Section
    Name As String
    ID As Long                          'Index of the name in `TokenStrings`
    Size As Long                        'Total number of Chunk bytes for the Section
    Slot As OZ80_SLOT                   'The Slot (paging) pattern assigned
    Chunks() As Long                    'A list of which Chunks are in the Section
    ChunkCount As Long                  'Number of the above, 1-based
End Type

'Whilst in the syntax `SLOT` uses a list (i.e. `SLOT 0, 1, 2`), we convert that into _
 a bit pattern to make it quick and easy to work with instead of iterating an array
Private Enum OZ80_SLOT
    SLOT0 = 2 ^ 0
    SLOT1 = 2 ^ 1
    SLOT2 = 2 ^ 2
    SLOT1KB = 1024
End Enum

Private Sections() As oz80Section
Private SectionCount As Long

'--------------------------------------------------------------------------------------

'The database of defined named values _
 (i.e. Constants, Sections, Labels &c.)
'Private Values As Scripting.Dictionary

Private Type oz80Name
    Name As String
    Value As Double
End Type

Private Names() As oz80Name
Private NameCount As Long

'--------------------------------------------------------------------------------------
'Type Definition:
'--------------------------------------------------------------------------------------

Private Type OZ80Property
    Name As String
    Offset As Long
    
    Unit As OZ80_TOKEN
    Size As Long
End Type

Private Type OZ80Type
    Name As String
    Base As Long
    Size As Long                        'Complete size of the Type in bits
    
    Properties() As OZ80Property
    PropertyCount As Long
End Type

Private Types() As OZ80Type
Private TypeCount As Long

'--------------------------------------------------------------------------------------

Private Type oz80RAMPage
    Type As OZ80Type
    
    Used As Long                        'Number of used bytes
    Last As Long                        'Index of the used byte closest to the end
End Type

Private RAMPages() As oz80RAMPage
Private RAMPagesCount As Long

'--------------------------------------------------------------------------------------
'Z80 Instructions:
'--------------------------------------------------------------------------------------
'There are up to three parameters for Z80 instructions. _
 These can be Registers, Values or Z80 Memory References

Private Type oz80Param
    Mask As OZ80_PARAM                  'Bit-pattern (as below) identifiying features
    Token As OZ80_TOKEN                 'Which word it is, from the source code
    
    Value As Double                     'Value of the Param, could be 'INDEFINITE'
    Expression As Long                  'The Token index where the Value Expression is
    Defer As Long                       'If the Expression needs to be deferred
    
    LogText As Long                     'When logging, which cached text to use
End Type

'In order to compare the hundreds of permutations of parameters for Z80 instructions, _
 we assign each parameter type a single bit. We can thus check very quickly if a given _
 parameter falls within an allowed list of accepted types

Private Enum OZ80_PARAM
    PARAM_NONE
    
    [_PARAM_REGS_BEGIN] = 1
    PARAM_REG_B = 2 ^ 0
    PARAM_REG_C = 2 ^ 1
    PARAM_REG_D = 2 ^ 2
    PARAM_REG_E = 2 ^ 3
    PARAM_REG_H = 2 ^ 4
    PARAM_REG_L = 2 ^ 5
    PARAM_MEM_HL = 2 ^ 6
    PARAM_REG_A = 2 ^ 7
    
    'The presence of an IX/IY prefix on the opcode changes H/L to IXH/IYH or IXL/IYL
     'respectively, but only on instructions that use single byte opcodes.
     'This is officially undocumented, but obviously fair game for old systems
    PARAM_REG_IXH = 2 ^ 8
    PARAM_REG_IXL = 2 ^ 9
    PARAM_REG_IYH = 2 ^ 10
    PARAM_REG_IYL = 2 ^ 11
    
    'The presence of an IX/IY prefix on the opcode changes a memory reference "[HL]"
     'to IX/IY, with an offset value e.g. "[IX+$FF]"
    PARAM_MEM_IX = 2 ^ 12
    PARAM_MEM_IY = 2 ^ 13
    [_PARAM_REGS_END] = PARAM_MEM_IY
    
    'A couple of undocumented instructions allow for IX/IY memory references,
     'but not the standard "[HL]" reference
    PARAM_MEM_IXY = PARAM_MEM_IX Or PARAM_MEM_IY
    'And this is the common "[HL|IX+$FF|IY+$FF]" form that is used often throughout
    PARAM_MEM_HLIXY = PARAM_MEM_HL Or PARAM_MEM_IXY
    
    'The main 8-bit registers are a common instruction parameter
    PARAM_REGS_ABCDEHL = PARAM_REG_A Or PARAM_REG_B Or PARAM_REG_C Or PARAM_REG_D Or PARAM_REG_E Or PARAM_REG_E Or PARAM_REG_H Or PARAM_REG_L
    'The Z80 clumps HL/IX & IY memory references together with 8-bit registers when
     'building opcodes, i.e. "A|B|C|D|E|H|L|[HL|IX+$FF|IY+$FF]"
    PARAM_REGS_ABCDEHL_MEM_HLIXY = PARAM_REGS_ABCDEHL Or PARAM_MEM_HLIXY
    'The use of the IX/IY prefix turns H/L into IXH/IXL/IYH/IYL in many instances
    PARAM_REGS_IXHL = PARAM_REG_IXH Or PARAM_REG_IXL
    PARAM_REGS_IYHL = PARAM_REG_IYH Or PARAM_REG_IYL
    PARAM_REGS_IXYHL = PARAM_REGS_IXHL Or PARAM_REGS_IYHL
    PARAM_REGS_ABCDEIXYHL_MEM_HLIXY = PARAM_REGS_ABCDEHL_MEM_HLIXY Or PARAM_REGS_IXYHL
    
    'Very uncommon 8-bit registers
    PARAM_REG_I = 2 ^ 14                'Interrupt register
    PARAM_REG_R = 2 ^ 15                'Refresh register, pseudo-random
    
    'The 16-bit register pairs
    PARAM_REG_AF = 2 ^ 16               'The Accumulator and the processor Flags
    PARAM_REG_BC = 2 ^ 17               'Registers B & C
    PARAM_REG_DE = 2 ^ 18               'Registers D & E
    PARAM_REG_HL = 2 ^ 19               'Registers H & L
    PARAM_REG_SP = 2 ^ 20               'Stack Pointer
    
    PARAM_REG_IX = 2 ^ 21               'Register IX
    PARAM_REG_IY = 2 ^ 22               'Register IY
    
    PARAM_VAL = 2 ^ 23                  'Immediate Value, e.g. `JP $FFFF`
    
    PARAM_REGS_BC_DE_SP = PARAM_REG_BC Or PARAM_REG_DE Or PARAM_REG_SP
    'Some instructions accept BC/DE/HL/SP, but not IX & IY due to existing prefixes
    PARAM_REGS_BC_DE_HL_SP = PARAM_REGS_BC_DE_SP Or PARAM_REG_HL
    
    'HL, IX & IY are synonymous as they use an opcode prefix to determine which
    PARAM_REGS_HL_IXY = PARAM_REG_HL Or PARAM_REG_IX Or PARAM_REG_IY
    'PUSH / POP allow AF but not SP
    PARAM_REGS_AF_BC_DE_HL_IXY = PARAM_REG_AF Or PARAM_REG_BC Or PARAM_REG_DE Or PARAM_REGS_HL_IXY
    'The LD instruction can take most 16-bit registers
    PARAM_REGS_BC_DE_HL_SP_IXY = PARAM_REGS_BC_DE_HL_SP Or PARAM_REG_IX Or PARAM_REG_IY
    
    '..................................................................................
    
    'Register C & Flag C cannot be distinguished by the tokeniser (it isn't aware of
     'context) so they are treated as the same thing. Another bit covers NC/Z/NZ so
     'that these are not accidentally taken as Register C elsewhere
    PARAM_FLAGS_CZ = PARAM_REG_C Or (2 ^ 24)
    PARAM_FLAGS_MP = (2 ^ 25)
    
    PARAM_FLAGS = PARAM_FLAGS_CZ Or PARAM_FLAGS_MP
    
    '..................................................................................
    
    'The IN and OUT instructions can use port "C" (which is, in reality, BC)
    PARAM_MEM_BC = 2 ^ 26
    PARAM_MEM_DE = 2 ^ 27
    PARAM_MEM_SP = 2 ^ 28
    
    PARAM_MEM_VAL = 2 ^ 29
    
    '..................................................................................
    
    'This is a shorthand to check for any instance of IX/IY so that we can add the
     'relevant opcode prefix with the simplest of tests
    PARAM_ANY_IX = PARAM_REG_IX Or PARAM_REG_IXH Or PARAM_REG_IXL Or PARAM_MEM_IX
    PARAM_ANY_IY = PARAM_REG_IY Or PARAM_REG_IYH Or PARAM_REG_IYL Or PARAM_MEM_IY
    PARAM_ANY_IXY = PARAM_ANY_IX Or PARAM_ANY_IY
    
    'This is shorthand to check for any 8/16-bit parameter Value
    PARAM_ANY_VAL = PARAM_VAL Or PARAM_MEM_VAL
End Enum

'A Z80 instruction can have up to 3 parameters; _
 other features are a main Value and an Offset (for indexed lookups, like `IX+$FF`)

Private Type oz80Instruction
    Param1 As oz80Param                 'The first parameter, if present
    Param2 As oz80Param                 'The second parameter, if present
    Param3 As oz80Param                 'The third parameter, if present (undocumented)
    
    'So that we don't need to check Param 1 & 2 several times, _
     we extract the Value from the relevant one (if present)
    Value As Double
    
    HasValue As Long                    'If the instruction implies a Value
    HasOffset As Long                   'If the instruction implies an Offset
End Type

'This is unused at this moment, but represents a further attempt to fit all _
 combinations of the three parameters in a single bit-mask:

Private Enum OZ80_MASK
    'Param 1 Bits
    '..................................................................................
    MASK_P1_A = (2 ^ 0)                 'Param 1 contains `A`
    MASK_P1_BDEHL = (2 ^ 1)             'Param 1 contains `B|D|E|H|L`
    MASK_P1_C = (2 ^ 2)                 'Param 1 contains `C` (Register or Flag)
    MASK_P1_IXYHL = (2 ^ 3)             'Param 1 contains `IXH|IYH|IXL|IYL`
    MASK_P1_IR = (2 ^ 4)                'Param 1 contains `I|R`
    MASK_P1_VAL = (2 ^ 5)               'Param 1 contains `$FF|$FFFF` (8/16-bit number)
    MASK_P1_AF = (2 ^ 6)                'Param 1 contains `AF`
    MASK_P1_BC = (2 ^ 7)                'Param 1 contains `BC`
    MASK_P1_DE = (2 ^ 8)                'Param 1 contains `DE
    MASK_P1_HL = (2 ^ 9)                'Param 1 contains `HL`
    MASK_P1_IXY = (2 ^ 10)              'Param 1 contains `IX` or `IY`
    MASK_P1_SP = (2 ^ 11)               'Param 1 contains `SP` (also `[SP]`)
    MASK_P1_MEM_VAL = (2 ^ 12)          'Param 1 contains `[$FF|$FFFF]`
    MASK_P1_MEM_BC = (2 ^ 13)           'Param 1 contains `[BC]` (also `[C]`)
    MASK_P1_MEM_DE = (2 ^ 14)           'Param 1 contains `[DE]`
    MASK_P1_MEM_HLIXY = (2 ^ 15)        'Param 1 contains `[HL|IX|IY]`, e.g. `jp`
    MASK_P1_NZZNCPOPEPM = (2 ^ 16)      'Param 1 contains `NZ|Z|NC|PE|PO|P|M` (flags)
    
    'Param 2 Bits
    '..................................................................................
    MASK_P2_A = (2 ^ 30)                'Param 2 contains `A`
    MASK_P2_BCDEHL = (2 ^ 29)           'Param 2 contains `B|C|D|E|H|L`
    MASK_P2_IXYHL = (2 ^ 28)            'Param 2 contains `IXH|IYH|IXL|IYL`
    MASK_P2_IR = (2 ^ 27)               'Param 2 contains `I|R`
    MASK_P2_VAL = (2 ^ 26)              'Param 2 contains `$FF|$FFFF`
    MASK_P2_AF = (2 ^ 25)               'Param 2 contains `AF`
    MASK_P2_BC_DE_SP = (2 ^ 24)         'Param 2 contains `BC|DE|SP`
    MASK_P2_HL = (2 ^ 23)               'Param 2 contains `HL`
    MASK_P2_IX = (2 ^ 22)               'Param 2 contains `IX`
    MASK_P2_IY = (2 ^ 21)               'Param 2 contains `IY`
    MASK_P2_MEM_BC_DE = (2 ^ 20)        'Param 2 contains `[BC]` (also `[C]`) or `[DE]`
    MASK_P2_MEM_HL = (2 ^ 19)           'Param 2 contains `[HL]`
    MASK_P2_MEM_IXY = (2 ^ 18)          'Param 2 contains `[IX+$FF|IY+$FF]`
    MASK_P2_MEM_VAL = (2 ^ 17)          'Param 2 contains `[$FF|$FFFF]
    
    'Param 3 (undocumented) Bits
    '..................................................................................
    MASK_P3_ABCDEHL = -1                'Param 3 contains `A|B|C|D|E|H|L`
    
    'Parameter Tests
    '..................................................................................
    TEST_P1_A = MASK_P1_A
    TEST_P1_ABCDEHL = TEST_P1_A Or MASK_P1_BDEHL Or MASK_P1_C
    TEST_P1_ABCDEHL_IXYHL_MEM_HLIXY = TEST_P1_ABCDEHL Or MASK_P1_IXYHL Or MASK_P1_MEM_HLIXY
    TEST_P1_IR = MASK_P1_IR
    
    TEST_P1_AF = MASK_P1_AF
    TEST_P1_DE = MASK_P1_DE
    [_TEST_P1_BC_DE] = MASK_P1_BC Or TEST_P1_DE
    TEST_P1_SP = MASK_P1_SP
    TEST_P1_BC_DE_SP = [_TEST_P1_BC_DE] Or TEST_P1_SP
    
    TEST_P1_HL = MASK_P1_HL
    TEST_P1_IXY = MASK_P1_IXY
    [_TEST_P1_HL_IXY] = TEST_P1_HL Or TEST_P1_IXY
    TEST_P1_BC_DE_SP_HLIXY = TEST_P1_BC_DE_SP Or [_TEST_P1_HL_IXY]
    TEST_P1_AF_BC_DE_HLIXY = TEST_P1_AF Or [_TEST_P1_BC_DE] Or [_TEST_P1_HL_IXY]
    
    TEST_P1_VAL = MASK_P1_VAL
    
    TEST_P1_MEM_VAL = MASK_P1_MEM_VAL
    TEST_P1_MEM_BC = MASK_P1_MEM_BC
    TEST_P1_MEM_DE = MASK_P1_MEM_DE
    TEST_P1_MEM_HLIXY = MASK_P1_MEM_HLIXY
    
    'Flags `NZ|Z|NC|C|PO|PE|P|M`
    TEST_P1_NZZNCC = MASK_P1_NZZNCPOPEPM Or MASK_P1_C
    
    TEST_P2_A = MASK_P2_A
    TEST_P2_ABCDEHL = TEST_P2_A Or MASK_P2_BCDEHL
    [_TEST_P2_IXYHL] = MASK_P2_IXYHL
    TEST_P2_ABCDEHL_IXYHL = TEST_P2_ABCDEHL Or [_TEST_P2_IXYHL]
    TEST_P2_IR = MASK_P2_IR
    
    TEST_P2_AF = MASK_P2_AF
    TEST_P2_BC_DE_SP = MASK_P2_BC_DE_SP
    TEST_P2_HL = MASK_P2_HL
    TEST_P2_BC_DE_SP_HL = TEST_P2_BC_DE_SP Or TEST_P2_HL
    TEST_P2_BC_DE_SP_IX = TEST_P2_BC_DE_SP Or MASK_P2_IX
    TEST_P2_BC_DE_SP_IY = TEST_P2_BC_DE_SP Or MASK_P2_IY
    TEST_P2_HLIXY = TEST_P2_HL Or MASK_P2_IX Or MASK_P2_IY
    
    TEST_P2_MEM_IXY = MASK_P2_MEM_IXY
    [_TEST_P2_MEM_HL_IXY] = TEST_P2_MEM_IXY Or MASK_P2_HL
    TEST_P2_ABCDEHL_IXYHL_MEM_HLIXY = TEST_P2_ABCDEHL_IXYHL Or [_TEST_P2_MEM_HL_IXY]
    
    TEST_P2_VAL = MASK_P2_VAL
    TEST_P2_MEM_VAL = MASK_P2_MEM_VAL
    
    TEST_P2_MEM_BC_DE = MASK_P2_MEM_BC_DE
    
End Enum

'--------------------------------------------------------------------------------------
'Z80 Opcodes:
'--------------------------------------------------------------------------------------
'Z80 opcodes are made up of various bit patterns that can be broken down into the
'following layout: (this information from <z80.info/decoding.htm>)

'             .-+-+-+-+-+-+-+-.
' Bit Number: |7|6|5|4|3|2|1|0|
'             ¦-+-+-+-+-+-+-+-¦
'  Component: | X |  Y  |  Z  |
'             '---+---+-+-----'
'                 | P |Q|
'                 '---+-'

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
Private OpcodeY_RST(0 To &H38&) As Long
Private OpcodeY_Bit(0 To 7) As Long
'Sure this is a waste of memory, but it is fast as we don't have to convert _
 bit-powers to linear numbers for many instructions being assembled
Private OpcodeY_Register([_PARAM_REGS_BEGIN] To [_PARAM_REGS_END]) As Long
Private OpcodeZ_Register([_PARAM_REGS_BEGIN] To [_PARAM_REGS_END]) As Long

'--------------------------------------------------------------------------------------
'ROM Layout:
'--------------------------------------------------------------------------------------
'As PROCs and TABLEs are parsed they are assembled into binary "Chunks". After all _
 source code is assembled, the Chunks are arranged according to their attributes and _
 placed into the output ROM file

Private Type oz80Chunk
    FullName As String                  'Fully-qualified name, i.e. "::section:label"
    LabelName As String                 'Quick-access to the Label name
    SectionName As String               'Quick-access to the Chunk's Section name
    Slot As OZ80_SLOT                   'Slot pattern assigned to the Chunk
    
    SectionID As Long                   'Index number to `Sections` array
    
    Data() As Byte                      'The binary content of the Chunk
    Buffer As Long                      'Buffer size for the data (power of 2)
    Size As Long                        'Length of the actual data in bytes
    
    AddrReal As Long                    'Location of Chunk within the ROM
    AddrSlot As Long                    'As above, but memory-mapped within 64 K
    
    TableIndex As Long                  'Current `INDEX` number in the Chunk
    'TODO: Replace this with something much, much faster
    Properties As Scripting.Dictionary
    
    LogASM As bluArrayStrings           'Disassembly log
End Type
'We'll erase a Chunk by setting it to this empty instance
Private EmptyChunk As oz80Chunk

Private Chunks() As oz80Chunk           'All Chunks are collected here
Private ChunkCount As Long              '1-based count of Chunks
Private ChunksSize As Long              'Total size of all Chunks, in bytes

'These will be references to the special Chunks at fixed positions within the ROM _
 (when laying out the ROM, we'll need to position these first)
Private Chunk_Header As Long            'The System-defined header
Private Chunk_Interrupt00 As Long       'Z80 Interrupt located at $0000
Private Chunk_Interrupt08 As Long       'Z80 intettupt located at $0008
Private Chunk_Interrupt10 As Long       'Z80 intettupt located at $0010
Private Chunk_Interrupt18 As Long       'Z80 intettupt located at $0018
Private Chunk_Interrupt20 As Long       'Z80 intettupt located at $0020
Private Chunk_Interrupt28 As Long       'Z80 intettupt located at $0028
Private Chunk_Interrupt30 As Long       'Z80 intettupt located at $0030
Private Chunk_Interrupt38 As Long       'Z80 IRQ (located at $0038)
Private Chunk_Interrupt66 As Long       'Z80 NMI (located at $0066)
Private Chunk_Start As Long             'Boot vector (often $0000)

'--------------------------------------------------------------------------------------

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

Private ROM() As Byte                   'The ROM's binary contents
Private ROMSize As Long                 'The ROM file size, e.g. 256 KB (power of 2)
Private ROMUsed As Long                 'Where the current end of the data is

'--------------------------------------------------------------------------------------
'Deferred Expressions:
'--------------------------------------------------------------------------------------
'Certain Values cannot be known until the size of all Chunks has been determined first. _
 When such a Value is encountered, the Expression is 'deferred' until after ROM Layout

Private Type oz80Deferred
    File As Long                        'File index number
    TokenIndex As Long                  'Token index number, within said File
    
    ChunkID As Long                     'Which Chunk to write the Value to
    ROMOffset As Long                   'Which starting byte within the ROM
    
    Kind As OZ80_DEFERRED               'Which type of data to output
    From As Long                        'The address to calculate relative from
End Type

Private Enum OZ80_DEFERRED
    ASM_VAL8                            'An 8-bit value
    ASM_VAL16                           'A 16-bit value
    ASM_JR                              'A relative jump (8-bits)
    ASM_JP                              'An absolute jump (16-bits)
    ASM_RST                             'An `rst` Z80 instruction
End Enum

Private Deferred() As oz80Deferred
Private DeferredCount As Long

'--------------------------------------------------------------------------------------
'Logging:
'--------------------------------------------------------------------------------------
'OZ80's internal dialogue will be dumped out to a text file as it goes, _
 this is not kept in memory
Private Log As bluLog

'Logging level
Public Enum OZ80_LOG
    OZ80_LOG_ERROR                      'Error messages have the highest priority
    OZ80_LOG_ACTION                     'The key important happenings
    OZ80_LOG_INFO                       'Optional information, not actions happening
    OZ80_LOG_STATUS                     'Display variable values &c. when assigned
    OZ80_LOG_DEBUG                      'Internal information for debugging purposes
End Enum

'For logging, we will want to get a text representation of any of the Tokens _
 (oh how I wish VB6 supported static constant arrays)
Private TokenNames(0 To OZ80_TOKEN.[_TOKEN_LAST] - 1) As String

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
Private Const STR_CONST                 As String = "CONST"
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
Private Const STR_DI                    As String = "di"
Private Const STR_DJNZ                  As String = "djnz"
Private Const STR_E                     As String = "E"
Private Const STR_EI                    As String = "ei"
Private Const STR_EX                    As String = "ex"
Private Const STR_EXX                   As String = "exx"
Private Const STR_F                     As String = "F"
Private Const STR_HLT                   As String = "hlt"
Private Const STR_HALT                  As String = "halt"
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
Private Const STR_INDEX_LABEL           As String = ".index"
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
Private Const STR_PE                    As String = "PE"
Private Const STR_PO                    As String = "PO"
Private Const STR_POP                   As String = "pop"
Private Const STR_PROC                  As String = "PROC"
Private Const STR_PUSH                  As String = "push"
Private Const STR_R                     As String = "R"
Private Const STR_RAM                   As String = "RAM"
Private Const STR_RAMPAGE               As String = "RAMPAGE"
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
Private Const STR_TYPE                  As String = "TYPE"
Private Const STR_WORD                  As String = "WORD"
Private Const STR_XOR                   As String = "xor"
Private Const STR_Z                     As String = "Z"

'Formatting of output messages
Private Const LOG_VERSION               As String = "OZ80MANDIAS v{$}.{$},{$}"
Private Const LOG_PHASE1                As String = "Phase 1: Structural Pass"
Private Const LOG_PHASE2                As String = "Phase 2: Layout ROM"
Private Const LOG_PHASE3                As String = "Phase 3: Defered Calculation"
Private Const LOG_PHASE4                As String = "Phase 4: Post Processing"
Private Const LOG_ASM_COMPLETE          As String = "Assembly Complete"
Private Const LOG_TIME                  As String = "Time: {$}s"
Private Const LOG_SECTION               As String = "SECTION {$}"
Private Const LOG_LINEDOUBLE            As String = "========================================================================="
Private Const LOG_FILE_LENGTH           As String = "Length: {$} Characters"
Private Const LOG_FILE_TOKENS           As String = "Tokens: {$}"

'For diaplying a disassembly
Private Const LOG_ASM_COLS              As String = "ADDR | PX OP OF CB LO HI | MNEMONIC  PARAM1    PARAM2    PARAM3"
Private Const LOG_ASM_HEAD              As String = "-----¦-------------------¦-----------------------------------------------"
Private Const LOG_ASM                   As String = "···· | ·· ·· ·· ·· ·· ·· | ·········                                     "
Private Const LOG_ASM_FOOT              As String = "-----+-------------------'-----------------------------.-----------------"
Private Const LOG_DATA_COLS             As String = "ADDR | 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F | 0123456789ABCDEF"
Private Const LOG_DATA_HEAD             As String = "-----¦-------------------------------------------------¦-----------------"
Private Const LOG_DATA                  As String = "···· | ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· ·· | ················"

'For displaying a Token's data
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
Private Const FORMAT_TOKEN_SQUARE       As String = "TYPE/MEM"
Private Const FORMAT_TOKEN_TEXT         As String = "    TEXT"
Private Const FORMAT_TOKEN_Z80          As String = "     Z80"

'To save having to create and destroy short strings all the time, _
 we'll keep this one around for formatting log messages
Private LogStr As bluString

'Which type of data the Chunk is recording (Z80 assembly or Data). _
 This is used to produce the disassembly log
Private ChunkMode As CHUNK_MODE
Private Enum CHUNK_MODE
    MODE_UNSET
    MODE_Z80                            'When adding Z80 instructions
    MODE_DATA                           'When adding data
End Enum

'When generating the disassembly log, we don't show data as 1-byte per line, _
 but render it in typical hex-bin format of 16 hex-bytes followed by ASCII. _
 We therefore need to prepare a line of 16-bytes before logging it
Private ChunkDataIndex As Long
'This will be the current line of disassembled log data (See `LOG_DATA`)
Private ChunkDataLog As String

'Lookup table of hexadecimal prettyprint, _
 saves slow & repetitive text manipulation when logging
Private HexStr8(-1 To &HFF&) As String * 2
Private HexStr16(0 To &HFFFF&) As String * 4

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
    
    'Add the whole lot up...
    [_PARAMTEXT_TOTAL] = PARAMTEXT_MEM_VAL16 + [_PARAMTEXT_MEM_VAL16]
End Enum

'...and reserve 10 characters of text for each
Private ParamText(0 To [_PARAMTEXT_TOTAL]) As String * 10

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
    'Unlike the assemblers of the past, we are not practically memory-constrained; _
     even though OZ80 would be hard-pushed to use more than a few hundred megs of _
     RAM, there is always virtual memory, so our work does not have a hard limit. _
     Therefore we aim to trade off memory-usage for speed with a large set of lookup _
     tables which we set up with the class. If you don't want to waste memory whilst _
     the assembler isn't running, only insantiate OZ80 when needed
    Dim i As Long
    
    'If too many possible tokens are created, things will screw up
    Debug.Assert OZ80_TOKEN.[_TOKEN_LAST] < 256
    Debug.Print "[OZ80] No. Tokens Defined: " & (OZ80_TOKEN.[_TOKEN_LAST] - 1)
    Debug.Print "[OZ80] Size of param text: " & [_PARAMTEXT_TOTAL] * 20 & " bytes"
    
    'Register the events we'll be timing with the profiler
    With Profiler
        Call .RegisterEvent(EVENT_INITIALIZE, "Initialize")
        Call .RegisterEvent(EVENT_FILELOAD_API, "bluString.FileLoad - api_CreateFile")
        Call .RegisterEvent(EVENT_TOKENISE, "FileTokenise")
        Call .RegisterEvent(EVENT_PARSE_WORD, "FileParse:ReadWord")
        Call .RegisterEvent(EVENT_BUILDTREE, "BuildSyntaxTree")
        Call .RegisterEvent(EVENT_FORMATTOKEN, "FormatToken")
        Call .RegisterEvent(EVENT_ASSEMBLE, "Assemble")
        Call .RegisterEvent(EVENT_Z80ASSEMBLE, "Z80_Assemble")
        Call .RegisterEvent(EVENT_ROM_LAYOUT, "ROM_Layout")
        Call .RegisterEvent(EVENT_ROM_ADDCHUNK, "ROM_AddChunk")
    End With
    
    '*** TIMER BEGIN ***
    'Begin using our profiler to time all this caching and initialising
    Call Profiler.EventBegin(EVENT_INITIALIZE)
    
    'This is a re-usable temporary String for logging
    Set LogStr = New bluString
    
    'Manually create a 64-bit Double "Infinite" number, which cannot be written _
     implicitly in VB6. This method is based upon <stackoverflow.com/a/896292>. _
     We use this to mark numbers as "not yet defined" whilst still allowing full _
     64-Bit floating-point calculations
    Dim Pointer As Long
    Let Pointer = VarPtr(INDEFINITE)
    Let W32TLB.MemByte(Pointer + 6) = 240
    Let W32TLB.MemByte(Pointer + 7) = 127
    'Manually create the largest possible 64-bit number
    Let Pointer = VarPtr(NUMBER_MAX)
    Let W32TLB.MemDouble(Pointer) = -1
    Let W32TLB.MemByte(Pointer + 6) = 239
    Let W32TLB.MemByte(Pointer + 7) = 127
    'Manually create the largest negative 64-bit number
    Let Pointer = VarPtr(NUMBER_MIN)
    Let W32TLB.MemDouble(Pointer) = -1
    Let W32TLB.MemByte(Pointer + 6) = 239
    Let W32TLB.MemByte(Pointer + 7) = 255
    'Define min/max numbers that can be multiplied by 1000/1024 without overflow
    Let NUMBER_MAX_K = NUMBER_MAX / 1000
    Let NUMBER_MIN_K = -NUMBER_MAX_K
    Let NUMBER_MAX_KB = NUMBER_MAX / 1024
    Let NUMBER_MIN_KB = -NUMBER_MAX_KB
    
    'Hexadecimal ASCII to decimal lookup table:
    '----------------------------------------------------------------------------------
    For i = 0 To 255: Let Hex2Dec(i) = -1: Next i
    Let Hex2Dec(UCS2.DIGIT_ZERO) = 0
    Let Hex2Dec(UCS2.DIGIT_ONE) = 1
    Let Hex2Dec(UCS2.DIGIT_TWO) = 2
    Let Hex2Dec(UCS2.DIGIT_THREE) = 3
    Let Hex2Dec(UCS2.DIGIT_FOUR) = 4
    Let Hex2Dec(UCS2.DIGIT_FIVE) = 5
    Let Hex2Dec(UCS2.DIGIT_SIX) = 6
    Let Hex2Dec(UCS2.DIGIT_SEVEN) = 7
    Let Hex2Dec(UCS2.DIGIT_EIGHT) = 8
    Let Hex2Dec(UCS2.DIGIT_NINE) = 9
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_A) = 10
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_A) = 10
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_B) = 11
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_B) = 11
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_C) = 12
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_C) = 12
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_D) = 13
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_D) = 13
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_E) = 14
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_E) = 14
    Let Hex2Dec(UCS2.LATIN_SMALL_LETTER_F) = 15
    Let Hex2Dec(UCS2.LATIN_CAPITAL_LETTER_F) = 15
    
    '8 & 16-bit hexadecimal numbers:
    '----------------------------------------------------------------------------------
    'TODO: We need to make this stupidly fast...
    Let HexStr8(-1) = "??"
    For i = 0 To &HF&:          Let HexStr8(i) = "0" & Hex$(i):     Next i
    For i = &H10& To &HFF&:     Let HexStr8(i) = Hex$(i):           Next i
    
    For i = 0 To &HF&:          Let HexStr16(i) = "000" & Hex$(i):  Next i
    For i = &H10& To &HFF&:     Let HexStr16(i) = "00" & Hex$(i):   Next i
    For i = &H100& To &HFFF&:   Let HexStr16(i) = "0" & Hex$(i):    Next i
    For i = &H1000& To &HFFFF&: Let HexStr16(i) = Hex$(i):          Next i
    
    'Token Names:
    '----------------------------------------------------------------------------------
    Let i = 0
    'Z80 Instructions .................................................................
    Let i = i + 1: Let TokenNames(TOKEN_Z80_ADC) = STR_ADC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_ADD) = STR_ADD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_AND) = STR_AND
    Let i = i + 1: Let TokenNames(TOKEN_Z80_BIT) = STR_BIT
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CALL) = STR_CALL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CCF) = STR_CCF
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CP) = STR_CP
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CPD) = STR_CPD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CPDR) = STR_CPDR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CPI) = STR_CPI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CPIR) = STR_CPIR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_CPL) = STR_CPL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_DAA) = STR_DAA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_DEC) = STR_DEC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_DI) = STR_DI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_DJNZ) = STR_DJNZ
    Let i = i + 1: Let TokenNames(TOKEN_Z80_EI) = STR_EI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_EX) = STR_EX
    Let i = i + 1: Let TokenNames(TOKEN_Z80_EXX) = STR_EXX
    Let i = i + 1: Let TokenNames(TOKEN_Z80_HALT) = STR_HALT
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IM) = STR_IM
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IN) = STR_IN
    Let i = i + 1: Let TokenNames(TOKEN_Z80_INC) = STR_INC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IND) = STR_IND
    Let i = i + 1: Let TokenNames(TOKEN_Z80_INDR) = STR_INDR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_INI) = STR_INI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_INIR) = STR_INIR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_JP) = STR_JP
    Let i = i + 1: Let TokenNames(TOKEN_Z80_JR) = STR_JR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_LD) = STR_LD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_LDD) = STR_LDD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_LDDR) = STR_LDDR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_LDI) = STR_LDI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_LDIR) = STR_LDIR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_NEG) = STR_NEG
    Let i = i + 1: Let TokenNames(TOKEN_Z80_NOP) = STR_NOP
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OR) = STR_OR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OUT) = STR_OUT
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OUTD) = STR_OUTD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OTDR) = STR_OTDR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OUTI) = STR_OUTI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_OTIR) = STR_OTIR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_POP) = STR_POP
    Let i = i + 1: Let TokenNames(TOKEN_Z80_PUSH) = STR_PUSH
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RES) = STR_RES
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RET) = STR_RET
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RETI) = STR_RETI
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RETN) = STR_RETN
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RLA) = STR_RLA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RL) = STR_RL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RLC) = STR_RLC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RLCA) = STR_RLCA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RLD) = STR_RLD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RR) = STR_RR
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RRA) = STR_RRA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RRC) = STR_RRC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RRCA) = STR_RRCA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RRD) = STR_RRD
    Let i = i + 1: Let TokenNames(TOKEN_Z80_RST) = STR_RST
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SBC) = STR_SBC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SCF) = STR_SCF
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SET) = STR_SET
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SLA) = STR_SLA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SLL) = STR_SLL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SRA) = STR_SRA
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SRL) = STR_SRL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SUB) = STR_SUB
    Let i = i + 1: Let TokenNames(TOKEN_Z80_XOR) = STR_XOR
    'Z80 Registers / Flags ............................................................
    Let i = i + 1: Let TokenNames(TOKEN_Z80_A) = STR_A
    Let i = i + 1: Let TokenNames(TOKEN_Z80_AF) = STR_AF
    Let i = i + 1: Let TokenNames(TOKEN_Z80_B) = STR_B
    Let i = i + 1: Let TokenNames(TOKEN_Z80_C) = STR_C
    Let i = i + 1: Let TokenNames(TOKEN_Z80_BC) = STR_BC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_D) = STR_D
    Let i = i + 1: Let TokenNames(TOKEN_Z80_E) = STR_E
    Let i = i + 1: Let TokenNames(TOKEN_Z80_H) = STR_H
    Let i = i + 1: Let TokenNames(TOKEN_Z80_DE) = STR_DE
    Let i = i + 1: Let TokenNames(TOKEN_Z80_H) = STR_H
    Let i = i + 1: Let TokenNames(TOKEN_Z80_HL) = STR_HL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_I) = STR_I
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IX) = STR_IX
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IXH) = STR_IXH
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IXL) = STR_IXL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IY) = STR_IY
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IYH) = STR_IYH
    Let i = i + 1: Let TokenNames(TOKEN_Z80_IYL) = STR_IYL
    Let i = i + 1: Let TokenNames(TOKEN_Z80_L) = STR_L
    Let i = i + 1: Let TokenNames(TOKEN_Z80_M) = STR_M
    Let i = i + 1: Let TokenNames(TOKEN_Z80_NC) = STR_NC
    Let i = i + 1: Let TokenNames(TOKEN_Z80_NZ) = STR_NZ
    Let i = i + 1: Let TokenNames(TOKEN_Z80_P) = STR_P
    Let i = i + 1: Let TokenNames(TOKEN_Z80_PE) = STR_PE
    Let i = i + 1: Let TokenNames(TOKEN_Z80_PO) = STR_PO
    Let i = i + 1: Let TokenNames(TOKEN_Z80_R) = STR_R
    Let i = i + 1: Let TokenNames(TOKEN_Z80_SP) = STR_SP
    Let i = i + 1: Let TokenNames(TOKEN_Z80_Z) = STR_Z
    'Operators ........................................................................
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_ADD) = ChrW$(SYNTAX_OPERATOR_ADD)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_SUB) = ChrW$(SYNTAX_OPERATOR_SUB)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_MUL) = ChrW$(SYNTAX_OPERATOR_MUL)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_DIV) = ChrW$(SYNTAX_OPERATOR_DIV)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_POW) = ChrW$(SYNTAX_OPERATOR_POW)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_MOD) = ChrW$(SYNTAX_OPERATOR_MOD)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_REP) = "x"
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_OR) = ChrW$(SYNTAX_OPERATOR_OR)
    Let i = i + 1: Let TokenNames(TOKEN_OPERATOR_AND) = ChrW$(SYNTAX_OPERATOR_AND)
    'Keywords .........................................................................
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_BOOL) = STR_BOOL
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_BYTE) = STR_BYTE
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_CONST) = STR_CONST
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_HELP) = STR_HELP
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_HI) = STR_HI
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_INCLUDE) = STR_INCLUDE
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_INDEX) = STR_INDEX
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_INTERRUPT) = STR_INTERRUPT
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_LO) = STR_LO
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_LONG) = STR_LONG
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_NYBL) = STR_NYBL
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_PARAMS) = STR_PARAMS
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_PROC) = STR_PROC
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_RAM) = STR_RAM
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_RAMPAGE) = STR_RAMPAGE
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_RETURN) = STR_RETURN
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_SECTION) = STR_SECTION
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_SLOT) = STR_SLOT
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_START) = STR_START
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_SYSTEM) = STR_SYSTEM
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_TABLE) = STR_TABLE
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_TRIP) = STR_TRIP
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_TYPE) = STR_TYPE
    Let i = i + 1: Let TokenNames(TOKEN_KEYWORD_WORD) = STR_WORD
    '
    Let i = i + 1: Let TokenNames(TOKEN_NUMBER) = vbNullString
    Let i = i + 1: Let TokenNames(TOKEN_SUFFIX_K) = STR_K
    Let i = i + 1: Let TokenNames(TOKEN_SUFFIX_KB) = STR_KB
    '
    Let i = i + 1: Let TokenNames(TOKEN_BRACES_OPEN) = ChrW$(SYNTAX_BRACES_OPEN)
    Let i = i + 1: Let TokenNames(TOKEN_BRACES_CLOSE) = ChrW$(SYNTAX_BRACES_CLOSE)
    Let i = i + 1: Let TokenNames(TOKEN_PARENS_OPEN) = ChrW$(SYNTAX_PARENS_OPEN)
    Let i = i + 1: Let TokenNames(TOKEN_PARENS_CLOSE) = ChrW$(SYNTAX_PARENS_CLOSE)
    Let i = i + 1: Let TokenNames(TOKEN_SQUARE_OPEN) = ChrW$(SYNTAX_SQUARE_OPEN)
    Let i = i + 1: Let TokenNames(TOKEN_SQUARE_CLOSE) = ChrW$(SYNTAX_SQUARE_CLOSE)
    '
    Let i = i + 1: Let TokenNames(TOKEN_NAME) = vbNullString
    Let i = i + 1: Let TokenNames(TOKEN_PROPERTY_NEW) = ChrW$(SYNTAX_PROPERTY)
    Let i = i + 1: Let TokenNames(TOKEN_PROPERTY_ANON) = ChrW$(SYNTAX_PROPERTY)
    Let i = i + 1: Let TokenNames(TOKEN_TEXT) = ChrW$(SYNTAX_TEXT)
    
    Let i = i + 1: Let TokenNames(TOKEN_DOC) = "``"
    
    'The token names list needs to be kept up to date, _
     so if we've missed one I can alert you here
    Debug.Assert i = [_TOKEN_LAST]
    
    'Token Attributes:
    '----------------------------------------------------------------------------------
    Let i = 0
    Let i = i + 1: Let TokenAttributes(TOKEN_NONE) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_ADC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_ADD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_AND) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_BIT) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CALL) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CCF) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CP) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CPD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CPDR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CPI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CPIR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_CPL) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_DAA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_DEC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_DI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_DJNZ) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_EI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_EX) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_EXX) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_HALT) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IM) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IN) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_INC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IND) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_INDR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_INI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_INIR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_JP) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_JR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_LD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_LDD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_LDDR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_LDI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_LDIR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_NEG) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_NOP) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OUT) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OUTD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OTDR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OUTI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_OTIR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_POP) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_PUSH) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RES) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RET) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RETI) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RETN) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RLA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RL) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RLCA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RLC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RLD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RRA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RR) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RRCA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RRC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RRD) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_RST) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SBC) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SCF) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SET) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SLA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SRA) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SLL) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SRL) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SUB) = ATTR_Z80_MNEMONIC
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_XOR) = ATTR_Z80_MNEMONIC
    '
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_A) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_AF) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_B) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_BC) = ATTR_Z80_REGISTER Or ATTR_PARAM
    '`C` can be both a Register or a Flag
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_C) = ATTR_Z80_REGFLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_D) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_DE) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_E) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_H) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_HL) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_I) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IX) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IXH) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IXL) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IY) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IYH) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_IYL) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_L) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_M) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_NC) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_NZ) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_P) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_PE) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_PO) = ATTR_Z80_FLAG Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_R) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_SP) = ATTR_Z80_REGISTER Or ATTR_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_Z80_Z) = ATTR_Z80_FLAG Or ATTR_PARAM
    '
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_ADD) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_SUB) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_MUL) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_DIV) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_POW) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_MOD) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_REP) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_OR) = ATTR_OPERTATOR
    Let i = i + 1: Let TokenAttributes(TOKEN_OPERATOR_AND) = ATTR_OPERTATOR
    '
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_BOOL) = ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_BYTE) = ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_CONST) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_HELP) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_HI) = ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_INCLUDE) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_INDEX) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_INTERRUPT) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_LO) = ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_LONG) = ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_NYBL) = ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_PARAMS) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_PROC) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_RAM) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_RAMPAGE) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_RETURN) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_SECTION) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_SLOT) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_START) = ATTR_KEYWORD Or ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_SYSTEM) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_TABLE) = ATTR_KEYWORD
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_TRIP) = ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_TYPE) = ATTR_KEYWORD Or ATTR_UNIT
    Let i = i + 1: Let TokenAttributes(TOKEN_KEYWORD_WORD) = ATTR_UNIT
    '
    Let i = i + 1: Let TokenAttributes(TOKEN_NUMBER) = ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_SUFFIX_K) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_SUFFIX_KB) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_BRACES_OPEN) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_BRACES_CLOSE) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_SQUARE_OPEN) = ATTR_Z80_PARAM
    Let i = i + 1: Let TokenAttributes(TOKEN_SQUARE_CLOSE) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_PARENS_OPEN) = ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_PARENS_CLOSE) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_NAME) = ATTR_LISTEXPR
    Let i = i + 1: Let TokenAttributes(TOKEN_PROPERTY_NEW) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_PROPERTY_ANON) = ATTR_NONE
    Let i = i + 1: Let TokenAttributes(TOKEN_TEXT) = ATTR_LIST
    Let i = i + 1: Let TokenAttributes(TOKEN_DOC) = ATTR_NONE
    'The token attributes list needs to be kept up to date, _
     so if we've missed one I can alert you here
    Debug.Assert i = [_TOKEN_LAST]
    
    'Z80 Param Text:
    '----------------------------------------------------------------------------------
    'Index 0 acts as an "unimplemented" marker so we can see when certain cache _
     indexes haven't been implemented yet
    Let ParamText(PARAMTEXT_NONE) = "!!!!!!!!!!"
    
    For i = 0 To ([_TOKEN_Z80_END] - [_TOKEN_Z80_BEGIN])
        Let ParamText(PARAMTEXT_Z80 + i) = TokenNames([_TOKEN_Z80_BEGIN] + i)
    Next i
    
    For i = 0 To ([_TOKEN_REGS_END] - [_TOKEN_REGS_BEGIN])
        Let ParamText(PARAMTEXT_REGS + i) = TokenNames([_TOKEN_REGS_BEGIN] + i)
    Next i
    
    For i = 0 To [_PARAMTEXT_VAL8]
        Let ParamText(PARAMTEXT_VAL8 + i) = _
            ChrW$(SYNTAX_NUMBER_HEX) & HexStr8(i)
    Next i
    
    For i = 0 To [_PARAMTEXT_VAL16]
        Let ParamText(PARAMTEXT_VAL16 + i) = _
            ChrW$(SYNTAX_NUMBER_HEX) & HexStr16(i)
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
            HexStr8(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    'Register IY memory reference, with value -- i.e. "[IY+$FF]"
    For i = 0 To [_PARAMTEXT_MEM_IY_VAL]
        Let ParamText(PARAMTEXT_MEM_IY_VAL + i) = _
            ChrW$(SYNTAX_SQUARE_OPEN) & STR_IY & ChrW$(SYNTAX_OPERATOR_ADD) & _
            HexStr8(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    
    '16-bit memory references, i.e "[$FFFF]"
    For i = 0 To [_PARAMTEXT_MEM_VAL16]
        Let ParamText(PARAMTEXT_MEM_VAL16 + i) = _
            ChrW$(SYNTAX_SQUARE_OPEN) & ChrW$(SYNTAX_NUMBER_HEX) & _
            HexStr16(i) & ChrW$(SYNTAX_SQUARE_CLOSE)
    Next i
    
    '----------------------------------------------------------------------------------
    
    'Populate the look-up table for converting register tokens into their bit masks
    Let TokenRegister(TOKEN_Z80_A) = PARAM_REG_A
    Let TokenRegister(TOKEN_Z80_B) = PARAM_REG_B
    Let TokenRegister(TOKEN_Z80_C) = PARAM_REG_C
    Let TokenRegister(TOKEN_Z80_D) = PARAM_REG_D
    Let TokenRegister(TOKEN_Z80_E) = PARAM_REG_E
    Let TokenRegister(TOKEN_Z80_H) = PARAM_REG_H
    Let TokenRegister(TOKEN_Z80_L) = PARAM_REG_L
    Let TokenRegister(TOKEN_Z80_I) = PARAM_REG_I
    Let TokenRegister(TOKEN_Z80_R) = PARAM_REG_R
    Let TokenRegister(TOKEN_Z80_AF) = PARAM_REG_AF
    Let TokenRegister(TOKEN_Z80_BC) = PARAM_REG_BC
    Let TokenRegister(TOKEN_Z80_DE) = PARAM_REG_DE
    Let TokenRegister(TOKEN_Z80_HL) = PARAM_REG_HL
    Let TokenRegister(TOKEN_Z80_SP) = PARAM_REG_SP
    Let TokenRegister(TOKEN_Z80_IX) = PARAM_REG_IX
    Let TokenRegister(TOKEN_Z80_IXH) = PARAM_REG_IXH
    Let TokenRegister(TOKEN_Z80_IXL) = PARAM_REG_IXL
    Let TokenRegister(TOKEN_Z80_IY) = PARAM_REG_IY
    Let TokenRegister(TOKEN_Z80_IYH) = PARAM_REG_IYH
    Let TokenRegister(TOKEN_Z80_IYL) = PARAM_REG_IYL
    Let TokenRegister(TOKEN_Z80_NC) = PARAM_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_Z) = PARAM_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_NZ) = PARAM_FLAGS_CZ
    Let TokenRegister(TOKEN_Z80_P) = PARAM_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PE) = PARAM_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_PO) = PARAM_FLAGS_MP
    Let TokenRegister(TOKEN_Z80_M) = PARAM_FLAGS_MP
    
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
    
    Let OpcodeY_Register(PARAM_REG_B) = Y0
    Let OpcodeY_Register(PARAM_REG_C) = Y1
    Let OpcodeY_Register(PARAM_REG_D) = Y2
    Let OpcodeY_Register(PARAM_REG_E) = Y3
    Let OpcodeY_Register(PARAM_REG_H) = Y4
    Let OpcodeY_Register(PARAM_REG_IXH) = Y4
    Let OpcodeY_Register(PARAM_REG_IYH) = Y4
    Let OpcodeY_Register(PARAM_REG_L) = Y5
    Let OpcodeY_Register(PARAM_REG_IXL) = Y5
    Let OpcodeY_Register(PARAM_REG_IYL) = Y5
    Let OpcodeY_Register(PARAM_MEM_HL) = Y6
    Let OpcodeY_Register(PARAM_MEM_IX) = Y6
    Let OpcodeY_Register(PARAM_MEM_IY) = Y6
    Let OpcodeY_Register(PARAM_REG_A) = Y7
    
    Let OpcodeY_RST(&H0) = Y0
    Let OpcodeY_RST(&H8) = Y1
    Let OpcodeY_RST(&H10) = Y2
    Let OpcodeY_RST(&H18) = Y3
    Let OpcodeY_RST(&H20) = Y4
    Let OpcodeY_RST(&H28) = Y5
    Let OpcodeY_RST(&H30) = Y6
    Let OpcodeY_RST(&H38) = Y7
    
    Let OpcodeZ_Register(PARAM_REG_B) = Z0
    Let OpcodeZ_Register(PARAM_REG_C) = Z1
    Let OpcodeZ_Register(PARAM_REG_D) = Z2
    Let OpcodeZ_Register(PARAM_REG_E) = Z3
    Let OpcodeZ_Register(PARAM_REG_H) = Z4
    Let OpcodeZ_Register(PARAM_REG_IXH) = Z4
    Let OpcodeZ_Register(PARAM_REG_IYH) = Z4
    Let OpcodeZ_Register(PARAM_REG_L) = Z5
    Let OpcodeZ_Register(PARAM_REG_IXL) = Z5
    Let OpcodeZ_Register(PARAM_REG_IYL) = Z5
    Let OpcodeZ_Register(PARAM_MEM_HL) = Z6
    Let OpcodeZ_Register(PARAM_MEM_IX) = Z6
    Let OpcodeZ_Register(PARAM_MEM_IY) = Z6
    Let OpcodeZ_Register(PARAM_REG_A) = Z7
        
    Call Profiler.EventEnd(EVENT_INITIALIZE)
    '*** TIMER END ***
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Call Reset
    
    'Free the look-up tables:
    Erase OpcodeP_RegisterPair, OpcodeY_Bit, OpcodeY_Condition, _
          OpcodeY_Register, OpcodeY_RST, OpcodeZ_Register, TokenRegister
    Erase Hex2Dec
    
    'Erase the caches
    Erase HexStr8, HexStr16
    Erase TokenNames, ParamText
    
    Set Profiler = Nothing
    Set LogStr = Nothing
End Sub

'Reset : Clears the internal state
'======================================================================================
Private Sub Reset()
    'Reset the error number
    Let Error = OZ80_ERROR_NONE
    
    'Stop the log (this will close the file handle, if open)
    Set Log = Nothing
    
    'Clear the database of source files
    Erase Files: Let File = 0
    
    'Clear the syntax tree (source code machine representation)
    Erase Tree.Nodes
    Let Tree.Count = 0
    Let Tree.Buffer = 0
    Let Tree.Index = 0
    
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
    Let ChunksSize = 0
    
    'Reset Sections
    Erase Sections: Let SectionCount = 0
    'Erase RAM Pages
    Erase RAMPages: Let RAMPagesCount = 0
    
    'Reset the name value datastore
    Erase Names: Let NameCount = 0
    
    'Erase the Types (which includes `RAM`)
    Erase Types
    Let TypeCount = 0
    
    'Clear ROM layout
    Erase Layout: Let LayoutCount = 0
    Erase ROM
    'TODO: Default ROM size should be per-system / user-selectable?
    Let ROMSize = 8192
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Abort : Stop the assembly when running (e.g. user cancelled)
'======================================================================================
Public Sub Abort()
    'OZ80 checks the error state regularly during assembly, _
     meaning that it will cancel at the next check
    Let Error = OZ80_ERROR_ABORT
End Sub

'Assemble : Compile a Z80 binary from a given starting OZ80 source code file
'======================================================================================
'SourceFile     | Path to Source code file to begin with
'[LogFile]      | Path to desired log file. Will default to SourceFile + ".log"
'---------------¦----------------------------------------------------------------------
'Returns        | The error number
'======================================================================================
Public Function Assemble( _
             ByVal SourceFile As String, _
    Optional ByRef LogFile As String = vbNullString _
) As OZ80_ERROR
    'This will be a crude overall timer for the assembly
    Dim Begin As Single
    Let Begin = Timer
    
    'Initialise the internal state
    Call Reset
    
    'Begin logging
    '----------------------------------------------------------------------------------
    'This is OZ80's write-only logger. Anything sent to this is appended to the log
    Set Log = New bluLog
    
    'If a log file is not provided, _
     we'll take the source file's name
    If LogFile = vbNullString Then
        'Replace the file-extension with ".log"
        Let LogFile = blu.FileSystem.PathChangeExtension(SourceFile, ".log")
    End If
    
    'Re-base the log file
    Let LogFile = AssembleFile_SelectPath(LogFile)
    
    If Not Log.SetFilePath(LogFile) Then
        'ERROR: The log file could not be used
        Stop
    End If
    
    'Announce ourselves
    Call RaiseMessage(OZ80_LOG_INFO, LOG_VERSION, App.Major, App.Minor, App.Revision)
    
    '----------------------------------------------------------------------------------
    'Phase 1: Parse & Assemble
    '----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE1)
    
    'User-defined Sections begin from index 1 and up. _
     Index 0 is reserved as a default for Chunks without a specified Section
    ReDim Sections(0 To 0) As oz80Section
    
    'Prepare the default RAM page, _
     until the System is defined, the default RAM page will have no set size
    ReDim RAMPages(0) As oz80RAMPage
    Let RAMPages(0).Type.Name = "$"
    
    'When building the abstract syntax tree, nodes are created ahead of time and _
     populated with the details once each word is recognised. This makes it easier _
     to pick the relationship (sibling or child) of nodes as we go. The one tree is _
     shared for all source code files, so we reset it here
    Let Tree.Count = 1
    Let Tree.Buffer = 100000
    ReDim Tree.Nodes(1 To Tree.Buffer) As oz80TreeNode
    
    Let Tree.Index = 1
    
    '----------------------------------------------------------------------------------
    
    'Start with parsing the given file; the assembler does not work with the original _
     text directly. Source code is first parsed into tokens and then executed, once _
     an `INCLUDE` statement occurs, the next file is parsed and executed in-situ
    Call AssembleFile(SourceFile)
    
    'If there was a problem parsing the file, an error message will _
     already have been raised by an event, we only need to exit
    If Error Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    'Phase 2: Layout the ROM _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE2)
    
    'Now that the size of each Chunk is known, we can arrange the Chunks in the ROM _
     space and determine each Label location
'    Call ROMLayout
    
    If Error Then GoTo Finish
    
'    '----------------------------------------------------------------------------------
'    'Phase 3: Calculate Deferred Expressions _
'     ----------------------------------------------------------------------------------
'    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE3)
'
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
''    Dim Key As Variant
''    For Each Key In Values.Keys
''        Call RaiseMessage(OZ80_LOG_INFO, Key & " = " & Hex(Values.Item(Key)))
''    Next
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
''    Call RaiseMessage(OZ80_LOG_INFO, vbNullString)
'
'    Dim i As Long
'    For i = 1 To DeferredCount: With Deferred(i)
'
''        Call RaiseMessage(OZ80_LOG_INFO, "File : {$}", Files(.File).Path)
''        Call RaiseMessage(OZ80_LOG_INFO, _
''            "Line : {$} Col : {$}", _
''            CStr(Files(.File).Tokens(.TokenIndex).Line), _
''            CStr(Files(.File).Tokens(.TokenIndex).Col) _
''        )
''        Call RaiseMessage(OZ80_LOG_INFO, Chunks(.ChunkID).LabelName)
'
'        'Point OZ80 at the Expression
'        Let File = .File: Let TokenIndex = .TokenIndex - 1
'        '(this will cache the Token data)
'        Call NextToken
'
'        'Now calculate the Expression
'        Dim Value As Double
'        Call ParseExpression(Value, .ChunkID)
'
'        If Error Then
'            'TODO: Unexpected; considering the Expression will already have _
'                   been walked when it was deferred!
'            Stop
'        End If
'
'        If Value = INDEFINITE Then
'            Dim IndefiniteCount
'            Let IndefiniteCount = IndefiniteCount + 1
'
'            Call RaiseMessage(OZ80_LOG_ERROR, "INDEFINITE!")
'            Debug.Print "File : " & Files(.File).Path
'            Debug.Print "Line : " & CStr(Files(.File).Stream.Tokens(.TokenIndex).Line) & " Col : " & (Files(.File).Stream.Tokens(.TokenIndex).Col)
'            Debug.Print Chunks(.ChunkID).LabelName
'            Debug.Print
'
'            Let TokenIndex = .TokenIndex - 1: Call NextToken
'            Call RaiseError(OZ80_ERROR_INDEFINITE, "INDEFINITE!")
'            GoTo Finish
'
''            Stop
'
'        Else
''            Call RaiseMessage(OZ80_LOG_STATUS, Hex(Value))
'            If .Kind = ASM_JR Then
'                Dim From As Long
'                Let From = Chunks(.ChunkID).AddrReal + .ROMOffset
'
'                'The Value given should be the absolute address of the destination
'                Dim Diff As Long
'                Let Diff = Value - (Chunks(.ChunkID).AddrReal + .From)
'
'                If Diff < 0 _
'                    Then Let ROM(From) = (255 + Diff) And &HFF& _
'                    Else Let ROM(From) = (Diff - 1) And &HFF&
'
'            ElseIf .Kind = ASM_JP Then
'                Let From = Chunks(.ChunkID).AddrReal + .ROMOffset
'
'                Let ROM(From) = Value And &HFF&
'                Let ROM(From + 1) = (Value And &HFF00&) \ &H100&
'
'            ElseIf .Kind = ASM_VAL8 Then
'                Let From = Chunks(.ChunkID).AddrReal + .ROMOffset
'                'TODO: 8-bit limit check?
'                Let ROM(From) = Value And &HFF&
'
'            ElseIf .Kind = ASM_VAL16 Then
'                Let From = Chunks(.ChunkID).AddrReal + .ROMOffset
'
'                Let ROM(From) = Value And &HFF&
'                Let ROM(From + 1) = (Value And &HFF00&) \ &H100&
'
'            ElseIf .Kind = ASM_RST Then
'                Let ROM( _
'                    Chunks(.ChunkID).AddrReal + .ROMOffset _
'                ) = X3 Or OpcodeY_RST(Value) Or Z7
'            Else
'                'TODO: Any other kind of deferred expressions
'                Stop
'            End If
'        End If
'
''        Call RaiseMessage(OZ80_LOG_INFO, vbNullString)
'    End With: Next i
'
''    Call RaiseMessage(OZ80_LOG_INFO, "Number of Deferred Expressions: " & DeferredCount)
''    Call RaiseMessage(OZ80_LOG_INFO, "Number of Indefinite Expressions: " & IndefiniteCount)
    
    '----------------------------------------------------------------------------------
    'Phase 4: Post Processing, e.g. Checksum _
     ----------------------------------------------------------------------------------
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_PHASE4)
    
    'TODO
    
    '-----------------------------------------------------------------------------------
    
    'Finally, write the ROM to disk
    Dim FileNumber As Integer: Let FileNumber = FreeFile
    Open App.Path & "\ROM.sms" For Binary Access Write As #FileNumber
    Put #FileNumber, , ROM
    Close #FileNumber
    
    Call RaiseMessage(OZ80_LOG_ACTION, LOG_ASM_COMPLETE)
    Call RaiseMessage(OZ80_LOG_INFO, LOG_TIME, Format$(Timer - Begin, "0.000"))
    
Finish:
    '----------------------------------------------------------------------------------
    'Return the error number
    Let Assemble = Error
    
    'Clean up (resets error number)
    Call Reset
End Function

'AssembleFile_SelectPath : Given a relative path, convert it to absolute
'======================================================================================
'FilePath       | The user-provided path, which could be relative or absolute
'---------------¦----------------------------------------------------------------------
'Returns        | A fixed, absolute path (relative from the previously parsed file)
'======================================================================================
Private Function AssembleFile_SelectPath( _
    ByRef FilePath As String _
) As String
    'If the file is already absolute, it does not need a base path
    If blu.FileSystem.PathIsAbsolute(FilePath) Then
        Let AssembleFile_SelectPath = FilePath
    Else
        'Path is relative! If no other file has been parsed yet _
         we will need to use the current directory as a base
        If File = 0 Then
            Let AssembleFile_SelectPath = blu.FileSystem.PathCombine( _
                blu.FileSystem.WorkingDirectory, FilePath _
            )
        Else
            'When another file has been parsed before, _
             we work relatively from that file's directory
            Let AssembleFile_SelectPath = blu.FileSystem.PathCombine( _
                blu.FileSystem.FileFolder(Files(File).Path), _
                FilePath _
            )
        End If
    End If
End Function

'AssembleFile : Parse a source file and then begin assembling it
'======================================================================================
'SourceFile     | File path to the source code file (".oz80")
'======================================================================================
Private Sub AssembleFile( _
    ByRef SourceFile As String _
)
    '[1] Normalise the file path: _
     ----------------------------------------------------------------------------------
    'The file path handed to us might be just a file name and/or contain a relative _
     path (i.e. "..\"). We need to normalise this given path into an absolute path _
     (e.g. "C:\Folder\File.ext")
    Dim FilePath As String
    Let FilePath = AssembleFile_SelectPath(SourceFile)
    
    '[2] Check file exists: _
     ----------------------------------------------------------------------------------
    If Not blu.FileSystem.FileExists(FilePath) Then Call RaiseError( _
        OZ80_ERROR_FILE_NOTFOUND, 0, 0, _
        "The following file could not be found:" & vbNewLine & vbNewLine & FilePath _
    ): Exit Sub
    
    'TODO: Cannot assemble two of the same file (i.e. cannot `INCLUDE` a file _
     that's already been included)
    
    '[3] Tokenise the file: _
     ----------------------------------------------------------------------------------
    'Check to see if the file has already been tokenised
    Dim i As Long
    For i = 1 To FileCount
        If Files(i).Path = FilePath Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE, 0, 0, _
            "File already assembled! You cannot `INCLUDE` a file that has already " _
          & "been included elsewhere" _
        ): Exit Sub
    Next
    
    'Add a File structure to hold the machine representation of the text file; _
     the assembler doesn't work with the original text directly
    Let FileCount = FileCount + 1: Let File = FileCount
    ReDim Preserve Files(1 To FileCount) As oz80File
    
    'Record the file path for easy reference in the future
    Let Files(File).Path = FilePath
    
    'Explode the source code file into tokens:
    
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_TOKENISE)
    Call AssembleFile_Tokenise(FilePath, Files(File).Stream)
    Call Profiler.EventEnd(EVENT_TOKENISE)
    '*** TIMER END ***
    
    'If the file failed to parse, an error code / message will already have _
     been raised via an event, we only need to exit
    If Error Then Exit Sub
    
    'When we begin assembling from the Syntax Tree, we want to do so _
     from the beginning of the source file now being added
    Let Tree.Index = Tree.Count
    Let Node = Tree.Nodes(Tree.Index)
    
    'Now convert the linear word-by-word Token Stream in a strucutred, _
     validated Abstract Syntax Tree:
    
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_BUILDTREE)
    Call AssembleFile_SyntaxTree(Files(File).Stream)
    Call Profiler.EventEnd(EVENT_BUILDTREE)
    '*** TIMER END ***
    
    'If the source code structure couldn't be validated, _
     the error will already have been raised and we can exit now
    If Error Then Exit Sub
    
    '[4] Assemble the file: _
     ----------------------------------------------------------------------------------
'    Call RaiseMessage(OZ80_LOG_ACTION, "Assembling ""{$}""", FilePath)
    
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_ASSEMBLE)
        
    'Now begin assembling the file. If an `INCLUDE` statement is encountered, _
     the referenced file will also be tokenised and assembled
    Call AssembleRoot
    
    Call Profiler.EventEnd(EVENT_ASSEMBLE)
    '*** TIMER END ***
    
'    If not Error Then Call RaiseMessage(OZ80_LOG_INFO, "Assembly Complete for ""{$}""", FilePath)
End Sub

'AssembleFile_Tokenise : Parse Source Code Text into Tokens
'======================================================================================
'FilePath       |
'TokenStream    | The `TokenStream` type to populate
'======================================================================================
Private Sub AssembleFile_Tokenise( _
    ByRef FilePath As String, _
    ByRef TokenStream As oz80TokenStream _
)
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_FILELOAD_API)
    
    'Setup the bluString that's going to hold our source text
    Dim Source As bluString
    Set Source = New bluString
    
    'Attempt to load the source file
    'TODO: More precise error handling?
    If _
        Source.Load(FilePath, LoadOption_NormalizeWhitespace) <> 0 _
    Then Call RaiseError( _
        OZ80_ERROR_FILE_READ, 0, 0, _
        "Could not read the file" _
    ): Exit Sub
    
    'For speed we will walk through the source text as Unicode characters; _
     this array will be mapped to the source text held within the bluString
    Dim Buffer() As Integer
    'This does the mapping (the data is not duplicated). _
     The buffer has to be released with `CharArrayUnaccess` when done
    Let Buffer = Source.CharArrayAccess()
    
    Call Profiler.EventEnd(EVENT_FILELOAD_API)
    '*** TIMER END ***
    
    '----------------------------------------------------------------------------------
    
    'Which character in the source code we're parsing. Since the way the routine is _
     written (we increase the position before each word begins) begin with -1 so that _
     we may really begin at 0
    Dim SourcePos As Long
    Let SourcePos = -1
    
    'Line and column number in the source text, as it is read. 1-based. _
     The Line & Col number that a word starts on is kept in `TokenLine` & `TokenCol`
    Dim Line As Long, Col As Long
    Let Line = 1
    
    'Get a reference to the Unicode properties look-up table. This doesn't need to be _
     manually freed as that Class only goes out of scope when the application quits
    Dim Unicode() As Integer
    Let Unicode = blu.Strings.UnicodeProperties()
    
    'We need to be able to tell the difference between a property being used as a _
     parameter on a Z80 instruction, and a label-property definition (within a chunk) _
     which occurs on a new line. We keep track of the previous line number used and _
     look for a property occuring on a new line
    Dim PrevLine As Long
    Let PrevLine = 1
    
    'We will capture the word as we read it here. The bluString is both fast and _
     allows us to access the string as bytes or text
    Dim Word As bluString
    'Set the buffer for capturing the word to 1024 chars to avoid lots of memory _
     re-allocations if larger strings in the source are encountered
    Set Word = New bluString
    Let Word.Buffer = BUFFER_1K
    
    'Where in the source file the captured word begins. _
     This will be used for error information right the way through
    Dim WordLine As Long
    Dim WordCol As Long
    
    'These will be the deduced information from the word; _
     when we work out the scope, these will be written to the syntax tree
    Dim WordToken As OZ80_TOKEN         'An index number to represent the word
    Dim WordAttr As OZ80_ATTR           'Attributes of the word, e.g. Keyword, List
    Dim WordValue As Double             'If a number, the numerical value
    
    'Prepare a large buffer to accept the Token Stream (using `ReDim` for every token _
     added is *incredibly* slow). If the tokens overflow this buffer, we'll increase _
     it by another chunk there and then
    Let TokenStream.Buffer = TOKEN_BUFFER
    ReDim TokenStream.Tokens(0 To TOKEN_BUFFER) As oz80Token
    
    'Begin:
    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
    
SkipWord:
    'When we need to ignore a word (comments for example), it's faster to jump up _
     here and drop into the routine again. We also need to ensure that our profiling _
     timer does not stretch across multiple ignored words & real words
    Call Profiler.EventEnd(EVENT_PARSE_WORD)
    
ReadWord:
    '*** TIMER BEGIN ***
    Call Profiler.EventBegin(EVENT_PARSE_WORD)
    
    'Reset the captured word
    Call Word.Clear
    
    'A cache is *much* faster than accessing the property many times over
    Dim WordLength As Long: Let WordLength = 0
    
    Dim WordContinue As OZ80_SYNTAX
    Dim WordBlock As Long
    
    Let WordAttr = ATTR_NONE
    Let WordValue = 0
    
    '----------------------------------------------------------------------------------
SkipChar:
    'Move on to the next character in the source
    Let SourcePos = SourcePos + 1
ReadChar:
    'Read a single character:
    Dim Char As OZ80_SYNTAX
    Let Char = Buffer(SourcePos)
    'If the file ends? (the array we get from the bluString is null-terminated, _
     so we don't need to do length comparisons every loop!)
    If Char = ASC_NUL Then
        'If no word is currently underway then we are all done
        If WordLength = 0 Then GoTo Finish
        'Otherwise process this final word
        'TODO: Would error upon return to SkipChar?
        GoTo TokenWord
    End If
    
    '[a]    Start of a word:
    '----------------------------------------------------------------------------------
    'The start of a word determines much of how we will process it. _
     Words end on whitespace, except when comments or quotes are started
    If WordLength = 0 Then
        'Skip whitespace before the word begins:
        If Char = ASC_SPC Then Let Col = Col + 1: GoTo SkipChar
        'Tab goes to the next tab stop boundary
        If Char = ASC_TAB _
            Then Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH): GoTo SkipChar
        'Start a new line, reset the column count
        If Char = ASC_LF Then Let Line = Line + 1: Let Col = 0: GoTo SkipChar
        'All other text moves along one column
        Let Col = Col + 1
        
        'One-letter words:
        '(Commas, brackets and operators can be used without spaces between; _
          therefore when we encounter one we end the word immediately)
        If Char = SYNTAX_NEXT Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_ADD Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_SUB Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_MUL Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_DIV Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_POW Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_MOD Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_OR Then GoTo EndWord
        If Char = SYNTAX_OPERATOR_AND Then GoTo EndWord
        If Char = SYNTAX_BRACES_OPEN Then GoTo EndWord
        If Char = SYNTAX_BRACES_CLOSE Then GoTo EndWord
        If Char = SYNTAX_PARENS_OPEN Then GoTo EndWord
        If Char = SYNTAX_PARENS_CLOSE Then GoTo EndWord
        If Char = SYNTAX_SQUARE_OPEN Then GoTo EndWord
        If Char = SYNTAX_SQUARE_CLOSE Then GoTo EndWord
        
        'Encountered a comment marker...
        If Char = SYNTAX_COMMENT Then
            'Enable the line-parsing mode
            Let WordContinue = SYNTAX_COMMENT
            'Do a look-ahead to see if this is a multi-line comment
            If Buffer(SourcePos + 1) = SYNTAX_COMMENT Then Let WordBlock = 1
            
        'Is this text? (in which case, ignore wordbreaks until text end)
        ElseIf Char = SYNTAX_TEXT Then
            'Enable the line-parsing mode
            Let WordContinue = SYNTAX_TEXT
            Let WordBlock = 1
            
        'Is this documentation text?
        ElseIf Char = SYNTAX_DOC Then
            'Do a look-ahead to see if this is documentation text
            If Buffer(SourcePos + 1) = SYNTAX_DOC Then Let WordContinue = SYNTAX_DOC
        End If
        
        'Capture the start of the word
        Let WordLine = Line: Let WordCol = Col
        Dim FirstChar As OZ80_SYNTAX
        Let FirstChar = Char
        
    '[b]    Processing of single words: (end on whitespace)
    '----------------------------------------------------------------------------------
    ElseIf WordContinue = 0 Then
        'Handle whitespace:
        If Char = ASC_SPC Then Let Col = Col + 1: GoTo TokenWord
        'Tab goes to the next tab stop boundary
        If Char = ASC_TAB _
            Then Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH): GoTo TokenWord
        'End the word with the line-break
        If Char = ASC_LF Then Let Line = Line + 1: Let Col = 0: GoTo TokenWord
        'All other text moves along one column
        Let Col = Col + 1
        
        'One-letter words:
        '(Commas, brackets and operators can be used without spaces between; _
          therefore when we encounter one we end the word immediately)
        If Char = SYNTAX_NEXT Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_ADD Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_SUB Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_MUL Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_DIV Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_POW Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_MOD Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_OR Then GoTo WordBreak
        If Char = SYNTAX_OPERATOR_AND Then GoTo WordBreak
        If Char = SYNTAX_BRACES_OPEN Then GoTo WordBreak
        If Char = SYNTAX_BRACES_CLOSE Then GoTo WordBreak
        If Char = SYNTAX_PARENS_OPEN Then GoTo WordBreak
        If Char = SYNTAX_PARENS_CLOSE Then GoTo WordBreak
        If Char = SYNTAX_SQUARE_OPEN Then GoTo WordBreak
        If Char = SYNTAX_SQUARE_CLOSE Then GoTo WordBreak
        
        'A register hint is split off and ignored
        If Char = SYNTAX_HINT Then GoTo WordBreak
        If Char = SYNTAX_SHADOW Then GoTo WordBreak
        
        'If a comment appears at the end of a word without any space to separate, _
         we need to first keep the word before switching to comment parsing
        If Char = SYNTAX_COMMENT Then GoTo WordBreak
        
    '[c]    Processing of multi-words: (e.g. line-comments, documentation text)
    '----------------------------------------------------------------------------------
    ElseIf WordBlock = 0 Then
        'Handle whitespace:
        If Char = ASC_TAB Then
            'Tab goes to the next tab stop boundary
            Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH)
            
        ElseIf Char = ASC_LF Then
            'Start a new line, reset the column count
            Let Line = Line + 1: Let Col = 0
            'If a comment, throw it away
            If WordContinue = SYNTAX_COMMENT Then
                Let WordContinue = 0
                GoTo SkipWord
            Else
                Let WordContinue = 0
                GoTo TokenWord
            End If
        
        ElseIf Char = SYNTAX_DOC Then
            Let Col = Col + 1
            'Are we processing documentation text?
            If WordContinue = SYNTAX_DOC Then
                'Do a look-ahead for the triple-backtick
                If Buffer(SourcePos + 1) = SYNTAX_DOC Then Let WordBlock = 1
            End If
        Else
            'All other text moves along one column
            Let Col = Col + 1
        End If
        
    '[d]    Processing of multi-line words (e.g. multi-line comments)
    '----------------------------------------------------------------------------------
    Else
        'Handle whitespace:
        If Char = ASC_TAB Then
            'Tab goes to the next tab stop boundary
            Let Col = Col + TAB_WIDTH - (Col Mod TAB_WIDTH)
            
        ElseIf Char = ASC_LF Then
            'If currently in a quote, this is an error!
            If WordContinue = SYNTAX_TEXT Then Call RaiseError( _
                OZ80_ERROR_INVALID, Line, Col, _
                "Quoted text must be terminated before a line-break. " _
              & "If you want to include line-breaks in quoted text, use the ASCII " _
              & "code 13 between pieces of quoted text. For example: " & vbNewLine _
              & vbNewLine _
              & vbTab & "WORD" & vbTab & """The quick brown fox"", 13, " & vbNewLine _
              & vbTab & vbTab & """jumps over the lazy dog.""" _
            ): GoTo Finish
            'Start a new line, reset the column count
            Let Line = Line + 1: Let Col = 0
            
        ElseIf Char = SYNTAX_TEXT Then
            Let Col = Col + 1
            'If processing quoted text, end on the quote
            If WordContinue = SYNTAX_TEXT Then
                'Capture the text
                Let WordContinue = 0: Let WordBlock = 0
                GoTo TokenWord
            End If
            
        ElseIf Char = SYNTAX_COMMENT Then
            Let Col = Col + 1
            'Are we parsing a comment?
            If WordContinue = SYNTAX_COMMENT Then
                'Do a look ahead for the second semi-colon
                If Buffer(SourcePos + 1) = SYNTAX_COMMENT Then
                    'Discard the comment
                    Let WordContinue = 0: Let WordBlock = 0
                    GoTo SkipWord
                End If
            End If
        
        ElseIf Char = SYNTAX_DOC Then
            Let Col = Col + 1
            'Are we parsing documentation text?
            If WordContinue = SYNTAX_DOC Then
                'Do a look ahead for another backtick
                If Buffer(SourcePos + 1) = SYNTAX_DOC Then
                    'Is this a third back-tick?
                    If Buffer(SourcePos - 1) = SYNTAX_DOC Then
                        'Capture the documentation text
                        Let WordContinue = 0: Let WordBlock = 0
                        GoTo TokenWord
                    End If
                End If
            End If
            
        Else
            'All other text moves along one column
            Let Col = Col + 1
        End If
        
    End If
           
    '----------------------------------------------------------------------------------
    
    'Since this is a valid letter, capture it
    Call Word.CharPush(Char): Let WordLength = WordLength + 1
    GoTo SkipChar
    
EndWord:
    Call Word.Clear.CharPush(Char): Let WordLength = 1
    Let WordLine = Line: Let WordCol = Col
    Let FirstChar = Char
    GoTo TokenWord
    
WordBreak:
    'If the word is longer than one character, then save the current character _
     for use in the next word. i.e. when a comma is reached, process the text _
     before the comma as a word, and take the comma as the next word
    Let SourcePos = SourcePos - 1: Let Col = Col - 1
    
    
    '----------------------------------------------------------------------------------
TokenWord:
    
    'Jump ahead based on the ASCII code of the first letter of the word; _
     this is both very fast and allows us to ignore case by mapping the _
     lower-case and upper-case ASCII to the same destination label!
     
    'As VB6 doesn't allow us to intersperse code with comments, _
     I'll annotate the jump list here
     
    ' ! " # $ % & ' ( )                 'Apostrophe is register hint and skipped
    ' * + , - . /                       'Comma is purely cosmetic
    ' 1 2 3 4 5 5 6 7 8 9 0             'Numbers are all handled as one
    ' : ; < = > ? @                     'Much of this is unimplemented at the moment
    ' A B C D E F G H I J K L M         'There are no F or G words
    ' N O P Q R S T U V W X Y Z         'There are no Q, U, V or Y words
    ' [ \ ] ^ _ `                       'No Underscore words
    ' a b c d e f g h i j k l m         'There are no F or G words
    ' n o p q r s t u v w x y z         'There are no Q, U, V or Y words
    ' { | } ~                           'Tilde is unused at this moment
    
    On FirstChar - 32 GoTo _
        isName, smark, isName, isName, perc, amp, SkipWord, lbr, rbr _
      , star, plus, SkipWord, dash, isName, fs _
      , num, num, num, num, num, num, num, num, num, num _
      , isName, SkipWord, WordError, WordError, WordError, WordError, WordError _
      , a, b, c, d, e, WordError, WordError, h, i, j, k, l, m _
      , n, o, p, WordError, r, s, t, WordError, WordError, w, x, WordError, z _
      , lsb, bs, rsb, acute, WordError, backtick _
      , a, b, c, d, e, WordError, WordError, h, i, j, k, l, m _
      , n, o, p, WordError, r, s, t, WordError, WordError, w, x, WordError, z _
      , lcb, bar, rcb, WordError
    
    'Null, or not an ASCII char from 33 to 126
    GoTo WordError
    
    'Just a loop counter
    Dim i As Long
    'These are used whilst converting words into real numbers
    Dim Number As Double, Multiplier As Long
    
isName:
    '[a]    Defined Names & Hexadecimal Numbers
    '----------------------------------------------------------------------------------
    '[a.1]  Check the one-letter names:
    '..................................................................................
    If WordLength = 1 Then
        'Which one is it?
        If FirstChar = SYNTAX_PROPERTY Then
            'An anonymous property, "."
            Let WordToken = TOKEN_PROPERTY_ANON
            Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_PROPERTY
            
        ElseIf FirstChar = SYNTAX_LABEL Then
            'Self-reference Label, ":"
            Let WordToken = TOKEN_NAME
            Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_LABEL
            
        ElseIf FirstChar = SYNTAX_RAM Then
            'Default RAM Page, "$", the Token uses a Value of 0. _
             The default RAM Page is pre-defined, so is not created here
            Let WordToken = TOKEN_NAME
            Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_RAM
            
        Else
            'It's life Jim, but not as we know it
            Call RaiseError( _
                OZ80_ERROR_INVALID, Line, Col, _
                "Invalid Name. A 1-letter Name was not recognised, " _
              & "only ':', '.' and '$' are defined 1-letter names" _
            ): GoTo Finish
        End If
        GoTo NextWordSkipAttr
    End If
    
    'Get raw access to the characters in the word; this will be faster than using _
     the `Word.Char` accessor repeatedly, though we will have to avoid leaving the _
     function without calling `CharArrayUnaccess` first, to free the reference
    Dim Chars() As Integer
    Let Chars = Word.CharArrayAccess()
    
    'The Name's prefix will determine where we begin typical validation; _
     some prefixes are longer than others
    Dim BeginAt As Long
    
    'The beginning of the name will determine what is allowed to follow
    If FirstChar = SYNTAX_NUMBER_HEX Then
        
        'Hexadecimal Number:
        '..............................................................................
        'Check for a valid hexadecimal digit after the "$"
        If (Unicode(Chars(1)) And C1_XDIGIT) <> 0 Then
            
            'Any more than 8 hexadecimal digits will be an overflow _
             (32-bit number maximum). Note "9" because of the "$" prefix
            If WordLength > 9 Then Call RaiseError( _
                OZ80_ERROR_INVALID, Line, Col, _
                "Invalid hexadecimal number. Hexadecimal numbers begin with '$', " _
              & "followed by up to 8 hexadecimal digits (0-9 & A-F). Your word is " _
              & "too long. Note that RAM Names also begin with '$', but are followed " _
              & "by dot or underscore to avoid clashes with valid hexadecimal " _
              & "numbers like $DEADBEEF" _
            ): GoTo NameErrorOut
            
            'Prime the value with the first hexadecimal digit
            Let WordValue = Hex2Dec(Chars(1))
            
            'Walk the hexadecimal digits
            For i = 2 To WordLength - 1
                'Thanks to Vadym Stetsiak for the simple, fast hex ASCII conversion:
                '<vadmyst.blogspot.co.uk/2012/02/fast-convertion-of-hex-string-into.html>
                Let WordValue = WordValue * 16
                Let WordValue = WordValue + Hex2Dec(Chars(i))
                
                'ERROR: Invalid hexadecimal number!
                If WordValue < 0 Then Call RaiseError( _
                    OZ80_ERROR_INVALID, Line, Col, _
                    "Not a valid hexadecimal number." & vbNewLine & vbNewLine _
                  & "A hexadecimal number begins with '$' followed by 1-8 " _
                  & "hexadecimal digits (0-9, A-F)." _
                ): GoTo NameErrorOut
            Next i
            
            'Add the token for the number
            Let WordToken = TOKEN_NUMBER
            'Free the array reference into the bluString
            Call Word.CharArrayUnaccess(Chars)
            'No need to process as a Name, continue parsing
            GoTo NextWord
            
        End If
        
        'RAM Name:
        '..............................................................................
        Let WordToken = TOKEN_NAME
        Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_RAM Or ATTR_PARAM
        Let BeginAt = 1
        
    ElseIf FirstChar = SYNTAX_PROPERTY Then
        'TODO: NEW/USE property tokens (or do away with this distinction entirely)
        'If the property appears on a new line, it's a label-property definition _
         i.e. it defines a new sub-label of a chunk
        If WordLine <> PrevLine _
            Then Let WordToken = TOKEN_PROPERTY_NEW _
            Else Let WordToken = TOKEN_NAME
        'A property can only be followed by other properties
        Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_PROPERTY
        'Begin validating after the dot
        Let BeginAt = 1
        
    ElseIf FirstChar = SYNTAX_LABEL Then
        'Could be a Label, ":", or a Section, "::"
        If Chars(1) = SYNTAX_LABEL Then
            'It's a Section name
            Let WordToken = TOKEN_NAME
            Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_SECTION
            Let BeginAt = 2
        Else
            'It's a Label
            Let WordToken = TOKEN_NAME
            Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_LABEL
            Let BeginAt = 1
        End If
        
    ElseIf FirstChar = SYNTAX_CONSTANT Then
        Let WordToken = TOKEN_NAME
        Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_CONST
        Let BeginAt = 1
    
    ElseIf FirstChar = SYNTAX_TYPE Then
        Let WordToken = TOKEN_NAME
        Let WordAttr = TokenAttributes(WordToken) Or ATTR_NAME_TYPE
        Let BeginAt = 1
    Else
        Call RaiseError( _
            OZ80_ERROR_INTERNAL, Line, Col, _
            "Unimplemented Name checking. The word was recognised as a Name, but the " _
          & "first character was not handled. This is a mistake on OZ80's part" _
        ): GoTo NameErrorOut
    End If
    
    '[a.4]
    '..................................................................................
    
    'Now walk the remainder of the Name and validate
    Dim c As Long
    For c = BeginAt To WordLength
        'Have we encountered a property?
        If Chars(c) = SYNTAX_PROPERTY Then
            Let c = c + 1
            Let WordAttr = WordAttr Or ATTR_NAME_PROPERTY
            
            'Dot cannot be followed by another dot or number
            If Chars(c) = SYNTAX_PROPERTY Then GoTo NameError
            If (Chars(c) >= ASC_0) And (Chars(c) <= ASC_9) Then GoTo NameError
            'Also dot cannot be the last character
            If Chars(c) = ASC_NUL Then GoTo NameError
        
        'Have we encountered a Label?
        ElseIf Chars(c) = SYNTAX_LABEL Then
            'A Label can only occur after a Section,
            If (WordAttr And ATTR_NAME_SECTION) = 0 Then GoTo NameError
            'And it cannot occur twice
            If (WordAttr And ATTR_NAME_LABEL) <> 0 Then GoTo NameError
            
            Let c = c + 1
            Let WordAttr = WordAttr Or ATTR_NAME_LABEL
            
            'Colon cannot be followed by dot or number
            If Chars(c) = SYNTAX_PROPERTY Then GoTo NameError
            If (Chars(c) >= ASC_0) And (Chars(c) <= ASC_9) Then GoTo NameError
            'Can't have two colons, since a Section prefix is already parsed
            If Chars(c) = SYNTAX_LABEL Then GoTo NameError
            'Also colon cannot be the last character
            If Chars(c) = ASC_NUL Then GoTo NameError
        End If
    Next c
    
    'Add the Name string to the string storage in the Token Stream: _
     We're not going to bother to de-duplicate these as it'll add too much time _
     searching vs. what little extra memory is used (the Token Stream is discarded _
     after the Syntax Tree for it has been built)
    With TokenStream
        Let .StringCount = .StringCount + 1
        ReDim Preserve .Strings(1 To .StringCount) As String
        Let .Strings(.StringCount) = Word.Text
        'Add the index key to the token so that we can retrieve the string later; _
         the one-letter names will have a Value of 0 to distinguish them
        Let WordValue = .StringCount
    End With
    
    'Free the array reference into the bluString
    Call Word.CharArrayUnaccess(Chars)
    GoTo NextWordSkipAttr
    
NameError:
    '..................................................................................
    'TODO:
    Call RaiseError( _
        OZ80_ERROR_INVALID, Line, Col, _
        "Not a valid Name" _
    )
NameErrorOut:
    'Free the array reference into the bluString
    Call Word.CharArrayUnaccess(Chars)
    GoTo Finish
    
smark:
    'Quoted Text:
    '.......................................................................... `"` ...
    Let WordToken = TOKEN_TEXT
    
    With TokenStream
        Let .StringCount = .StringCount + 1
        ReDim Preserve .Strings(1 To .StringCount) As String
        Let .Strings(.StringCount) = Word.Mid(1).Text
        'Add the index key to the token so that we can retrieve the string later; _
         the one-letter names will have a Value of 0 to distinguish them
        Let WordValue = .StringCount
    End With
    
    GoTo NextWord
    
perc:
    '.......................................................................... `%` ...
    'ERROR: Cannot be longer than 32-bits
    If WordLength > 33 Then
        GoTo ErrorOverflow
    End If
    Let Number = 0: Let Multiplier = 0
    For i = WordLength - 1 To 1 Step -1
        If Word.Char(i) = ASC_1 Then
            Let Number = Number + 2 ^ Multiplier
        ElseIf Word.Char(i) <> ASC_0 Then
            'Error - not a valid binary number
            Stop
        End If
        Let Multiplier = Multiplier + 1
    Next i
    'Tokenise number
    Let WordToken = TOKEN_NUMBER
    Let WordValue = Number
    GoTo NextWord
    
amp:
    '.......................................................................... `&` ...
    Let WordToken = TOKEN_OPERATOR_AND
    GoTo NextWord
    
lbr:
    '.......................................................................... `(` ...
    Let WordToken = TOKEN_PARENS_OPEN
    GoTo NextWord
    
rbr:
    '.......................................................................... `)` ...
    Let WordToken = TOKEN_PARENS_CLOSE
    GoTo NextWord
    
star:
    '.......................................................................... `*` ...
    Let WordToken = TOKEN_OPERATOR_MUL
    GoTo NextWord
    
plus:
    '.......................................................................... `+` ...
    Let WordToken = TOKEN_OPERATOR_ADD
    GoTo NextWord
    
dash:
    '.......................................................................... `-` ...
    Let WordToken = TOKEN_OPERATOR_SUB
    GoTo NextWord
    
fs:
    '.......................................................................... `/` ...
    Let WordToken = TOKEN_OPERATOR_DIV
    GoTo NextWord
    
num:
    '........................................................................ `0-9` ...
    'Validate the number
    Let Number = 0: Let Multiplier = 1
    For i = WordLength - 1 To 0 Step -1
        'ERROR: If not a 0-9 digit then this is not a decimal number!
        If Not Word.CharIsNumeric(i, True) Then Call RaiseError( _
            OZ80_ERROR_INVALID, Line, Col, _
            "Not a valid decimal number." & vbNewLine & vbNewLine _
          & "Any word that begins with a number is treated as a decimal number and " _
          & "must contain numeric characters (0-9) and optionally one decimal point. " _
          & "For numbers with a decimal point, there must always be at least one " _
          & "number preceeding the decimal point and another proceeding it, " _
          & "i.e. '9.9'." _
        ): GoTo Finish
        
        'Convert the number from ASCII codes into a value
        On Error GoTo ErrorOverflow
        'I've tried a lookup table here and it's exactly the same speed
        Let Number = Number + (Word.Char(i) - ASC_0) * Multiplier
        'Don't allow the multiplier to overflow when the number doesn't
        If i > 0 Then Let Multiplier = Multiplier * 10
        On Error GoTo 0
    Next i
    'Add the token with the number value
    Let WordToken = TOKEN_NUMBER
    Let WordValue = Number
    GoTo NextWord
    
lthan:
    '.......................................................................... `<` ...
    'TODO: Less than operator
    GoTo NextWord
    
equal:
    '.......................................................................... `=` ...
    'TODO: Equals operator
    GoTo NextWord
    
gthan:
    '.......................................................................... `>` ...
    'TODO: Greater than operator
    GoTo NextWord
    
qmark:
    '.......................................................................... `?` ...
    'TODO: Ternery operator?
    GoTo NextWord
    
at:
    '.......................................................................... `@` ...
    'TODO: Macros
    GoTo NextWord
    
a:
    '.......................................................................... `A` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_A
    ElseIf Word.Equals(STR_AF) Then Let WordToken = TOKEN_Z80_AF
    'Z80 Instructions:
    ElseIf Word.Equals(STR_ADC) Then Let WordToken = TOKEN_Z80_ADC
    ElseIf Word.Equals(STR_ADD) Then Let WordToken = TOKEN_Z80_ADD
    ElseIf Word.Equals(STR_AND) Then Let WordToken = TOKEN_Z80_AND
    Else: GoTo WordError: End If
    GoTo NextWord
    
b:
    '.......................................................................... `B` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_B
    ElseIf Word.Equals(STR_BC) Then Let WordToken = TOKEN_Z80_BC
    'Z80 Instructions:
    ElseIf Word.Equals(STR_BIT) Then Let WordToken = TOKEN_Z80_BIT
    'Keywords:
    ElseIf Word.Equals(STR_BOOL) Then Let WordToken = TOKEN_KEYWORD_BOOL
    ElseIf Word.Equals(STR_BYTE) Then Let WordToken = TOKEN_KEYWORD_BYTE
    Else: GoTo WordError: End If
    GoTo NextWord
    
c:
    '.......................................................................... `C` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_C
    'Z80 Instructions:
    ElseIf Word.Equals(STR_CALL) Then Let WordToken = TOKEN_Z80_CALL
    ElseIf Word.Equals(STR_CCF) Then Let WordToken = TOKEN_Z80_CCF
    ElseIf Word.Equals(STR_CP) Then Let WordToken = TOKEN_Z80_CP
    ElseIf Word.Equals(STR_CPD) Then Let WordToken = TOKEN_Z80_CPD
    ElseIf Word.Equals(STR_CPDR) Then Let WordToken = TOKEN_Z80_CPDR
    ElseIf Word.Equals(STR_CPI) Then Let WordToken = TOKEN_Z80_CPI
    ElseIf Word.Equals(STR_CPIR) Then Let WordToken = TOKEN_Z80_CPIR
    ElseIf Word.Equals(STR_CPL) Then Let WordToken = TOKEN_Z80_CPL
    'Keywords:
    ElseIf Word.Equals(STR_CONST) Then Let WordToken = TOKEN_KEYWORD_CONST
    Else: GoTo WordError: End If
    GoTo NextWord
    
d:
    '.......................................................................... `D` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_D
    ElseIf Word.Equals(STR_DE) Then Let WordToken = TOKEN_Z80_DE
    'Z80 Instructions:
    ElseIf Word.Equals(STR_DAA) Then Let WordToken = TOKEN_Z80_DAA
    ElseIf Word.Equals(STR_DEC) Then Let WordToken = TOKEN_Z80_DEC
    ElseIf Word.Equals(STR_DI) Then Let WordToken = TOKEN_Z80_DI
    ElseIf Word.Equals(STR_DJNZ) Then Let WordToken = TOKEN_Z80_DJNZ
    Else: GoTo WordError: End If
    GoTo NextWord
    
e:
    '.......................................................................... `E` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_E
    'Z80 Instructions:
    ElseIf Word.Equals(STR_EI) Then Let WordToken = TOKEN_Z80_EI
    ElseIf Word.Equals(STR_EX) Then Let WordToken = TOKEN_Z80_EX
    ElseIf Word.Equals(STR_EXX) Then Let WordToken = TOKEN_Z80_EXX
    Else: GoTo WordError: End If
    GoTo NextWord
        
h:
    '.......................................................................... `H` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_H
    ElseIf Word.Equals(STR_HL) Then Let WordToken = TOKEN_Z80_HL
    'Z80 Instructions:
    ElseIf Word.Equals(STR_HLT) Then Let WordToken = TOKEN_Z80_HALT
    ElseIf Word.Equals(STR_HALT) Then Let WordToken = TOKEN_Z80_HALT
    'Keywords:
    ElseIf Word.Equals(STR_HELP) Then Let WordToken = TOKEN_KEYWORD_HELP
    ElseIf Word.Equals(STR_HI) Then Let WordToken = TOKEN_KEYWORD_HI
    Else: GoTo WordError: End If
    GoTo NextWord
    
i:
    '.......................................................................... `I` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_I
    ElseIf Word.Equals(STR_IX) Then Let WordToken = TOKEN_Z80_IX
    ElseIf Word.Equals(STR_IY) Then Let WordToken = TOKEN_Z80_IY
    ElseIf Word.Equals(STR_IXH) Then Let WordToken = TOKEN_Z80_IXH
    ElseIf Word.Equals(STR_IXL) Then Let WordToken = TOKEN_Z80_IXL
    ElseIf Word.Equals(STR_IYH) Then Let WordToken = TOKEN_Z80_IYH
    ElseIf Word.Equals(STR_IYL) Then Let WordToken = TOKEN_Z80_IYL
    'Z80 Instructions:
    ElseIf Word.Equals(STR_IM) Then Let WordToken = TOKEN_Z80_IM
    ElseIf Word.Equals(STR_IN) Then Let WordToken = TOKEN_Z80_IN
    ElseIf Word.Equals(STR_INC) Then Let WordToken = TOKEN_Z80_INC
    ElseIf Word.Equals(STR_IND) Then Let WordToken = TOKEN_Z80_IND
    ElseIf Word.Equals(STR_INDR) Then Let WordToken = TOKEN_Z80_INDR
    ElseIf Word.Equals(STR_INI) Then Let WordToken = TOKEN_Z80_INI
    ElseIf Word.Equals(STR_INIR) Then Let WordToken = TOKEN_Z80_INIR
    'Keywords:
    ElseIf Word.Equals(STR_INCLUDE) Then Let WordToken = TOKEN_KEYWORD_INCLUDE
    ElseIf Word.Equals(STR_INDEX) Then Let WordToken = TOKEN_KEYWORD_INDEX
    ElseIf Word.Equals(STR_INTERRUPT) Then Let WordToken = TOKEN_KEYWORD_INTERRUPT
    Else: GoTo WordError: End If
    GoTo NextWord
    
j:
    '.......................................................................... `J` ...
    'Z80 Instructions:
    If Word.Equals(STR_JP) Then
        Let WordToken = TOKEN_Z80_JP
    ElseIf Word.Equals(STR_JR) Then Let WordToken = TOKEN_Z80_JR
    'There are no `J` registers
    Else: GoTo WordError: End If
    GoTo NextWord
    
k:
    '.......................................................................... `K` ...
    'Number suffixes K & KB
    If WordLength = 1 Then
        Let WordToken = TOKEN_SUFFIX_K
    ElseIf Word.Equals(STR_KB) Then Let WordToken = TOKEN_SUFFIX_KB
    'There are no `K` registers
    Else: GoTo WordError: End If
    GoTo NextWord
    
l:
    '.......................................................................... `L` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_L
    'Keywords
    ElseIf Word.Equals(STR_LO) Then Let WordToken = TOKEN_KEYWORD_LO
    ElseIf Word.Equals(STR_LONG) Then Let WordToken = TOKEN_KEYWORD_LONG
    'Z80 instructions
    ElseIf Word.Equals(STR_LD) Then Let WordToken = TOKEN_Z80_LD
    ElseIf Word.Equals(STR_LDD) Then Let WordToken = TOKEN_Z80_LDD
    ElseIf Word.Equals(STR_LDDR) Then Let WordToken = TOKEN_Z80_LDDR
    ElseIf Word.Equals(STR_LDI) Then Let WordToken = TOKEN_Z80_LDI
    ElseIf Word.Equals(STR_LDIR) Then Let WordToken = TOKEN_Z80_LDIR
    Else: GoTo WordError: End If
    GoTo NextWord
    
m:
    '.......................................................................... `M` ...
    'Z80 flag conditions
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_M
    'There is no `M` register
    Else: GoTo WordError: End If
    GoTo NextWord
    
n:
    '.......................................................................... `N` ...
    'Z80 Instructions:
    If Word.Equals(STR_NEG) Then
        Let WordToken = TOKEN_Z80_NEG
    ElseIf Word.Equals(STR_NOP) Then Let WordToken = TOKEN_Z80_NOP
    'Keywords:
    ElseIf Word.Equals(STR_NYBL) Then Let WordToken = TOKEN_KEYWORD_NYBL
    'Z80 flag conditions
    ElseIf Word.Equals(STR_NC) Then Let WordToken = TOKEN_Z80_NC
    ElseIf Word.Equals(STR_NZ) Then Let WordToken = TOKEN_Z80_NZ
    'There are no `N` registers
    Else: GoTo WordError: End If
    GoTo NextWord
    
o:
    '.......................................................................... `O` ...
    'Z80 Instructions:
    If Word.Equals(STR_OR) Then
        Let WordToken = TOKEN_Z80_OR
    ElseIf Word.Equals(STR_OUT) Then Let WordToken = TOKEN_Z80_OUT
    ElseIf Word.Equals(STR_OUTD) Then Let WordToken = TOKEN_Z80_OUTD
    ElseIf Word.Equals(STR_OTDR) Then Let WordToken = TOKEN_Z80_OTDR
    ElseIf Word.Equals(STR_OUTI) Then Let WordToken = TOKEN_Z80_OUTI
    ElseIf Word.Equals(STR_OTIR) Then Let WordToken = TOKEN_Z80_OTIR
    'There are no `O` registers
    Else: GoTo WordError: End If
    GoTo NextWord
    
p:
    '.......................................................................... `P` ...
    'Z80 flag conditions
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_P
    ElseIf Word.Equals(STR_PE) Then Let WordToken = TOKEN_Z80_PE
    ElseIf Word.Equals(STR_PO) Then Let WordToken = TOKEN_Z80_PO
    'Z80 instructions
    ElseIf Word.Equals(STR_POP) Then Let WordToken = TOKEN_Z80_POP
    ElseIf Word.Equals(STR_PUSH) Then Let WordToken = TOKEN_Z80_PUSH
    'Keywords
    ElseIf Word.Equals(STR_PARAMS) Then Let WordToken = TOKEN_KEYWORD_PARAMS
    ElseIf Word.Equals(STR_PROC) Then Let WordToken = TOKEN_KEYWORD_PROC
    Else: GoTo WordError: End If
    GoTo NextWord
    
r:
    '.......................................................................... `R` ...
    'Z80 Registers:
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_R
    'Z80 Instructions:
    ElseIf Word.Equals(STR_RES) Then Let WordToken = TOKEN_Z80_RES
    ElseIf Word.Equals(STR_RET) Then Let WordToken = TOKEN_Z80_RET
    ElseIf Word.Equals(STR_RETI) Then Let WordToken = TOKEN_Z80_RETI
    ElseIf Word.Equals(STR_RETN) Then Let WordToken = TOKEN_Z80_RETN
    ElseIf Word.Equals(STR_RL) Then Let WordToken = TOKEN_Z80_RL
    ElseIf Word.Equals(STR_RLA) Then Let WordToken = TOKEN_Z80_RLA
    ElseIf Word.Equals(STR_RLC) Then Let WordToken = TOKEN_Z80_RLC
    ElseIf Word.Equals(STR_RLCA) Then Let WordToken = TOKEN_Z80_RLCA
    ElseIf Word.Equals(STR_RLD) Then Let WordToken = TOKEN_Z80_RLD
    ElseIf Word.Equals(STR_RR) Then Let WordToken = TOKEN_Z80_RR
    ElseIf Word.Equals(STR_RRA) Then Let WordToken = TOKEN_Z80_RRA
    ElseIf Word.Equals(STR_RRC) Then Let WordToken = TOKEN_Z80_RRC
    ElseIf Word.Equals(STR_RRCA) Then Let WordToken = TOKEN_Z80_RRCA
    ElseIf Word.Equals(STR_RRD) Then Let WordToken = TOKEN_Z80_RRD
    ElseIf Word.Equals(STR_RST) Then Let WordToken = TOKEN_Z80_RST
    'Keywords:
    ElseIf Word.Equals(STR_RAM) Then Let WordToken = TOKEN_KEYWORD_RAM
    ElseIf Word.Equals(STR_RAMPAGE) Then Let WordToken = TOKEN_KEYWORD_RAMPAGE
    ElseIf Word.Equals(STR_RETURN) Then Let WordToken = TOKEN_KEYWORD_RETURN
    Else: GoTo WordError: End If
    GoTo NextWord
    
s:
    '.......................................................................... `S` ...
    'Z80 Instructions:
    If Word.Equals(STR_SBC) Then
        Let WordToken = TOKEN_Z80_SBC
    ElseIf Word.Equals(STR_SCF) Then Let WordToken = TOKEN_Z80_SCF
    ElseIf Word.Equals(STR_SET) Then Let WordToken = TOKEN_Z80_SET
    ElseIf Word.Equals(STR_SLA) Then Let WordToken = TOKEN_Z80_SLA
    ElseIf Word.Equals(STR_SLL) Then Let WordToken = TOKEN_Z80_SLL
    ElseIf Word.Equals(STR_SRA) Then Let WordToken = TOKEN_Z80_SRA
    ElseIf Word.Equals(STR_SRL) Then Let WordToken = TOKEN_Z80_SRL
    ElseIf Word.Equals(STR_SUB) Then Let WordToken = TOKEN_Z80_SUB
    'Keywords
    ElseIf Word.Equals(STR_SECTION) Then Let WordToken = TOKEN_KEYWORD_SECTION
    ElseIf Word.Equals(STR_SLOT) Then Let WordToken = TOKEN_KEYWORD_SLOT
    ElseIf Word.Equals(STR_START) Then Let WordToken = TOKEN_KEYWORD_START
    ElseIf Word.Equals(STR_SYSTEM) Then Let WordToken = TOKEN_KEYWORD_SYSTEM
    'Registers
    ElseIf Word.Equals(STR_SP) Then Let WordToken = TOKEN_Z80_SP
    Else: GoTo WordError: End If
    GoTo NextWord
    
t:
    '.......................................................................... `T` ...
    'Keywords:
    If Word.Equals(STR_TABLE) Then
        Let WordToken = TOKEN_KEYWORD_TABLE
    ElseIf Word.Equals(STR_TYPE) Then Let WordToken = TOKEN_KEYWORD_TYPE
    ElseIf Word.Equals(STR_TRIP) Then Let WordToken = TOKEN_KEYWORD_TRIP
    'There is no `T` register
    Else: GoTo WordError: End If
    GoTo NextWord
    
w:
    '.......................................................................... `W` ...
    'Keywords:
    If Word.Equals(STR_WORD) Then
        Let WordToken = TOKEN_KEYWORD_WORD
    'There is no `W` register
    Else: GoTo WordError: End If
    GoTo NextWord
    
x:
    '.......................................................................... `X` ...
    'If only `x`, then it's the repeat operator
    If WordLength = 1 Then
        Let WordToken = TOKEN_OPERATOR_REP
    'Z80 Instructions:
    ElseIf Word.Equals(STR_XOR) Then Let WordToken = TOKEN_Z80_XOR
    'There is no `X` register
    Else: GoTo WordError: End If
    GoTo NextWord
    
z:
    '.......................................................................... `Z` ...
    'Z80 flag conditions
    If WordLength = 1 Then
        Let WordToken = TOKEN_Z80_Z
    'There is no 'z' register
    Else: GoTo WordError: End If
    GoTo NextWord
    
lsb:
    '.......................................................................... `[` ...
    Let WordToken = TOKEN_SQUARE_OPEN
    GoTo NextWord
    
bs:
    '.......................................................................... `\` ...
    Let WordToken = TOKEN_OPERATOR_MOD
    GoTo NextWord
    
rsb:
    '.......................................................................... `]` ...
    Let WordToken = TOKEN_SQUARE_CLOSE
    GoTo NextWord
    
acute:
    '.......................................................................... `^` ...
    Let WordToken = TOKEN_OPERATOR_POW
    GoTo NextWord
    
backtick:
    '..........................................................................  `  ...
    'TODO: Any better way of doing this? (persisting some info from `ReadWord`?)
    If WordLength > 2 Then
        If Word.Char(1) = SYNTAX_DOC Then
            Let WordToken = TOKEN_DOC
        Else
            GoTo SkipWord
        End If
    Else
        If Word.Char(1) = SYNTAX_DOC Then
            Let WordToken = TOKEN_DOC
        Else
            GoTo SkipWord
        End If
    End If
    
    With TokenStream
        Let .StringCount = .StringCount + 1
        ReDim Preserve .Strings(1 To .StringCount) As String
        'TODO: Strip the prefix?
        Let .Strings(.StringCount) = Word.Text
        'Add the index key to the token so that we can retrieve the string later; _
         the one-letter names will have a Value of 0 to distinguish them
        Let WordValue = .StringCount
    End With
    
    GoTo NextWord
    
lcb:
    '.......................................................................... `{` ...
    Let WordToken = TOKEN_BRACES_OPEN
    GoTo NextWord
    
bar:
    '.......................................................................... `|` ...
    Let WordToken = TOKEN_OPERATOR_OR
    GoTo NextWord
    
rcb:
    '.......................................................................... `}` ...
    Let WordToken = TOKEN_BRACES_CLOSE
    GoTo NextWord
    
NextWord:
    '----------------------------------------------------------------------------------
    'Look up the token attributes (e.g. if it's a Keyword, Expression, List &c.)
    Let WordAttr = TokenAttributes(WordToken)
    
NextWordSkipAttr:
    With TokenStream
        'Add the information gleaned to the current token _
         and move to the next empty slot
        With .Tokens(.Length)
            Let .Kind = WordToken
            Let .Attr = WordAttr
            Let .Line = WordLine
            Let .Col = WordCol
            Let .Value = WordValue
        End With
        
        'Remember the line number used for this token to check for properties occuring _
         on a new line -- this constitutes a property definition in a Chunk
        Let PrevLine = WordLine
        
        'Move to the next token number
        Let .Length = .Length + 1
        
        'Note that the tokens array is given a large size to begin with as ReDimming _
         it with every token added is *incredibly* slow. If we go over this buffer _
         size, we need to increase it by a large chunk. Once parsing is complete the _
         buffer will be trimmed down to the final size
        If .Length > .Buffer Then
            Let .Buffer = .Buffer + TOKEN_BUFFER
            ReDim Preserve .Tokens(0 To .Buffer) As oz80Token
        End If
    End With
    
    'Word has been parsed, validated and added to the syntax tree
    '----------------------------------------------------------------------------------
    GoTo SkipWord
    
    
    '----------------------------------------------------------------------------------
    'Errors:
    '----------------------------------------------------------------------------------
WordError:
    '..................................................................................
    'Not a recognised word (typo perhaps?)
    Call RaiseError( _
        OZ80_ERROR_INVALID, Line, Col, _
        "Encountered an unknown word; it is probably a typo" _
    )
    GoTo Finish
    
ErrorOverflow:
    '..................................................................................
    On Error GoTo 0
    Call RaiseError( _
        OZ80_ERROR_OVERFLOW, Line, Col, _
        "Overflow" _
    )
    GoTo Finish
    
    '----------------------------------------------------------------------------------
    
Finish:
    Call Source.CharArrayUnaccess(Buffer)
    Call Source.Destroy
    Set Source = Nothing
End Sub

'AssembleFile_SyntaxTree : Convert a Token Stream into an Abstract Syntax Tree
'======================================================================================
'TokenStream    | The Token Stream to build a Syntax Tree from
'======================================================================================
Private Sub AssembleFile_SyntaxTree( _
    ByRef TokenStream As oz80TokenStream _
)   'This function is tuned for speed -- it does not conform to "best practices", _
     or even structured code; GOTOs are fast, deal with it
    
    'The current Token being processed
    Dim Token As oz80Token
    'When a Token contains a string, it'll be made available here
    Dim TokenText As bluString
    Set TokenText = New bluString
    
    'As we join the nodes in the Syntax Tree, we'll need to keep track of the _
     previous node so that we can link together a continuous chain
    Dim NodePrev As Long
    'When a particular branch ends, we want to be able to quickly get back to the _
     branch-point without having to traverse every node. We keep track of the last _
     parent-node as we go for this reason
    Dim NodeParent As Long
    'As we add nodes in a List, we must link the new node to the previous, though _
     we must avoid linking the first node (child) with its previous (parent); _
     ergo: when NodeSibling = 0, we do not need to link the node
    Dim NodeSibling As Long
    
    'Our current position within the Token Stream
    Dim TokenIndex As Long
    'Seed the routine with the first Token in the Stream
    GoSub GetToken
    
    'WARNING: Severe abuse of `Gosub`/`Return` follows. I'll provide a comment _
     any time I use `Goto` with the intention of using the destination's `Return`
    
ScopeRoot:
    '----------------------------------------------------------------------------------
    '[A]    SCOPE <root>
    '----------------------------------------------------------------------------------
    'Upon the start of a file, only a specific set of statements are allowed
    
    '..................................................................................
    '[A.1]  `CONST !<name> <expr>`
    '..................................................................................
    If Token.Kind = TOKEN_KEYWORD_CONST Then
        'Add the `CONST` keyword to the syntax tree
        GoSub AddNodeSibling
        
        'The next word must be a Constant name
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_CONST Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "A Constant name must follow the `CONST` statement. A Constant name " _
          & "begins with an exclamation mark; for example: " & vbNewLine _
          & vbNewLine _
          & vbTab & "CONST !PROBLEMS 99" _
        ): GoTo Finish
        
        'Predefine the Constant
        Let TokenText.Text = TokenStream.Strings(Token.Value)
        GoSub PredefineName
        
        'Add the constant name to the tree
        GoSub AddNodeParam
        
        'An optional Expression follows to define the value:
        If (Token.Attr And ATTR_EXPR) <> 0 Then
            'Prepare the empty node as a parameter of the previous, _
             the Expression scope will fill in the details
            GoSub AddParam
            'Add the Expression to the parameter branch
            GoSub ScopeExpr
            'Leave the parameter branch and return to the constant name
            GoSub UpScope
        End If
        
        'Leave the parameter branch (Constant name) and return to root
        GoSub UpScope
        
    '..................................................................................
    '[A.2]  `INCLUDE <list>`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_INCLUDE Then
        'Add this word to the syntax tree
        GoSub AddNodeSibling
        
        'The List that follows (the file name) is not optional
        If (Token.Attr And ATTR_LIST) = 0 Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "The `INCLUDE` statement must be followed by a List " _
          & "(the file to include). For example: " & vbNewLine _
          & vbNewLine _
          & vbTab & "INCLUDE ""library.oz80""" _
        ): GoTo Finish
        
        'Prepare a parameter branch
        GoSub AddParam
        'Look for the list that follows
        GoSub ScopeList
        'Leave the parameter branch and return to root
        GoSub UpScope
        
    '..................................................................................
    '[A.3]  `SYSTEM <list>
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_SYSTEM Then
        'Add this word to the syntax tree
        GoSub AddNodeSibling
        
        'There must be a List following the `SYSTEM` statement
        If (Token.Attr And ATTR_LIST) = 0 Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "The `SYSTEM` statement must be followed by a List (the target System " _
          & "name). See the documentation for supported target Systems" _
        ): GoTo Finish
        
        'Prepare a parameter branch
        GoSub AddParam
        'Look for the List that follows
        GoSub ScopeList
        'Leave the parameter branch and return to root
        GoSub UpScope
        
    '..................................................................................
    '[A.4]  `SECTION ::<name> [SLOT <list>]`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_SECTION Then
        'Add this word to the syntax tree
        GoSub AddNodeSibling
        
        'The next word must be a Section name. This checks for a Section only, to _
         avoid accepting fully-qualified Labels, .e.g. "::section:label.property"
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "A Section name must follow the `SECTION` statement. " _
          & "Section names begin with two colons, e.g. ""::stuff""" _
        ): GoTo Finish
        
        'Add the Section name to the tree
        GoSub AddNodeParam
        
        'Check for the optional `SLOT` parameter
        If Token.Kind = TOKEN_KEYWORD_SLOT Then
            'Add the `SLOT` keyword to the tree
            GoSub AddNodeSibling
            
            'The Slot pattern that follows is not optional. We won't validate the _
             specific Slot pattern here as it could be dynamic and is System dependent
            If (Token.Attr And ATTR_LIST) = 0 Then Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "A List must follow the `SLOT` parameter of the `SECTION` statement. " _
              & "The format of the Slot List is dependent on the target System, " _
              & "so view the documentation for specifics" _
            ): GoTo Finish
            
            'Prepare a parameter branch
            GoSub AddParam
            'Fetch the Slot patten following
            GoSub ScopeList
            'Leave the parameter branch and return to keyword param
            GoSub UpScope
        End If
        
        'Leave the parameter branch and return to root
        GoSub UpScope
        
    '..................................................................................
    '[A.5]  `TYPE #<name> [...]`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_TYPE Then
        'Add this keyword to the syntax tree
        GoSub AddNodeSibling
        
        'The next word must be a Type name
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_TYPE Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected Type name following `TYPE` statement. " _
          & "See the documentation on how to define Types" _
        ): GoTo Finish
        
        'Add the type name as parameter to the `TYPE` keyword
        GoSub AddNodeParam
        'Leave the param branch...;
        GoSub UpScope
        'And walk the Type definition block
        GoSub ScopeType
        
    '..................................................................................
    '[A.6]  `RAMPAGE [$_<name>] [<address>] [<length>]`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_RAMPAGE Then
        'Add the `RAMPAGE` keyword to the tree
        GoSub AddNodeSibling
        
        'The first parameter is the RAM Page name
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_RAM Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected RAM Page name" _
        ): GoTo Finish
        
'        'Check to see if the RAM Name has already been pre-allocated
'        Dim RAMPage As Long
'        For i = 1 To RAMPagesCount - 1
'            If Word.Equals(RAMPages(i).Type.Name) Then Let RAMPage = i: Exit For
'        Next i
'
'        If RAMPage = 0 Then
'            'Pre-allocate the RAM Page
'            Let RAMPagesCount = RAMPagesCount + 1
'            ReDim Preserve RAMPages(0 To RAMPagesCount) As oz80RAMPage
'            Let RAMPages(RAMPagesCount).Type.Name = Word.Text
'            Let WordValue = RAMPagesCount
'        Else
'            Let WordValue = RAMPage
'        End If
        
        'Add the RAM Page name to the syntax tree
        GoSub AddNodeParam
        
        'The first optional parameter is the beginning address ($0 - $FFFF)
        If (Token.Attr And ATTR_EXPR) <> 0 Then
            'Create a parameter branch in the syntax tree
            GoSub AddNodeParam
            'Add the Expression
            GoSub ScopeExpr
            
            'The second optional parameter is the length of the page (1 B to 16 KB)
            If (Token.Attr And ATTR_EXPR) <> 0 Then
                'Create a parameter branch in the syntax tree
                GoSub AddNodeParam
                'Add the Expression
                GoSub ScopeExpr
                'Leave the parameter branch for the second optional param
                GoSub UpScope
            End If
            
            'Leave the parameter branch for the first optional param
            GoSub UpScope
        End If
        
        'Leave the parameter branch (RAM Page name)
        GoSub UpScope
        
    '..................................................................................
    '[A.7]  `RAM [$_<name>]`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_RAM Then
        'Add the `RAM` keyword to the tree
        GoSub AddNodeSibling
        
        'The parameter must be RAM Page name, not a fully-qualified RAM name
        If (Token.Attr And ATTR_NAMES) = ATTR_NAME_RAM Then
            'The RAM page name will be the parameter of the keyword
            GoSub AddNodeParam
            'Leave the parameter branch
            GoSub UpScope
        End If
        
        'As an aide, anything other than the opening square bracket _
         can provide a more specific error message
        If Token.Kind <> TOKEN_SQUARE_OPEN Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected either a RAM Page name, or an opening square bracket. " _
          & "The `RAM` keyword can optionally be followed by a RAM Page name " _
          & "and then must be followed by the opening square bracket of a RAM block" _
          & "definition" _
        ): GoTo Finish
        
        'Process the RAM block that follows
        GoSub ScopeType
        
    '..................................................................................
    '[A.8]  `START {[...]}`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_START Then
        'Add the `START` keyword to the tree
        GoSub AddNodeSibling
        'Process the code Chunk that follows "{...}"
        GoSub ScopeCode
        
    '..................................................................................
    '[A.9]  `TABLE :<name> [SECTION ::<name>] {[...]}`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_TABLE Then
        'Add the `TABLE` keyword to the tree
        GoSub AddNodeSibling
        
        'The Label name follows (cannot be anonymous)
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_LABEL Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected a Label name following the `TABLE` keyword. " _
          & "The Label name must not include a Section or Property component." _
        ): GoTo Finish
        
        'Add the Label name as a parameter
        GoSub AddNodeParam
        
        'Check for the optional Section name
        If Token.Kind = TOKEN_KEYWORD_SECTION Then
            'Add the `SECTION` keyword to the tree
            GoSub AddNodeSibling
            
            'A Section name must follow
            If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "Expected a Section name" _
            ): GoTo Finish
            
            'Add the Section name as parameter
            GoSub AddNodeParam
            GoSub UpScope
        End If
        
        'Process the table's guts
        GoSub UpScope
        GoSub ScopeCode
        
    '..................................................................................
    '[A.10] `PROC :<name> [SECTION ::<name> | INTERRUPT <expr>] [HELP <text>] _
                          [PARAMS <paramlist>] [RETURN <paramlist>]
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_PROC Then
        'Add this keyword to the tree
        GoSub AddNodeSibling
        
        'The Label name that follows (cannot be anonymous)
        If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_LABEL Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected a Label name following the `PROC` keyword. " _
          & "The Label name must not include a Section or Property component." _
        ): GoTo Finish
        
        'Add the Label as parameter to the keyword, _
         the `PROC` params that follow will be siblings of this
        GoSub AddNodeParam
        
ProcParams:
        'The `PROC` statement has a number of optional parameters, and not all can _
         be combined. We keep track of which we've encountered and avoid duplicates
        Dim IsProcSectionDefined As Long
        Dim IsProcInterruptDefined As Long
        Dim IsProcHelpDefined As Long
        Dim IsProcParamsDefined As Long
        Dim IsProcReturnDefined As Long
        
        '`PROC ... SECTION ::<name>`
        '..............................................................................
        If Token.Kind = TOKEN_KEYWORD_SECTION Then
            'The Section parameter cannot be defined twice!
            If IsProcSectionDefined <> 0 Then Call RaiseError( _
                OZ80_ERROR_DUPLICATE, Token.Line, Token.Col, _
                "`SECTION` parameter defined twice!" _
            ): GoTo Finish
            
            'Add the `SECTION` keyword to the tree
            GoSub AddNodeSibling
            
            'A Section name must follow
            If (Token.Attr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "Expected a Section name" _
            ): GoTo Finish
            
            'Add the Section name as parameter to the `SECTION` keyword
            GoSub AddNodeParam
            GoSub UpScope
            
            'Now mark Section parameter defined, _
             and go to the next parameter
            Let IsProcSectionDefined = 1
            GoTo ProcParams
            
        '`PROC ... INTERRUPT <expr>`
        '..............................................................................
        ElseIf Token.Kind = TOKEN_KEYWORD_INTERRUPT Then
            'The Interrupt parameter cannot be defined twice!
            If IsProcInterruptDefined <> 0 Then Call RaiseError( _
                OZ80_ERROR_DUPLICATE, Token.Line, Token.Col, _
                "`INTERRUPT` parameter defined twice!" _
            ): GoTo Finish
            
            'Add the `INTERRUPT` keyword
            GoSub AddNodeSibling
            
            'An Expression must follow
            If (Token.Attr And ATTR_EXPR) = 0 Then Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "Expected an Expression" _
            ): GoTo Finish
            
            'Add the Expression as parameter to the keyword
            GoSub AddParam
            GoSub ScopeExpr
            GoSub UpScope
            
            'Now mark Interrupt parameter defined, _
             and go to the next parameter
            Let IsProcInterruptDefined = 1
            GoTo ProcParams
            
        '`PROC ... HELP ``...`
        '..............................................................................
        ElseIf Token.Kind = TOKEN_KEYWORD_HELP Then
            'The Help parameter cannot be defined twice!
            If IsProcHelpDefined <> 0 Then Call RaiseError( _
                OZ80_ERROR_DUPLICATE, Token.Line, Token.Col, _
                "`HELP` parameter defined twice!" _
            ): GoTo Finish
            
            'Add the `HELP` keyword to the tree
            GoSub AddNodeSibling
            
            'Documentation text must follow
            If Token.Kind <> TOKEN_DOC Then Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "Expected documentation text following the `HELP` keyword" _
            ): GoTo Finish
            
            'The text will be parameter of the `HELP` keyword
            GoSub AddNodeParam
            
            'Now loop and add documentation text
            Do While Token.Kind = TOKEN_DOC
                GoSub AddNodeSibling
            Loop
            
            'Leave the documentation text parameter branch
            GoSub UpScope
            
            'Now mark Help parameter defined, _
             and go to the next parameter
            Let IsProcHelpDefined = 1
            GoTo ProcParams
            
        '`PROC ... PARAMS <params>`
        '..............................................................................
        ElseIf Token.Kind = TOKEN_KEYWORD_PARAMS Then
            'The `PARAMS` parameter cannot be defined twice!
            If IsProcParamsDefined <> 0 Then Call RaiseError( _
                OZ80_ERROR_DUPLICATE, Token.Line, Token.Col, _
                "`HELP` parameter defined twice!" _
            ): GoTo Finish
            
            'Add the `PARAMS` keyword to the tree, _
             and prepare a branch for the following list
            GoSub AddNodeSibling
            GoSub AddParam
            
            GoSub ScopeProcParams
            'Leave the parameter branch
            GoSub UpScope
            
            'Now mark `PARAMS` parameter defined, _
             and go to the next parameter
            Let IsProcParamsDefined = 1
            GoTo ProcParams
            
        '`PROC ... RETURN <params>`
        '..............................................................................
        ElseIf Token.Kind = TOKEN_KEYWORD_RETURN Then
            'The `RETURN` parameter cannot be defined twice!
            If IsProcReturnDefined <> 0 Then Call RaiseError( _
                OZ80_ERROR_DUPLICATE, Token.Line, Token.Col, _
                "`HELP` parameter defined twice!" _
            ): GoTo Finish
            
            'Add the `RETURN` keyword to the tree, _
             and prepare a branch for the following list
            GoSub AddNodeSibling
            GoSub AddParam
            GoSub ScopeProcParams
            GoSub UpScope
            
            'Now mark `RETURN` parameter defined, _
             and go to the next parameter
            Let IsProcReturnDefined = 1
            GoTo ProcParams
            
        '`PROC ... { ... }`
        '..............................................................................
        ElseIf Token.Kind = TOKEN_BRACES_OPEN Then
            'Leave the parameter branch, _
             the code Chunk will attach to the `PROC` keyword
            GoSub UpScope
            GoSub ScopeCode
            
        'ERROR: Not a valid `PROC` parameter
        '..............................................................................
        Else
            Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "TODO" _
            )
            GoTo Finish
        End If
        
        'Reset the param checks for the next `PROC`
        Let IsProcSectionDefined = 0
        Let IsProcInterruptDefined = 0
        Let IsProcHelpDefined = 0
        Let IsProcParamsDefined = 0
        Let IsProcReturnDefined = 0
        
    '..................................................................................
    Else
        'ERROR - Only valid statements allowed at root
        Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Unexpected word at the root scope (that is, statements not within " _
          & "other statement blocks such as `PROC`). This could be that you have " _
          & "one stray closing-brace '}' causing another scope to end incorrectly. " _
          & "it could also be a simple typo; only the following words are allowed " _
          & "in the root scope:" & vbNewLine _
          & vbNewLine _
          & vbTab & "CONST, INCLUDE, PROC, RAM, SECTION, START, SYSTEM, " _
                  & "TABLE, TYPE" _
        )
        GoTo Finish
    End If
    GoTo ScopeRoot
    
ScopeProcParams:
    '----------------------------------------------------------------------------------
    '[B]    SCOPE <params>
    '----------------------------------------------------------------------------------
    'Procedures (`PROC`), allow the documentation of the parameters they accept. _
     Because Procedures are assembly code, parameter use is not enforced; they are _
     provided purely as a programmer aid since it may not be obvious how the stack / _
     registers must be set up for any given Provedure. The `PARAMS` / `RETURN` _
     keywords accept a very limited list, applicable to assembly use
    
    'Once a `PROC` param list is exepected, it's not optional
    If (Token.Attr And ATTR_PARAM) = 0 Then Call RaiseError( _
        OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
        "Expected a `PROC` parameter: That is, a Z80 Register or a RAM Name" _
    ): GoTo Finish
    
    'As long as there are params, keep going:
    Do While (Token.Attr And ATTR_PARAM) <> 0
        'Add the param to the list
        GoSub AddNodeSibling
        'Check for documentation text to describe the param
        If Token.Kind = TOKEN_DOC Then
            'Documentation text is the prameter (syntax tree) of the `PROC` param
            GoSub AddParam
            'Allow multiple documentation texts in a row
            Do While Token.Kind = TOKEN_DOC: GoSub AddNodeSibling: Loop
            'Leave the parameter branch (for the documentation text) _
             and return to the `PROC` param list
            GoSub UpScope
        End If
    'Check for the next param in the list
    Loop
    'Once the list ends, return
    Return
    
ScopeCode:
    '----------------------------------------------------------------------------------
    '[C]    SCOPE <code>
    '----------------------------------------------------------------------------------
    'All assembly blocks must begin with an opening curly brace
    If Token.Kind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
        "Expected opening curly-brace" _
    ): GoTo Finish
    
    'Begin a child branch in the syntax tree
    GoSub AddChild
    
    'Continue walking the assembly block so long as the block remains open
    Do While Token.Kind <> TOKEN_BRACES_CLOSE
        
        'A Z80 instuction:
        If (Token.Attr And ATTR_Z80_MNEMONIC) <> 0 Then
            'TODO: To what extent we should validate Z80 instructions and their _
                   parameters before assembly I am not certain. The `Z80Assemble` _
                   function does this for us as part of necessary work to identify _
                   what opcodes are needed. For simplicity we will remain dumb here _
                   and look at shifting responsibility sometime in the future
            
            'Add the instruction mnemonic to the tree
            GoSub AddNodeSibling
            'TODO: No params?
            'Begin a parameter branch from the instruction
            GoSub AddParam
            'Add the Z80 parameters
            GoSub ScopeZ80Param
            'Leave the parameter branch, _
             returning to the Z80 instruction
            GoSub UpScope
            
        'A binary list: (Bytes, Words, Trips, Longs &c)
        ElseIf (Token.Attr And ATTR_UNIT) <> 0 Then
            'Add the binary type keyword
            GoSub AddNodeSibling
            'Begin a parameter branch for the following list
            GoSub AddParam
            'Walk the list
            'TODO: First element cannot be optional
            GoSub ScopeList
            'Leave the parameter branch
            GoSub UpScope
            
        'Define a sub-routine:
        ElseIf Token.Kind = TOKEN_PROPERTY_NEW Then
            'Add the property to the tree
            GoSub AddNodeSibling
            
        'Define a Table Index:
        ElseIf Token.Kind = TOKEN_KEYWORD_INDEX Then
            GoSub AddNodeSibling
            GoSub AddNodeParam
            GoSub UpScope
            
        Else
            Call RaiseError( _
                OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
                "U WOT M8?" _
            )
            GoTo Finish
        End If
    Loop
    
    GoSub UpScope
    GoSub NextToken
    Return
    
ScopeZ80Param:
    '----------------------------------------------------------------------------------
    '[D]    SCOPE <z80param>
    '----------------------------------------------------------------------------------
    'Memory reference?
    If Token.Kind = TOKEN_SQUARE_OPEN Then
        GoSub AddChild
        
        'IX or IY?
        If (Token.Kind = TOKEN_Z80_IX) _
        Or (Token.Kind = TOKEN_Z80_IY) _
        Then
            GoSub AddNodeSibling
            
            'Optional "+"?
            If Token.Kind = TOKEN_OPERATOR_ADD Then
                GoSub AddNodeParam
                GoSub AddParam
                'Expression follows (not optional, given the plus)
                GoSub ScopeExpr
                GoSub UpScope
                GoSub UpScope
            End If
            
        ElseIf (Token.Attr And ATTR_Z80_REGISTER) <> 0 Then
            GoSub AddNodeSibling
            
        ElseIf (Token.Attr And ATTR_EXPR) <> 0 Then
            GoSub ScopeExpr
            
        Else
            Call RaiseError( _
                OZ80_ERROR_INTERNAL, Token.Line, Token.Col, _
                "That's not something I know" _
            ): GoTo Finish
        End If
        
        If Token.Kind = TOKEN_SQUARE_CLOSE Then
            GoSub UpScope: GoSub NextToken
        Else
            Stop
        End If
        
    ElseIf (Token.Attr And ATTR_EXPR) <> 0 Then
        GoSub ScopeExpr
        
    ElseIf (Token.Attr And ATTR_Z80_REGFLAG) <> 0 Then
        'Add the register to the z80-list
        GoSub AddNodeSibling
        
    Else
        Return
    End If
    GoTo ScopeZ80Param
    
ScopeType:
    '----------------------------------------------------------------------------------
    '[E]    SCOPE <type>
    '----------------------------------------------------------------------------------
    
    'The Type definition must open with square brackets
    If Token.Kind <> TOKEN_SQUARE_OPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
        "Type definition block expected" _
    ): GoTo Finish
    
    GoSub AddChild
    
    'TODO: Could be empty
    Do
        'Each definition within the Type begins with _
         the Property name or the anonymous Property `.`
        If (Token.Kind <> TOKEN_PROPERTY_ANON) _
        And (Token.Kind <> TOKEN_PROPERTY_NEW) _
        Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "Expected a Property" _
        ): GoTo Finish
        
        GoSub AddNodeSibling
        
        'The unit size should follow
        If (Token.Attr And ATTR_UNIT) = 0 Then Call RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col, _
            "TODO!" _
        ): GoTo Finish
        'If it's `TYPE`, then either a Type name will follow, _
         or a Type definition block
        If Token.Kind = TOKEN_KEYWORD_TYPE Then
            'Add the `TYPE` unit to the tree
            GoSub AddNodeParam
            'Either a Type name follows or a Type definition block
            If (Token.Attr And ATTR_NAMES) = ATTR_NAME_TYPE Then
                GoSub AddNodeParam
                GoSub UpScope
            Else
                'Now process the Type-within-a-Type
                GoSub ScopeType
            End If
            GoSub UpScope
        Else
            'Add the Unit size to the tree
            GoSub AddNodeParam
            'If the repeat operator `x` follows, define the block size
            If Token.Kind = TOKEN_OPERATOR_REP Then
                GoSub AddNodeParam
                GoSub AddParam
                'TODO: Not optional
                GoSub ScopeExpr
                GoSub UpScope
                GoSub UpScope
            End If
            'TODO: Default Value
            GoSub UpScope
        End If
        
    Loop While Token.Kind <> TOKEN_SQUARE_CLOSE
    
    GoSub UpScope
    GoSub NextToken
    Return
    
ScopeList:
    '----------------------------------------------------------------------------------
    '[F]    SCOPE <list>
    '----------------------------------------------------------------------------------
    '[F.1]  Text is a list of its own, but not an Expression _
            (it's a series of bytes; more than one Value)
    '..................................................................................
    If Token.Kind = TOKEN_TEXT Then
        'Prepare the information for this
'        Let Tree(TreeNodeCount).Value = TokenStrings.Add(Word.Text)
        'TODO: Add the Text as a sibling in the List
        GoSub AddNodeSibling
        'Check for the next item in the list...
        GoTo ScopeList
        
    '[F.2]  The List scope also encompases the Expression scope, so we cannot _
            consider the List scope finished until we check for Expressions too
    '..................................................................................
    ElseIf (Token.Attr And ATTR_EXPR) <> 0 Then
        'Go fetch an Expression; we return here because _
         a List may comprise more than one Expression
        GoSub ScopeExpr
        'Now check again for another List item
        GoTo ScopeList
    End If
    'When no List item is found, _
     fold up to the parent scope
    Return
    
ScopeExpr:
    '----------------------------------------------------------------------------------
    '[G]    SCOPE <expr>
    '----------------------------------------------------------------------------------
    'When a non-Expression token occurs, the Expression ends and it's _
     up to the parent Scope to determine if the token is valid
    If (Token.Attr And ATTR_EXPR) = 0 Then
        Return
        
    '[G.1]  Defined Name, `!CONST`, `$_RAM`, `::Section`, `:Label` etc
    '..................................................................................
    ElseIf Token.Kind = TOKEN_NAME Then
        'TODO:
        GoSub AddNodeSibling
        
    '[G.2]  `HI <expr>`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_HI Then
        'TODO: `HI` function has a parameter Expression
        GoSub AddNodeSibling
        GoSub AddParam
        GoSub ScopeExpr
        GoSub UpScope
        
    '[G.3]  `LO <expr>`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_LO Then
        'TODO: `LO` function has a parameter Expression
        GoSub AddNodeSibling
        GoSub AddParam
        GoSub ScopeExpr
        GoSub UpScope
        
    '[G.4]  `START`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_KEYWORD_START Then
        'TODO
        GoSub AddNodeSibling
        
    '[G.5]  `255`, `$FF`, `%11111111`
    '..................................................................................
    ElseIf Token.Kind = TOKEN_NUMBER Then
        'TODO
        GoSub AddNodeSibling
        'Check for number prefix
        'TODO: This should be decimal only
        If (Token.Kind = TOKEN_SUFFIX_K) _
        Or (Token.Kind = TOKEN_SUFFIX_KB) Then
            GoSub AddNodeParam
            GoSub UpScope
        End If
        
    '[G.6]  `(` ...
    '..................................................................................
    ElseIf Token.Kind = TOKEN_PARENS_OPEN Then
        'TODO
        GoSub AddChild
        GoSub ScopeExpr
        GoSub UpScope
        
        'TODO: The parens should end here
        GoSub NextToken
        
    Else
        MsgBox "This token is classified as an Expression, but it not handled"
        Stop
    End If
    
    'To simplify a lot of routine, any Expression Value can be followed by an _
     Operator; extending the Expression. For a nice bit of optimisation, we can _
     simply drop through to the Operator scope below and borrow its `Return`!
    
ScopeOperator:
    '----------------------------------------------------------------------------------
    '[H]    SCOPE {operator}
    '----------------------------------------------------------------------------------
    'TODO: Repeat Operator `x` is not wanted here, needs to be reclassified
    If (Token.Attr And ATTR_OPERTATOR) <> 0 Then
        GoSub AddNodeParam
        GoSub AddParam
        GoSub ScopeExpr
        GoSub UpScope
        GoSub UpScope
    End If
    Return
    
    
    '[2.2]  Abstract Syntax Tree support routines:
    '----------------------------------------------------------------------------------
    'This variable keeps track of how deeply-nested our syntax tree is. If this isn't _
     zero when we hit the end of the file then there's an error with the user's code
    Dim ScopeDepth As Long
    
UpScope:
    'When a particular branch of the tree comes to an end, trace back
    '..................................................................................
    Let NodePrev = NodeParent
    If NodePrev <> 0 _
        Then Let NodeParent = Tree.Nodes(NodePrev).NodeParent _
        Else Let NodeParent = 0
    Let NodeSibling = NodePrev
    Let ScopeDepth = ScopeDepth - 1
    Return
    
AddChild:
    'This node is the child of the previous node:
    '..................................................................................
    'Set the length of the log message according to the scope depth
    Let ScopeDepth = ScopeDepth + 1
    'Display the current log line, as this node will appear on the next
'    Call RaiseMessage(OZ80_LOG_STATUS, LogStr.Text)
    'Prepare the next line
'    Call LogStr.Clear.Left(20 * (ScopeDepth + 1))
    
    If NodePrev <> 0 Then
        'Link the previous node with this one
        Let Tree.Nodes(NodePrev).NodeChild = Tree.Count
    End If
    'The parent node is now the previous _
     (the previous remains the same)
    Let NodeParent = NodePrev
    'This is the first sibling, don't link the parent & child as siblings!
    Let NodeSibling = 0
    'We do not need to add a new node, the current one will be used. _
     This is because braces and brackets are implied by the tree structure _
     and do not need to be explicitly added as a node
    GoTo NextToken
    
AddParam:
    'Makes the current (unset) node a param of the previous
    '..................................................................................
    'Set the length of the log message according to the scope depth
    Let ScopeDepth = ScopeDepth + 1
'    Call LogStr.Left(20 * (ScopeDepth + 1))
    
    If NodePrev <> 0 Then
        'Link the previous node with this one
        Let Tree.Nodes(NodePrev).NodeParam = Tree.Count
    End If
    'The parent node is now the previous _
     (the previous remains the same)
    Let NodeParent = NodePrev
    'This is the first sibling, don't link the parent & param as siblings!
    Let NodeSibling = 0
    Return
    
AddNodeParam:
    'This node is a parameter of the previous node:
    '..................................................................................
    'Set the length of the log message according to the scope depth
    Let ScopeDepth = ScopeDepth + 1
'    Call LogStr.Left(20 * (ScopeDepth + 1))
    
    If NodePrev <> 0 Then
        'Link the previous node with this one
        Let Tree.Nodes(NodePrev).NodeParam = Tree.Count
    End If
    'The previous node is now the parent node
    Let NodeParent = NodePrev
    'This is the first sibling, don't link the parent & param as siblings!
    Let NodeSibling = Tree.Count
    GoTo AddNode
    
AddNodeSibling:
    'This node is the sibling of the previous node:
    '..................................................................................
    If NodeSibling <> 0 Then
        'Display the current log line, as this node will appear on the next
'        Call RaiseMessage(OZ80_LOG_STATUS, LogStr.Text)
        'Prepare the next line
'        Call LogStr.Clear.Left(20 * (ScopeDepth + 1))
        
        'Link the previous node with this one
        Let Tree.Nodes(NodeSibling).NodeNext = Tree.Count
    End If
    'This will now be the sibling to link to
    Let NodeSibling = Tree.Count
    Let NodePrev = Tree.Count
    
AddNode:
    '..................................................................................
    'Add the current word to the syntax tree
    With Tree.Nodes(Tree.Count)
        Let .File = File
        Let .Token = Token
        Let .NodeParent = NodeParent
    End With
    
    'Formulate a log message:
    
'    'Numerical literals:
'    If WordToken = TOKEN_NUMBER Then
'        'Display the original text from the source file, since this will retain _
'         the formatting and base-type of the number (binary / decimal / hexadecimal)
'        Call LogStr.Overwrite(20 * ScopeDepth, Word.Truncate(19).Text)
'    'Defined Names:
'    ElseIf WordToken = TOKEN_NAME Then
'        Call LogStr.Overwrite(20 * ScopeDepth, Word.Truncate(19).Text)
'    'Property definitions:
'    ElseIf WordToken = TOKEN_PROPERTY_NEW Then
'        Call LogStr.Overwrite(20 * ScopeDepth, Word.Truncate(19).Text)
'    'Quoted text:
'    ElseIf WordToken = TOKEN_TEXT Then
'        Call LogStr.Overwrite(20 * ScopeDepth, Word.Truncate(17).Wrap("""").Text)
'    'Documentation text:
'    ElseIf WordToken = TOKEN_DOC Then
'        Call LogStr.Overwrite(20 * ScopeDepth, Word.Truncate(19).Text)
'    'For all other tokens, used the cached names
'    Else
'        Call LogStr.Overwrite(20 * ScopeDepth, TokenNames(WordToken))
'    End If
    
    'This node is now the reference
    Let NodePrev = Tree.Count
    
    'Define a new node
    Let Tree.Count = Tree.Count + 1
    If Tree.Count > Tree.Buffer Then
        Let Tree.Buffer = Tree.Buffer * 2
        ReDim Preserve Tree.Nodes(1 To Tree.Buffer) As oz80TreeNode
    End If
    
    'Now that word has been filed, pre-fetch the next one for processing. _
     NOTE: `Goto` instead of `Gosub` has been used here so that the `Return` in _
           `ReadWord` will return from here for us
'    GoTo ReadWord
    
    '----------------------------------------------------------------------------------
    
NextToken:
    Let TokenIndex = TokenIndex + 1
    'TODO: Handle unexpected end of stream
    If TokenIndex = TokenStream.Length Then GoTo Finish
GetToken:
    'Fetch the current Token Stream token; this will contain the token type, _
     attributes and source file line & column number, which will be passed on _
     to the Syntax Tree node for inclusion
    Let Token = TokenStream.Tokens(TokenIndex)
    
    'Each Token Stream has its own database of strings; one Token Stream represents _
     only one source code file. The Syntax Tree however combines all source files _
     used and deduplicates strings in its own string database so that every text node _
     points to a coninical string. This allows us to refer to strings with IDs _
     avoiding a lot of string searching / comparison
    Dim i As Long
    
'    'Is the token a Named Value? (e.g. Constant, Label, &c.)
'    '..................................................................................
'    Dim i As Long
'    If Token.Kind = TOKEN_NAME Then
'        'Take a copy of the text from the Token Stream's string database
'        Let TokenText.Text = TokenStream.Strings(Token.Value)
'
'        'TODO: Handle Scope; e.g. ::SECTION > :LABEL > .PROPERTY
'
'        Let i = 0
'        For i = 1 To NameCount
'            If TokenText.Equals(Names(i).Name) Then
'                Let Token.Value = i: Let i = 0: Exit For
'            End If
'        Next i
'        If i <> 0 Then
'            Let NameCount = NameCount + 1
'            ReDim Preserve Names(1 To NameCount) As oz80Name
'            With Names(NameCount)
'                Let .Name = TokenText.Text
'                Let .Value = INDEFINITE
'            End With
'            Let Token.Value = NameCount
'        End If
        
    'If it's a string of some kind, fetch the associated text
    '..................................................................................
    If Token.Kind = TOKEN_TEXT _
        Or Token.Kind = TOKEN_DOC Then
        'Take a copy of the text from the Token Stream's string database
        Let TokenText.Text = TokenStream.Strings(Token.Value)
        
        'Search through the Syntaxt Tree's string database for this token's string
        Let i = 0
        For i = 1 To Tree.StringCount
            'If the string is found...
            If TokenText.Equals(Tree.Strings(i)) Then
                'Re-use the existing string index
                Let Token.Value = i
                Let i = 0: Exit For
            End If
        Next i
        'If no matching string were found...
        If i <> 0 Then
            'Add the string to the end of the Syntaxt Tree string database
            Let Tree.StringCount = Tree.StringCount + 1
            ReDim Preserve Tree.Strings(1 To Tree.StringCount) As String
            Let Tree.Strings(Tree.StringCount) = TokenText.Text
            'The current token will be the basis of the same information stored _
             in the syntaxt tree. The Value will be the ID of the string
            Let Token.Value = Tree.StringCount
        End If
    End If
    Return
    
PredefineName:
    Let i = 0
    For i = 1 To NameCount
        If TokenText.Equals(Names(i).Name) Then
            Let Token.Value = i: Let i = 0: Exit For
        End If
    Next i
    If i <> 0 Then
        Let NameCount = NameCount + 1
        ReDim Preserve Names(1 To NameCount) As oz80Name
        With Names(NameCount)
            Let .Name = TokenText.Text
            Let .Value = INDEFINITE
        End With
        Let Token.Value = NameCount
    End If
    Return
    
Finish:
    '[5]    Finalisation
    '==================================================================================
    Call TokenText.Destroy
    Set TokenText = Nothing
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

''FormatToken : Format the details of the current token for logging
''======================================================================================
''Returns        | A new bluString containing the details of the current token
''======================================================================================
'Private Function FormatToken( _
') As bluString
'    '*** TIMER BEGIN ***
'    Call Profiler.EventBegin(EVENT_FORMATTOKEN)
'
'    'Set the template we'll be putting the data into
'    Dim Template As New bluString
'    Let Template.Text = FORMAT_TOKEN
'
'    'Format the Token number
'    Dim Format_TokenIndex As New bluString
'    Let Format_TokenIndex.Text = Format$(Format$(TokenIndex + 1, "#,#"), "@@@@@@@")
'
'    Dim Format_Msg As New bluString
'    Dim Format_TokenType As String
'
'    If TokenKind = TOKEN_BRACES_OPEN _
'    Or TokenKind = TOKEN_BRACES_CLOSE Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_CHUNK
'
'    ElseIf TokenKind = TOKEN_PARENS_OPEN _
'        Or TokenKind = TOKEN_PARENS_CLOSE Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_EXPR
'
'    ElseIf TokenKind = TOKEN_SQUARE_OPEN _
'        Or TokenKind = TOKEN_SQUARE_CLOSE Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_SQUARE
'
'    ElseIf (TokenAttr And ATTR_NAME_CONST) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_CONST
'
'    ElseIf (TokenAttr And ATTR_Z80_FLAG) <> 0 Then
'        '.............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_FLAG
'
'    ElseIf (TokenAttr And ATTR_NAME_TYPE) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_SQUARE
'
'    ElseIf (TokenAttr And ATTR_Z80_MNEMONIC) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_Z80
'
'    ElseIf (TokenAttr And ATTR_KEYWORD) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Call Format_Msg.UCase
'        Let Format_TokenType = FORMAT_TOKEN_KEYWORD
'
'    ElseIf (TokenAttr And ATTR_NAME_SECTION) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_SECTION
'
'    ElseIf (TokenAttr And ATTR_NAME_LABEL) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_LABEL
'
'    ElseIf TokenKind = TOKEN_NUMBER Then
'        '..............................................................................
'        'TODO: Don't try and display an indefinite value
'        Let Format_Msg.Text = Format$(TokenValue, "#,0")
'        Let Format_TokenType = FORMAT_TOKEN_NUMBER
'
'    ElseIf (TokenAttr And ATTR_OPERTATOR) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_OPERATOR
'
'    ElseIf TokenKind = TOKEN_SUFFIX_K _
'        Or TokenKind = TOKEN_SUFFIX_KB Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_PREFIX
'
'    ElseIf TokenKind = TOKEN_PROPERTY_NEW _
'        Or (TokenAttr And ATTR_NAME_PROPERTY) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_PROPERTY
'
'    ElseIf (TokenAttr And ATTR_NAME_RAM) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_RAM
'
'    ElseIf (TokenAttr And ATTR_Z80_REGISTER) <> 0 Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenNames(TokenKind)
'        Let Format_TokenType = FORMAT_TOKEN_REGISTER
'
'    ElseIf TokenKind = TOKEN_TEXT Then
'        '..............................................................................
'        Let Format_Msg.Text = TokenText
'        Call Format_Msg.Wrap(FORMAT_QUOTE)
'        Let Format_TokenType = FORMAT_TOKEN_TEXT
'
'    Else
'        Debug.Print TokenNames(TokenKind)
'        Stop
'    End If
'
'    Set FormatToken = Template.Format( _
'        Format_TokenIndex, _
'        Right$("    " & TokenLine, 5), Left$(TokenCol & "  ", 3), _
'        Format_TokenType, _
'        Format_Msg _
'    )
'
'    '*** TIMER END ***
'    Call Profiler.EventEnd(EVENT_FORMATTOKEN)
'End Function

'/// PARSING //////////////////////////////////////////////////////////////////////////

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
'<root>   | CONST       !{name}    [<expr>]               ' Constant
'         |                                               '
'         | ENUM        !{name}    [<expr>]     <type>    ' Enumerate Constants
'         |                                               '
'         | SECTION     ::{name}   [SLOT        <list>]   ' Define Section
'         |                                               '
'         | SYSTEM      <text>                            ' Define System
'         |                                               '
'         | RAM         $#{name}    <expr>      <expr>    ' Define RAM Bank
'         |                                               '
'         | <if>                                          '
'         |                                               '
'         | INCLUDE     <text>                            '
'         |                                               '
'         | TYPE        #{name}     <type>                '
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

''ScopeRoot : When we're not within the scope of anything else
''======================================================================================
'Private Sub ScopeRoot()
'    'When a file begins, it is at the top-most scope. Certain keywords are only _
'     allowed here. Once all nested scopes are exhausted, we return here
'
'Continue:
'    'Define Constant:
'    '----------------------------------------------------------------------------------
'    If TokenKind = TOKEN_KEYWORD_CONST Then
'        'Defining a constant
'        Call DefineConst
'
'    'Defining the target System
'    ElseIf TokenKind = TOKEN_KEYWORD_SYSTEM Then
'        'Hand off to the function that will process this
'        Call DefineSystem
'
'    'Defining a code Section
'    ElseIf TokenKind = TOKEN_KEYWORD_SECTION Then
'        'Hand off to the function that will process this
'        Call DefineSection
'
'    'Define a Type:
'    '----------------------------------------------------------------------------------
'    '`TYPE #<type> [ ... ]`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_TYPE Then
'        'Hand off to the function that will process this
'        Call DefineType
'
'    'Inlcude another source file:
'    '----------------------------------------------------------------------------------
'    '`INCLUDE <text>`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_INCLUDE Then
'        'Hand off to the function that will process this
'        Call ProcessInclude
'
'    'Define a Procedure:
'    '----------------------------------------------------------------------------------
'    '`PROC :<label> { [...] }`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_PROC Then
'        'Hand off to the function that will process this
'        Call DefineProc
'
'    'Define a RAM Page:
'    '----------------------------------------------------------------------------------
'    '`RAMPAGE [$_<name>] [<expr>] [<expr>]`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_RAMPAGE Then
'        Call DefineRAMPage
'
'    'Create some RAM variables:
'    '----------------------------------------------------------------------------------
'    '`RAM [<expr>] [ ... ]`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_RAM Then
'        'Hand off to the function that will process this
'        Call DefineRAM
'
'    'Define the starting Procedure:
'    '----------------------------------------------------------------------------------
'    '`START {...}`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_START Then
'        'Hand off to the function that will process this
'        Call DefineStart
'
'    'Defing a data table:
'    '----------------------------------------------------------------------------------
'    '`TABLE :<name> {...}`
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_TABLE Then
'        'Hand off to the function that will process this
'        Call DefineTable
'
'    End If
'
'    If Error Then Exit Sub
'    'As we loop, other functions will move the Token pointer forward (such as when _
'     evaluating expressions). Check if we've hit the end of the file
'    If EOF Then Exit Sub
'    'As long as there is content left to process in the file, keep going
'    GoTo Continue
'End Sub

Private Sub AssembleRoot()
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
     
Continue:
    If Node.Token.Kind = TOKEN_KEYWORD_CONST Then
        'Move on to the keyword parameter
        Let Node = Tree.Nodes(Node.NodeParam)
    End If
    
End Sub

Private Sub TreeNodeNext()
    With Tree
        Let .Index = .Nodes(.Index).NodeNext
        Let Node = .Nodes(.Index)
    End With
End Sub

''DefineConst : Define a Constant value
''======================================================================================
''CONST      !<name>     [<expr>]
''======================================================================================
'Private Sub DefineConst()
'    If NextTokenError() Then Exit Sub
'
'    '`ScopeRoot` has already verified that the current token is a Constant name; _
'     check if the Constant has already been defined
'    If Values.Exists(TokenText) Then
'        'ERROR: You can't redefine Constants
'        Call RaiseError(OZ80_ERROR_DUPLICATE, "")
'        Exit Sub
'    End If
'
'    'Remember the Constant's name before moving past it
'    Dim ConstName As String
'    Let ConstName = TokenText
'
'    'The constant value is optional (it will default to 0), so the file can end here _
'     (it's up to the parent to check if there's a unclosed scope problem)
'    If EOF Then
'        'Define the Constant as zero
'        Call Values.Add(ConstName, 0)
'        Exit Sub
'    End If
'    Call NextToken
'
'    'Calculate the Expression's Value:
'    '(If there was an error in the Expression, the error event _
'      will already have been raised, we only need to fold upwards)
'    Dim Value As Double
'    Call ParseExpression(Value)
'
'    If Error Then Exit Sub
'
'    'Define the Constant:
'    Call Values.Add(ConstName, Value)
'End Sub
'
''DefineType : Define a Type block
''======================================================================================
''TYPE       #<name>     [ ... ]
''======================================================================================
'Private Sub DefineType()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
'    Dim TypeName As String
'    Let TypeName = TokenText
'
''    Call RaiseMessage(OZ80_LOG_STATUS, "TYPE {$}", TypeName)
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
''    Call RaiseMessage(OZ80_LOG_INFO, _
''                      "ADDR | UNIT | QTY   | NAME")
''    Call RaiseMessage(OZ80_LOG_INFO, _
''                      "-----¦------¦-------¦---------------------------------------------------")
'
'    'Skip the param
'    If NextTokenError() Then Exit Sub
'
'    'Create the new Type object
'    Let TypeCount = TypeCount + 1
'    ReDim Preserve Types(0 To TypeCount) As OZ80Type
'    Let Types(TypeCount).Name = TypeName
'
'    Call ScopeType(Types(TypeCount))
'
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE & vbNewLine)
'End Sub
'
''======================================================================================
''ScopeType : Walks a Type block and defines the structure within
''======================================================================================
''BaseType       | Which Type to populate
''[BaseOffset]   | Where the properties should begin addressing from
''[BaseName]     | The name heirarchy to prepend (when nesting Types)
''======================================================================================
'Private Sub ScopeType( _
'             ByRef BaseType As OZ80Type, _
'    Optional ByVal BaseOffset As Long = 0, _
'    Optional ByVal BaseName As String = vbNullString _
')
'    Dim Value As Double
'
'    'This is a recursive function designed for walking nested Type definitions. _
'     The outer-most Type block will already have been created in the Type database, _
'     and we will define all the properties within it
'
'    'It should be noted that the optional params above are `ByVal`, this is _
'     critically important otherwise recursive calls cause the array to be locked, _
'     preventing us from adding to the array whilst we walk it
'
'    'Begin with a simple validation -- the Type bracket
'    If TokenKind <> TOKEN_SQUARE_OPEN Then
'        Call RaiseError(OZ80_ERROR_EXPECTED, "TODO"): Exit Sub
'    End If
'    'Move over the bracket
'    If NextTokenError() Then Exit Sub
'
'Continue:
'    'Property Name:
'    '----------------------------------------------------------------------------------
'    With BaseType
'
'    Dim IsAnon As Long: Let IsAnon = 0
'    If TokenKind = TOKEN_PROPERTY_ANON Then Let IsAnon = 1
'
'    'Type definitions should begin with a property
'    If (TokenKind <> TOKEN_PROPERTY_NEW) And (IsAnon = 0) Then
'        Call RaiseError(OZ80_ERROR_EXPECTED, "TODO"): Exit Sub
'    End If
'
'    'Define this Property:
'    Let .PropertyCount = .PropertyCount + 1
'    ReDim Preserve .Properties(1 To .PropertyCount) As OZ80Property
'
'    If TokenKind = TOKEN_PROPERTY_NEW Then
'        'Combine the Property name with the inherited heirarchy
'        Let .Properties(.PropertyCount).Name = BaseName & TokenText
'    End If
'
'    'Move over the Property name
'    If NextTokenError() Then Exit Sub
'
'    'Offset Value:
'    '----------------------------------------------------------------------------------
'    'An optional Value can be provided before the size to set the current enumeration _
'     address; that is, you can skip to anywhere in the address space
'
'    'Is there a Value preceeding the unit size?
'    '..................................................................................
'    If (TokenAttr And ATTR_EXPR) <> 0 Then
'        'Fetch the Value given. Any error, fold upwards
'        Call ParseExpression(Value)
'        If Error Then Exit Sub
'
'        'An indefinite Value cannot be used (yet?) as the addressing of the remainder _
'         of the Type will depend upon this Value and we haven't a means of deferring _
'         enmueration of Type structures
'        If Value = INDEFINITE Then
'            'Error, and fold up
'            Call RaiseError(OZ80_ERROR_INDEFINITE, "TODO")
'            Exit Sub
'        End If
'
'        'TODO: Sanity check (must be within $0000 - $FFFF)
'        'Set the address of this property to that given
'        Let .Properties(.PropertyCount).Offset = Value * 8
'        'The next property will begin after this new address
'        Let BaseOffset = .Properties(.PropertyCount).Offset
'
'    '..................................................................................
'    Else
'        'The new Property's reserved data size will begin ater the previous
'        Let .Properties(.PropertyCount).Offset = BaseOffset
'    End If
'
'    Let LogStr.Text = "     |      |       |                                                    "
'    Call LogStr.Overwrite(0, HexStr16(.Properties(.PropertyCount).Offset \ 8))
'    Call LogStr.Overwrite(22, .Properties(.PropertyCount).Name)
'
'
'    'Record the type of unit used
'    Let .Properties(.PropertyCount).Unit = TokenKind
'
'    'Nested Types:
'    '----------------------------------------------------------------------------------
'    'If the unit for a Property is `TYPE` then we append another Type at this point, _
'     either pre-existing (by name), or an anonymous Type, embedded here-in
'    If TokenKind = TOKEN_KEYWORD_TYPE Then
'        'Skip oiver the `TYPE` keyword
'        If NextTokenError() Then Exit Sub
'
'        Call LogStr.Overwrite(7, STR_TYPE)
''        Call RaiseMessage(OZ80_LOG_INFO, LogMsg.Text)
''        Call RaiseMessage(OZ80_LOG_INFO, _
''            "-----.------.-------.----------------------------------------------------" _
''        )
'
'        'Define the nested Type's address; that is, `#type.subtype` will give you the _
'         address of the nested Type, which will be the same as the first Property _
'         there-in (e.g. `#type.subtype.property`). This is so that you need not rely _
'         on the name/position of the first Property in your code
'        Call Values.Add( _
'            .Name & .Properties(.PropertyCount).Name, _
'            .Properties(.PropertyCount).Offset \ 8 _
'        )
'
'        'Is this a pre-existing Type name?
'        '..............................................................................
'        If (TokenAttr And ATTR_NAMES) = ATTR_NAME_TYPE Then
'            'Retrieve the name of the Type
'            Dim CopyTypeName As String
'            Let CopyTypeName = TokenText
'
'            'Look for it amongst the existing Types...
'            'TODO: I would like to do away with this and have the Type name token _
'             implicitly link to its Type structure; perhaps do this at Tokenisation...
'            Dim i As Long
'            For i = 1 To TypeCount
'                Dim CopyTypeIndex As Long
'                If Types(i).Name = CopyTypeName Then Let CopyTypeIndex = i: Exit For
'            Next i
'            'TODO: Type not found?
'            If CopyTypeIndex = 0 Then Stop
'
'            'Step over the Type name. There should exist at least the end-Type _
'             bracket, so end-of-file is an error
'            If NextTokenError() Then Exit Sub
'
'            'Define the size of the sub-Type. With thanks to below for the simple _
'             solution to always-round-up division for converting bits-to-bytes _
'             <visualbasic.ittoolbox.com/groups/technical-functional/visualbasic-l/always-round-up-to-whole-number-in-vb-60-1073167#GRB4513402>
'            Call Values.Add( _
'                .Name & .Properties(.PropertyCount).Name & ".size", _
'                Abs(Int(-Types(CopyTypeIndex).Size / 8)) _
'            )
'
'            'The Property name given to the sub-Type will need to be prepended to _
'             each Property within the sub-Type as we walk them; attaching the sub-Type _
'             to the parent Type at any nesting level
'
'            Dim CopyBaseName As String
'            Let CopyBaseName = .Properties(.PropertyCount).Name
'
'            'Walk the list of Properties within the sub-Type
'            For i = 1 To Types(CopyTypeIndex).PropertyCount
'                'Create the next node
'                Let .PropertyCount = .PropertyCount + 1
'                ReDim Preserve .Properties(1 To .PropertyCount) As OZ80Property
'
'                'Start assigning the Property details
'                With Types(CopyTypeIndex).Properties(i)
'                    Let LogStr.Text = "     |      |       |                                                    "
'                    Call LogStr.Overwrite(0, HexStr16(BaseOffset \ 8))
'                    Call LogStr.Overwrite(22, BaseName & CopyBaseName & .Name)
'                    Call LogStr.Overwrite(7, TokenNames(.Unit))
'
'                    'Copy the Property across
'                    Let BaseType.Properties(BaseType.PropertyCount).Offset = BaseOffset
'                    Let BaseType.Properties(BaseType.PropertyCount).Name = BaseName & CopyBaseName & .Name
'                    Let BaseType.Properties(BaseType.PropertyCount).Unit = .Unit
'                    Let BaseType.Properties(BaseType.PropertyCount).Size = .Size
'                    Let BaseOffset = BaseOffset + .Size
'
'                    'Add the size of this Property to the parent-Type
'                    Let BaseType.Size = BaseType.Size + .Size
'
''                    Call RaiseMessage(OZ80_LOG_INFO, LogMsg.Text)
'                End With
'
'                'Add the Type Property to the global Label database
'                Call Values.Add( _
'                    .Name & .Properties(.PropertyCount).Name, _
'                    .Properties(.PropertyCount).Offset \ 8 _
'                )
'                'Define the size in the Label database (in bytes)
'                Call Values.Add( _
'                    .Name & .Properties(.PropertyCount).Name & ".size", _
'                    Abs(Int(-.Properties(.PropertyCount).Size / 8)) _
'                )
'            Next i
'
'        'An anonymous (embedded) Type:
'        '..............................................................................
'        Else
'            'Walk the Type within this Type...
'            Call ScopeType(BaseType, BaseOffset, .Properties(.PropertyCount).Name)
'            'For any error within the sub-Type, don't continue with the outer Type
'            If Error Then Exit Sub
'        End If
'
''        Call RaiseMessage(OZ80_LOG_INFO, _
''            "-----'------'-------'----------------------------------------------------" _
''        )
'
'    'Unit:
'    '----------------------------------------------------------------------------------
'    'Each Property of a Type must define a size (Bytes, Words, Longs &c.), _
'     this is used to enumerate the addresses automatically
'    Else
'        'All size measurement is done in bits
'        Dim UnitBits As Long
'        If TokenKind = TOKEN_KEYWORD_BYTE Then
'            Let UnitBits = 8
'        ElseIf TokenKind = TOKEN_KEYWORD_WORD Then Let UnitBits = 16
'        ElseIf TokenKind = TOKEN_KEYWORD_TRIP Then Let UnitBits = 24
'        ElseIf TokenKind = TOKEN_KEYWORD_LONG Then Let UnitBits = 32
'        ElseIf TokenKind = TOKEN_KEYWORD_NYBL Then Let UnitBits = 4
'        ElseIf TokenKind = TOKEN_KEYWORD_BOOL Then Let UnitBits = 1
'        Else
'            'ERROR: Not a valid Unit size!
'            Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'            Exit Sub
'        End If
'
'        Call LogStr.Overwrite(7, TokenNames(TokenKind))
'
'        'Pass over the Unit token
'        If NextTokenError() Then Exit Sub
'
'        'Multiplier?
'        '..............................................................................
'        'A Unit size (e.g. `WORD` -- 16 bits) can be multiplied as needed using the _
'         'repeat' operator "x". E.g. `WORD x 6` will give you 6 pairs of bytes
'        If TokenKind = TOKEN_OPERATOR_REP Then
'            'Pass over the repeat operator
'            If NextTokenError() Then Exit Sub
'            'Get the multiplier Value
'            Call ParseExpression(Value)
'            If Error Then Exit Sub
'
'            'This cannot be indefinite because the size of structures must be known _
'             in order to layout the ROM and therefore define the Label addresses
'            If Value = INDEFINITE Then
'                'ERROR: Size of a structure cannot be indefinite
'                Call RaiseError(OZ80_ERROR_INDEFINITE, "TODO")
'                Exit Sub
'            Else
'                'TODO: Overflow check
'                'Assign the size in bits (unit x multiplier)
'                Let .Properties(.PropertyCount).Size = UnitBits * Value
'                Call LogStr.Overwrite(14, CStr(Value))
'            End If
'
'        'No multiplier -- just use a single unit
'        '..............................................................................
'        Else
'            Let .Properties(.PropertyCount).Size = UnitBits
'            Call LogStr.Overwrite(14, "1")
'        End If
'
'        'Shift forward the enumeration address according to the size of the structure; _
'         that is, the next Property shall begin after the previous number of bits
'        Let BaseOffset = .Properties(.PropertyCount).Offset + .Properties(.PropertyCount).Size
'        'And add this to the overal size of the whole Type
'        Let .Size = .Size + .Properties(.PropertyCount).Size
'
'        'Default Value?
'        '..............................................................................
'        'TODO: Is there a default value?
'        If (TokenAttr And ATTR_LIST) <> 0 Then
'            Dim Default As Scripting.Dictionary
'            Set Default = GetListAsDictionary()
'        End If
'
'        'Don't define anonymous Properties
'        If IsAnon = 0 Then
'            'Add the Type Property to the global value database
'            Call Values.Add( _
'                .Name & .Properties(.PropertyCount).Name, _
'                .Properties(.PropertyCount).Offset \ 8 _
'            )
'            'Define the size (in bytes too) with thanks to below for the simple solution _
'             <visualbasic.ittoolbox.com/groups/technical-functional/visualbasic-l/always-round-up-to-whole-number-in-vb-60-1073167#GRB4513402>
'            Call Values.Add( _
'                .Name & .Properties(.PropertyCount).Name & ".size", _
'                Abs(Int(-.Properties(.PropertyCount).Size / 8)) _
'            )
'        End If
'
''        Call RaiseMessage(OZ80_LOG_INFO, LogMsg.Text)
'    End If
'
'    End With
'    '----------------------------------------------------------------------------------
'
'    'Has the Type come to an end? If not, loop for the next Property
'    If TokenKind = TOKEN_SQUARE_CLOSE _
'        Then Call NextToken _
'        Else GoTo Continue
'
'    'Define the size of the whole type, with thanks to below for the simple solution _
'     <visualbasic.ittoolbox.com/groups/technical-functional/visualbasic-l/always-round-up-to-whole-number-in-vb-60-1073167#GRB4513402>
'    'We check for existence because of the ability to `APPEND` to a Type
'    'TODO: Do away with the use of a Dictionary object
'    If Values.Exists(BaseType.Name & BaseName & ".size") Then
'        Let Values.Item(BaseType.Name & BaseName & ".size") = Abs(Int(-BaseType.Size / 8))
'    Else
'        Call Values.Add(BaseType.Name & BaseName & ".size", Abs(Int(-BaseType.Size / 8)))
'    End If
'End Sub
'
''ProcessInclude : Process an Include statement
''======================================================================================
''INCLUDE    <text>
''======================================================================================
'Private Sub ProcessInclude()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
'    'Fetch the file path _
'     (if there's an error doing this, fold upwards)
'    Dim FilePath As bluString
'    If GetText(FilePath) Then Exit Sub
'
'    'Remember the current file being parsed
'    Dim ReturnFile As Long, ReturnToken As Long
'    Let ReturnFile = File
'    Let ReturnToken = TokenIndex
'
'    'Parse & assemble the included file
'    Call FileAssemble(FilePath.Text)
'    If Error Then Exit Sub
'
'    'Reattach the previous file
'    Let File = ReturnFile
'    Let TokenIndex = ReturnToken
'    Let TokensLength = Files(File).Stream.Length
'
'    With Files(File).Stream.Tokens(TokenIndex)
'        Let TokenKind = .Kind
'        Let TokenValue = .Value
'        Let TokenLine = .Line
'        Let TokenCol = .Col
'    End With
'
''    Call RaiseMessage( _
''        OZ80_LOG_ACTION, "Continuing Assembly of ""{$}""", _
''        Files(File).Path _
''    )
'End Sub
'
''DefineProc : Process a Procedure Chunk
''======================================================================================
''PROC   :<label>
''       [HELP       <text>]
''       [SECTION    ::<section>]    |   [INTERRUPT  <expr>]
''       [PARAMS     <params>]
''       [RETURN     <params>]
''{ [...] }
''======================================================================================
'Private Sub DefineProc()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then GoTo Finish
'
'    'ERROR: No Label name for Procedure
'    If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_LABEL Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, _
'        "Expected Label" _
'    ): GoTo Finish
'
'    'Remember the String of the Label name, _
'     we can't create the Label until the parameters are read
'    Dim LabelName As String
'    Let LabelName = TokenText
'
'    'ERROR: The file cannot end until the Chunk has been opened and closed
'    If NextTokenError() Then GoTo Finish
'
'    'The `INTERRUPT` parameter is optional, but can be zero, so use 'undefined'
'    Dim Interrupt As Double
'    Let Interrupt = INDEFINITE
'
'ReadParams:
'    'Keep track of which parameters have been included or not
'    Dim IsInterruptDefined As Boolean
'    Dim IsSectionDefined As Boolean
'    Dim IsParamsDefined As Boolean
'    Dim IsReturnDefined As Boolean
'    Dim IsHelpDefined As Boolean
'
'    '----------------------------------------------------------------------------------
'    '`SECTION` Parameter : Set which Section the Procedure belongs to _
'     ----------------------------------------------------------------------------------
'    If TokenKind = TOKEN_KEYWORD_SECTION Then
'        'ERROR: The Section parameter has already been specified
'        If IsSectionDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'ERROR: Cannot co-exist with the `INTERRUPT` parameter
'        If IsInterruptDefined Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        'ERROR: File ended before the Section name?
'        If NextTokenError() Then GoTo Finish
'
'        'ERROR: Not a Section name following the Section keyword
'        If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
'            OZ80_ERROR_EXPECTED, "TODO" _
'        ): GoTo Finish
'
'        Dim SectionName As String
'        Let SectionName = TokenText
'
'        Dim i As Long
'        'Look through the currently defined Sections for the Section name mentioned
'        For i = 1 To SectionCount
'            Dim SectionIndex As Long
'            If Sections(i).Name = SectionName Then SectionIndex = i: Exit For
'        Next i
'
'        'ERROR: Section has not been defined
'        If SectionIndex = 0 Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        Let IsSectionDefined = True
'
'        'Pass over the Section name, if the file ends, the Procedure isn't complete!
'        If NextTokenError() Then GoTo Finish
'
'        GoTo ReadParams
'
'    '----------------------------------------------------------------------------------
'    '`HELP` Parameter : Provide a description for the procedure
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_HELP Then
'        'ERROR: The `HELP` parameter has already been specified
'        If IsParamsDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'Pass over the `HELP` keyword, the file cannot end here
'        If NextTokenError() Then GoTo Finish
'
'        'Fetch the text description (ignored for now)
'        'For now, just read and discard only text tokens
'        Do
'            'If not a documentation text token, then continue
'            If TokenKind <> TOKEN_DOC Then Exit Do
'            'Move over the token, the file ending is an error
'            If NextTokenError() Then GoTo Finish
'        Loop
'
'        Let IsHelpDefined = True
'        GoTo ReadParams
'
'    '----------------------------------------------------------------------------------
'    '`PARAMS` Parameter : Document which Z80 parameters the Procedure uses _
'     ----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_PARAMS Then
'        'The `PARAMS` parameter is not used by the assembler at this time, _
'         neither is it a standard list as it consists of Registers and/or RAM Names, _
'         and optional `HELP` descriptions for each
'
'        'ERROR: The `PARAMS` parameter has already been specified
'        If IsParamsDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'Pass over the `PARAMS` keyword, the file cannot end here
'        If NextTokenError() Then GoTo Finish
'
'        Do
'            'Check for a valid `PARAMS` item
'            If (TokenAttr And ATTR_PARAM) <> 0 Then
'                'Skip over the list item, these are not recorded just yet _
'                 but will be in the future (for auto-documentation)
'                'ERROR: File cannot end in the midst of the List
'                If NextTokenError() Then GoTo Finish
'
'                'Check for documentation text for the list item
'                If TokenKind = TOKEN_DOC Then
'                    'Fetch the text description (ignored for now)
'                    'For now, just read and discard only text tokens
'                    Do
'                        'If not documentation text, then continue
'                        If TokenKind <> TOKEN_DOC Then Exit Do
'                        'Move over the token, the file ending is an error
'                        If NextTokenError() Then GoTo Finish
'                    Loop
'                End If
'
'            'The list ends when we hit one of the other `PROC` parameters or the _
'             opening brace. Anything else is considered a bad `PARAMS` list item
'            ElseIf (TokenKind = TOKEN_KEYWORD_HELP) _
'                Or (TokenKind = TOKEN_KEYWORD_RETURN) _
'                Or (TokenKind = TOKEN_KEYWORD_SECTION) _
'                Or (TokenKind = TOKEN_BRACES_OPEN) _
'            Then
'                'The list has come to the end, _
'                 look for more parameters
'                Let IsParamsDefined = True
'                GoTo ReadParams
'            Else
'                'ERROR: Only Registers and RAM Names are allowed at the moment
'                Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'                GoTo Finish
'            End If
'        Loop
'
'    '----------------------------------------------------------------------------------
'    '`RETURN` Parameter : Document what the Procedure sets / returns
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_RETURN Then
'        'The `RETURN` parameter is not used by the assembler at this time, _
'         neither is it a standard list as it consists of Registers and/or RAM Names
'
'        'ERROR: The `RETURN` parameter has already been specified
'        If IsReturnDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'Pass over the `RETURN` keyword, the file cannot end here
'        If NextTokenError() Then GoTo Finish
'
'        Do
'            'Check for a valid `RETURN` item
'            If (TokenAttr And ATTR_PARAM) <> 0 Then
'                'Skip over the list item, these are not recorded just yet _
'                 but will be in the future (for auto-documentation)
'                'ERROR: File cannot end in the midst of the List
'                If NextTokenError() Then GoTo Finish
'
'                'Check for documentation text for the list item
'                If TokenKind = TOKEN_DOC Then
'                    'Fetch the text description (ignored for now)
'                    'For now, just read and discard only text tokens
'                    Do
'                        'If not documentation text, then continue
'                        If TokenKind <> TOKEN_DOC Then Exit Do
'                        'Move over the token, the file ending is an error
'                        If NextTokenError() Then GoTo Finish
'                    Loop
'                End If
'
'            'The list ends when we hit one of the other `PROC` parameters or the _
'             opening brace. Anything else is considered a bad `RETURN` list item
'            ElseIf (TokenKind = TOKEN_KEYWORD_HELP) _
'                Or (TokenKind = TOKEN_KEYWORD_PARAMS) _
'                Or (TokenKind = TOKEN_KEYWORD_SECTION) _
'                Or (TokenKind = TOKEN_BRACES_OPEN) _
'            Then
'                'The list has come to the end, _
'                 look for more parameters
'                Let IsReturnDefined = True
'                GoTo ReadParams
'            Else
'                'ERROR: Only Registers and RAM Names are allowed at the moment
'                Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'                GoTo Finish
'            End If
'        Loop
'
'    '----------------------------------------------------------------------------------
'    '`INTERRUPT` Parameter : Define a fixed-location Interrupt Procedure
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_INTERRUPT Then
'        'The `INTERRUPT` parameter fixes a Procedure to a specific location in the _
'         ROM intended for hardware interrupts. This parameter cannot co-exist with _
'         the `SECTION` parameter
'
'        'ERROR: The `INTERRUPT` parameter has already been specified
'        If IsInterruptDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'ERROR: Cannot co-exist with the `SECTION` parameter
'        If IsSectionDefined Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        'ERROR: File cannot end before the parameter Expression!
'        If NextTokenError() Then GoTo Finish
'
'        'Is this an Expression at all?
'        If (TokenAttr And ATTR_EXPR) = 0 Then Call RaiseError( _
'            OZ80_ERROR_EXPECTED, "TODO" _
'        ): GoTo Finish
'
'        'Calculate the Expression:
'        '(If there was an error in the Expression, the error event _
'          will already have been raised, we only need to fold upwards)
'        Call ParseExpression(Interrupt)
'        If Error Then GoTo Finish
'
'        'An indefinite Value cannot be used for an Interrupt address, we will only be _
'         able to define the indefinite Values (i.e. Label locations in ROM) once _
'         the whole code is parsed and we use the Sections to layout the ROM
'        If Interrupt = INDEFINITE Then Call RaiseError( _
'            OZ80_ERROR_INDEFINITE, "TODO" _
'        ): GoTo Finish
'
'        'Validate the Value, only certain Interrupt addresses are allowed
'        If Not ((Interrupt = &H0) Or (Interrupt = &H8) Or (Interrupt = &H10) _
'            Or (Interrupt = &H18) Or (Interrupt = &H20) Or (Interrupt = &H28) _
'            Or (Interrupt = &H30) Or (Interrupt = &H38) Or (Interrupt = &H66)) _
'        Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        Let IsInterruptDefined = True
'        GoTo ReadParams
'    End If
'
'    '----------------------------------------------------------------------------------
'
'    'With the parameters read, the only allowed token is the opening of the Chunk
'    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, "TODO" _
'    ): GoTo Finish
'
'    'The chunk cannot be left open
'    If NextTokenError() Then GoTo Finish
'
'    'Now define the Label name; if a Section was given, it needs to be prepended
'    Dim FullName As String
'    Let FullName = SectionName & LabelName
'
'    'ERROR: Duplicate Chunk name
'    If ChunkExists(FullName) Then Call RaiseError( _
'        OZ80_ERROR_DUPLICATE, "TODO" _
'    ): GoTo Finish
'
'    'This will be the index of the new Chunk we create
'    Dim Chunk As Long
'
'    'Create the Chunk
'    If Interrupt = INDEFINITE Then
'        'A standard Procedure without fixed (interrupt) address
'        Let Chunk = Chunk_Begin(LabelName, SectionIndex)
'    Else
'        'A special fixed-Procedure for hardware interrupts
'        Let Chunk = Chunk_Begin(LabelName, -1, Interrupt)
'        'We keep direct references to Interrupt Procedures to make it easier to _
'         lay out the ROM contents
'        If Interrupt = &H0& Then
'            Let Chunk_Interrupt00 = Chunk
'        ElseIf Interrupt = &H8& Then Let Chunk_Interrupt08 = Chunk
'        ElseIf Interrupt = &H10& Then Let Chunk_Interrupt10 = Chunk
'        ElseIf Interrupt = &H18& Then Let Chunk_Interrupt18 = Chunk
'        ElseIf Interrupt = &H20& Then Let Chunk_Interrupt20 = Chunk
'        ElseIf Interrupt = &H28& Then Let Chunk_Interrupt28 = Chunk
'        ElseIf Interrupt = &H30& Then Let Chunk_Interrupt30 = Chunk
'        ElseIf Interrupt = &H38& Then Let Chunk_Interrupt38 = Chunk
'        ElseIf Interrupt = &H66& Then Let Chunk_Interrupt66 = Chunk
'        End If
'    End If
'
'    'If this Chunk belongs to a Section, add its size to the total count
'    If IsSectionDefined Then
'        Let Sections(SectionIndex).Size = _
'            Sections(SectionIndex).Size + Chunks(Chunk).Size
'    End If
'
'    'Count this toward the total size of all Chunks, which we will use to determine _
'     the size of output ROM needed
'    Let ChunksSize = ChunksSize + Chunks(Chunk).Size
'
'    'Log a header for the disassembly
''    Call Chunks(Chunk).LogASM.Add("PROC " & LabelName)
''    Call Chunks(Chunk).LogASM.Add(LOG_LINEDOUBLE)
'    Let ChunkMode = MODE_UNSET
'
'    'Enter the Procedure Chunk to process its innards
'    Call ScopeChunk(Chunk)
'
'    'Finalise the disassembly log
'    If ChunkMode = MODE_DATA Then
'        'If data bytes have been assembled, ensure the remainder are added to _
'         the log since we only output the disassembly line every 16 data bytes
''        If ChunkDataIndex <> 0 Then Call Chunks(Chunk).LogASM.Add(ChunkDataLog)
'    End If
''    Call Chunks(Chunk).LogASM.Add(LOG_LINEDOUBLE)
''    Call Chunks(Chunk).LogASM.Add("Size: " & Format$(Chunks(Chunk).Size, "#,0") & " bytes")
''    Call RaiseMessage(OZ80_LOG_DEBUG, Chunks(Chunk).LogASM.Concatenate())
'
'Finish:
'End Sub
'
''DefineSection : Define a Section
''======================================================================================
''SECTION    ::<section>    [SLOT   <list>]
''======================================================================================
'Private Sub DefineSection()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
'    'ERROR: Not a section name following the section keyword
'    If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, _
'        "Expected Section Name" _
'    ): Exit Sub
'
'    'Keep ahold of the Section identifier
'    Dim SectionName As String: Let SectionName = TokenText
'
'    'Has this Section already been defined?
'    Dim s As Long
'    For s = 1 To SectionCount
'        'ERROR: Section has already been defined!
'        If Sections(s).Name = SectionName Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): Exit Sub
'    Next s
'
'    '----------------------------------------------------------------------------------
'
'    'If the file ends here it's valid, as the params are optional
'    If EOF Then GoTo Finish
'    Call NextToken
'
'    'The only allowed Slot lists are: 1 KB | 0 | 1 | 2 | 0, 1 | 0, 1, 2 | 1, 2. _
'     These are stored as a bit-pattern allowing packing into one value, _
'     which when zero infers the default "let the assembler choose for me"
'    Dim Slot As OZ80_SLOT
'
'    'The only parameter is the Slot List
'    If TokenKind = TOKEN_KEYWORD_SLOT Then
'        'ERROR: File ending before we have the value
'        If NextTokenError() Then Exit Sub
'
'        'ERROR: There needs to be at least one item in the List
'        If (TokenAttr And ATTR_EXPR) = 0 Then Call RaiseError( _
'            OZ80_ERROR_EXPECTED, "TODO" _
'        ): Exit Sub
'
'        'Get the list of Slots assigned to the Section
'        Dim List As Scripting.Dictionary
'        Set List = GetListAsDictionary()
'
'        'Can't be more than three -- "0, 1, 2"
'        If List.Count > 3 Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): Exit Sub
'
'        'An indefinite value cannot be used for the Slot
'        If List.Item(0) = INDEFINITE Then GoTo ErrIndefinite
'
'        'Check for one item: 1KB, 0, 1 or 2
'        If List.Count = 1 Then
'            If List.Item(0) = 1024 Then Let Slot = SLOT1KB: GoTo Finish
'            If List.Item(0) = 0 Then Let Slot = SLOT0: GoTo Finish
'            If List.Item(0) = 1 Then Let Slot = SLOT1: GoTo Finish
'            If List.Item(0) = 2 Then Let Slot = SLOT2: GoTo Finish
'            'An invalid value
'            GoTo ErrInvalid
'        End If
'
'        'Check for two items: 0, 1 or 1, 2
'        If List.Count = 2 Then
'            'An indefinite value cannot be used for the Slot
'            If List.Item(1) = INDEFINITE Then GoTo ErrIndefinite
'            If List.Item(0) = 0 And List.Item(1) = 1 Then Let Slot = SLOT0 Or SLOT1: GoTo Finish
'            If List.Item(0) = 1 And List.Item(1) = 2 Then Let Slot = SLOT1 Or SLOT2: GoTo Finish
'            'An invalid value
'            GoTo ErrInvalid
'        End If
'
'        'Check for 0, 1, 2
'        If List.Item(0) <> 0 Then GoTo ErrInvalid
'        If List.Item(1) <> 1 Then GoTo ErrInvalid
'        If List.Item(2) <> 2 Then GoTo ErrInvalid
'        Let Slot = SLOT0 Or SLOT1 Or SLOT2
'    End If
'
'Finish:
'    '----------------------------------------------------------------------------------
'    'Record the result:
''    Call RaiseMessage(OZ80_LOG_STATUS, LOG_SECTION, SectionName)
'
'    Let SectionCount = SectionCount + 1
'    ReDim Preserve Sections(0 To SectionCount) As oz80Section
'
'    Let Sections(SectionCount).Name = SectionName
'    Let Sections(SectionCount).ID = SectionCount
'    Let Sections(SectionCount).Slot = Slot
'
'    Call Values.Add(SectionName, INDEFINITE)
'
'    Exit Sub
'
'    '----------------------------------------------------------------------------------
'ErrInvalid:
'    'Only a specific Slot pattern can be used
'    Call RaiseError(OZ80_ERROR_INVALID, "TODO")
'    Exit Sub
'
'ErrIndefinite:
'    'An indefinite Value cannot be used for a Slot number, we will only be _
'     able to define the indefinite Values (i.e. Label locations in ROM) once _
'     the whole code is parsed and we use the Sections to layout the ROM
'    Call RaiseError(OZ80_ERROR_INDEFINITE, "TODO")
'End Sub
'
''DefineRAM : Process a RAM block
''======================================================================================
''RAM       [<expr>]     [ ... ]
''======================================================================================
'Private Sub DefineRAM()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
''    Call RaiseMessage(OZ80_LOG_STATUS, "TYPE {$}", TypeName)
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
''    Call RaiseMessage(OZ80_LOG_INFO, _
''                      "ADDR | UNIT | QTY   | NAME")
''    Call RaiseMessage(OZ80_LOG_INFO, _
''                      "-----¦------¦-------¦----------------------------------------------------")
'
'    Dim TypeName As String
'    'Skip the param
'    If (TokenAttr And ATTR_NAMES) = ATTR_NAME_RAM Then
'        Let TypeName = TokenText
'        If NextTokenError() Then Exit Sub
'    Else
'        Let TypeName = "$"
'    End If
'
'    'Create the new Type object
'    Let TypeCount = TypeCount + 1
'    ReDim Preserve Types(0 To TypeCount) As OZ80Type
'    Let Types(TypeCount).Name = TypeName
'
'    'TODO: Use the right address
'    Call ScopeType(Types(TypeCount))
'
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE & vbNewLine)
'End Sub
'
''DefineRAMPage : Define a RAM Page
''======================================================================================
'Private Sub DefineRAMPage()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
'    'The RAM Page name follows
'    If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_RAM Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, _
'        "Expected RAM Page name" _
'    ): Exit Sub
'
'    If NextTokenError() Then Exit Sub
'
'    'Optional Expression may follow
'    If (TokenAttr And ATTR_EXPR) <> 0 Then
'        Dim ReturnValue As Double
'        Call ParseExpression(ReturnValue)
'
'        If (TokenAttr And ATTR_EXPR) <> 0 Then
'            Call ParseExpression(ReturnValue)
'        End If
'    End If
'End Sub
'
''DefineStart : Define the System's boot Procedure
''======================================================================================
''START      { ... }
''======================================================================================
'Private Sub DefineStart()
'    'You can't define the START procedure twice
'    If ChunkExists(STR_START) Then Call RaiseError( _
'        OZ80_ERROR_DUPLICATE, "TODO" _
'    ): Exit Sub
'
'    'ERROR: The file cannot end until the Chunk has been opened and closed
'    If NextTokenError() Then Exit Sub
'
'    'The only allowed token is the opening of the Chunk
'    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, "TODO" _
'    ): Exit Sub
'
'    'The chunk cannot be left open
'    If NextTokenError() Then Exit Sub
'
'    'Create the Chunk, and keep a reference to this for laying out the ROM
'    'TODO: Check that the System has been defined
'    Let Chunk_Start = Chunk_Begin(STR_START, -1, 0)
'
'    'Enter the Procedure Chunk to process its innards
'    Call ScopeChunk(Chunk_Start)
'End Sub
'
''DefineSystem : Define the target system (e.g. SEGA Master System)
''======================================================================================
''SYSTEM     <text>
''======================================================================================
'Private Sub DefineSystem()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then Exit Sub
'
'    'Check that the parameter is a list (or at least the start of one)
'    If (TokenAttr And ATTR_LIST) = 0 Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, "TODO" _
'    ): Exit Sub
'
'    'Fetch the System name string from the proceeding list _
'     (if there was any error during that, fold upwards)
'    Dim SystemName As bluString
'    If GetText(SystemName) Then Exit Sub
'
'    'Check if it's any of the allowed System names
'    If SystemName.Equals(STR_SMS) Then
'        'Configure the assembler for the SEGA Master System
'        Let System = SYSTEM_SMS
'    Else
'        'Not a valid system name
'        Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'    End If
'End Sub
'
''DefineTable : Process a data table Chunk
''======================================================================================
''TABLE  :<label>
''       [HELP       <text>]
''       [SECTION    ::<section>]
''{ [...] }
''======================================================================================
'Private Sub DefineTable()
'    'ERROR: The file ended before the first parameter
'    If NextTokenError() Then GoTo Finish
'
'    'ERROR: No Label name for Table
'    If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_LABEL Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, "TODO" _
'    ): GoTo Finish
'
'    'Remember the String of the Label name, _
'     we can't create the Label until the parameters are read
'    Dim LabelName As String
'    Let LabelName = TokenText
'
'    'ERROR: The file cannot end until the Chunk has been opened and closed
'    If NextTokenError() Then GoTo Finish
'
'ReadParams:
'    'Keep track of which params have been included or not
'    Dim IsSectionDefined As Boolean
'
'    '----------------------------------------------------------------------------------
'    '`SECTION` Parameter : Set which Section the Table belongs to _
'     ----------------------------------------------------------------------------------
'    If TokenKind = TOKEN_KEYWORD_SECTION Then
'        'ERROR: The Section parameter has already been specified
'        If IsSectionDefined Then Call RaiseError( _
'            OZ80_ERROR_DUPLICATE, "TODO" _
'        ): GoTo Finish
'
'        'ERROR: File ended before the Section name?
'        If NextTokenError() Then GoTo Finish
'
'        'ERROR: Not a Section name following the Section keyword
'        If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_SECTION Then Call RaiseError( _
'            OZ80_ERROR_EXPECTED, "TODO" _
'        ): GoTo Finish
'
'        Dim i As Long
'        'Look through the currently defined Sections for the Section name mentioned
'        For i = 1 To SectionCount
'            Dim SectionIndex As Long
'            If Sections(i).Name = TokenText Then SectionIndex = i: Exit For
'        Next i
'
'        'ERROR: Section has not been defined
'        If SectionIndex = 0 Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        Dim SectionName As String
'        Let SectionName = TokenText
'
'        Let IsSectionDefined = True
'
'        'Pass over the Section name, if the file ends, the Procedure isn't complete!
'        If NextTokenError() Then GoTo Finish
'
'        GoTo ReadParams
'    End If
'
'    'With the parameters read, the only allowed token is the opening of the Chunk
'    If TokenKind <> TOKEN_BRACES_OPEN Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED, "TODO" _
'    ): GoTo Finish
'
'    'The Chunk cannot be left open
'    If NextTokenError() Then GoTo Finish
'
'    '----------------------------------------------------------------------------------
'
'    'Now define the Label name; if a Section was given, it needs to be prepended
'    Dim FullLabelName As String
'    Let FullLabelName = SectionName & LabelName
'
'    'You can't name two Chunks the same thing
'    If ChunkExists(FullLabelName) Then Call RaiseError( _
'        OZ80_ERROR_DUPLICATE, "TODO" _
'    ): GoTo Finish
'
'    'Create the Chunk
'    Dim Chunk As Long
'    Let Chunk = Chunk_Begin(LabelName, SectionIndex)
'
'    'Log a header for the disassembly
''    Call Chunks(Chunk).LogASM.Add("TABLE " & LabelName)
''    Call Chunks(Chunk).LogASM.Add(LOG_LINEDOUBLE)
'    Let ChunkMode = MODE_UNSET
'
'    'Enter the Table Chunk to process its innards
'    Call ScopeChunk(Chunk)
'
'    'Finalise the disassembly log
'    If ChunkMode = MODE_DATA Then
''        If ChunkDataIndex <> 0 Then Call Chunks(Chunk).LogASM.Add(ChunkDataLog)
'    End If
''    Call Chunks(Chunk).LogASM.Add(LOG_LINEDOUBLE)
''    Call Chunks(Chunk).LogASM.Add("Size: " & Format$(Chunks(Chunk).Size, "#,0") & " bytes")
''    Call RaiseMessage(OZ80_LOG_DEBUG, Chunks(Chunk).LogASM.Concatenate())
'
'    'Count this toward the total size of all Chunks, _
'     which we will use to determine the size of output ROM needed
'    Let ChunksSize = ChunksSize + Chunks(Chunk).Size
'
'Finish:
'End Sub
'
''ScopeChunk: Within a code/data Chunk, i.e. `PROC`/`TABLE`
''======================================================================================
'Private Sub ScopeChunk( _
'    ByRef ChunkID As Long _
')
'    Dim ExprToken As Long
'    Dim Value As Double
'
'Continue:
'    '----------------------------------------------------------------------------------
'    'Z80 code?
'    '----------------------------------------------------------------------------------
'    If (TokenAttr And ATTR_Z80_MNEMONIC) <> 0 Then
'        'Assemble the Z80 instruction...
'        Call Z80Assemble(ChunkID): If Error Then Exit Sub
'
'        'Processing of the Z80 instruction (and parameters) will have moved the token _
'         pointer on to the next unread token already, so we can just loop right back
'        GoTo Continue
'
'    '----------------------------------------------------------------------------------
'    'Local Label definition?
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_PROPERTY_NEW Then
'        'Define the local property
'        'TODO: Duplicates?
'        Call Chunks(ChunkID).Properties.Add(TokenText, CDbl(Chunks(ChunkID).Size))
'        'Fully qualify the local Label name to its global identifier
'        Call AddUpdateLabel(Chunks(ChunkID).FullName & TokenText, INDEFINITE)
'
'    '----------------------------------------------------------------------------------
'    'INDEX
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_INDEX Then
'        'Move past the `INDEX` keyword; _
'         error if the file ends because the scope is obviously still open
'        If NextTokenError() Then Exit Sub
'
'        'The next token should be a Local Label name
'        If (TokenAttr And ATTR_NAMES) <> ATTR_NAME_PROPERTY Then
'            Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'            Exit Sub
'        End If
'
'        'Define the local property
'        Call Chunks(ChunkID).Properties.Add(TokenText, CDbl(Chunks(ChunkID).Size))
'
'        'Fully qualify the local Label name to its global identifier
'        'Add this to the label database, using the current offset position in the Chunk
'        Call AddUpdateLabel( _
'            Chunks(ChunkID).FullName & TokenText, INDEFINITE _
'        )
'
'        'Define the Index
'        Call AddUpdateLabel( _
'            Chunks(ChunkID).FullName & TokenText & STR_INDEX_LABEL, _
'            CDbl(Chunks(ChunkID).TableIndex) _
'        )
'
'        'Move to the next index
'        'TODO: Too many indicies at 256?
'        Let Chunks(ChunkID).TableIndex = Chunks(ChunkID).TableIndex + 1
'
'    '----------------------------------------------------------------------------------
'    'BYTE Data
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_BYTE Then
'        'Move past the `BYTE` keyword; _
'         error if the file ends because the scope is obviously still open
'        If NextTokenError() Then Exit Sub
'
'        'ERROR: There MUST be at least one List item after the keyword
'        If (TokenAttr And ATTR_LIST) = 0 Then
'            Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'            Exit Sub
'        End If
'
'        'Start walking the List
'        Do
'            If (TokenAttr And ATTR_EXPR) <> 0 Then
'                Let ExprToken = TokenIndex
'
'                Call ParseExpression(Value, ChunkID)
'                If Error Then Stop
'
'                If Value = INDEFINITE Then
'                    Call DeferExpression( _
'                        ASM_VAL8, File, ChunkID, ExprToken, Chunks(ChunkID).Size, 0 _
'                    )
'                    Call Chunk_AddBYTE(ChunkID, 0)
'                Else
'                    Call Chunk_AddBYTE(ChunkID, Value)
'                End If
'
'            ElseIf TokenKind = TOKEN_TEXT Then
'                'Walk the Text bytes
'                Let LogStr.Text = TokenText
'                Dim i As Long
'                For i = 0 To LogStr.Length - 1
'                    Call Chunk_AddBYTE(ChunkID, LogStr.Char(i) And &HFF&)
'                Next i
'                If NextTokenError Then Exit Sub
'
'            ElseIf (TokenAttr And ATTR_LIST) = 0 Then
'                Exit Do
'
'            Else
'                Stop
'                Call NextToken
'            End If
'        Loop
'        GoTo Continue
'
'    '----------------------------------------------------------------------------------
'    'WORD Data
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_WORD Then
'        'Move past the `WORD` keyword; _
'         error if the file ends because the scope is obviously still open
'        If NextTokenError() Then Exit Sub
'
'        'ERROR: There MUST be at least one List item after the keyword
'        If (TokenAttr And ATTR_LIST) = 0 Then
'            Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'            Exit Sub
'        End If
'
'        'Start walking the List
'        Do
'            If (TokenAttr And ATTR_EXPR) <> 0 Then
'                Let ExprToken = TokenIndex
'
'                Call ParseExpression(Value, ChunkID)
'                If Error Then Stop
'
'                If Value = INDEFINITE Then
'                    Call DeferExpression( _
'                        ASM_VAL16, File, ChunkID, ExprToken, Chunks(ChunkID).Size, 0 _
'                    )
'                    Call Chunk_AddWORD(ChunkID, 0)
'                Else
'                    Call Chunk_AddWORD(ChunkID, Value)
'                End If
'
'            ElseIf (TokenAttr And ATTR_LIST) = 0 Then
'                Exit Do
'            Else
'                'TODO: Text...
'                Call NextToken
'            End If
'        Loop
'        GoTo Continue
'
'    '----------------------------------------------------------------------------------
'    'The chunk might be empty `{}`, in which case we don't want to walk it
'    ElseIf TokenKind = TOKEN_BRACES_CLOSE Then
'        GoTo Finish
'    Else
'        'ERROR: Not a valid token for this scope
'        Call RaiseError(OZ80_ERROR_EXPECTED, "TODO")
'        Exit Sub
'    End If
'
'    '----------------------------------------------------------------------------------
'
'    'Move to the next token. If the file ends unexpectedly, the chunk has not been _
'     closed properly (closing brace is expected)
'    If NextTokenError() Then Exit Sub
'
'    'When the closing brace is encountered, we'll drop out of the scope
'    If TokenKind <> TOKEN_BRACES_CLOSE Then GoTo Continue
'
'Finish:
'    'Move onto the next token ready for the root scope to process. _
'     The file could end just after the chunk, so no error is necessary
'    Call NextToken
'End Sub
'
''Chunk_AddBYTE
''======================================================================================
''Value          | The "BYTE" to add to the current Chunk
''======================================================================================
'Private Sub Chunk_AddBYTE( _
'    ByRef ChunkID As Long, _
'    ByRef Value As Double _
')
'    With Chunks(ChunkID)
'        'Increase the Chunk accordingly
'        Let .Size = .Size + 1
'
'        'If the Chunk's buffer is full then double the buffer size to compensate. _
'        This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, ...
'        If .Size >= .Buffer Then
'            Let .Buffer = .Buffer * 2
'            ReDim Preserve .Data(1 To .Buffer) As Byte
'        End If
'
'        'Store the byte
'        Let .Data(.Size) = Value And &HFF&
'    End With
'
'    'Add this to the disassembly log
'    Call Chunk_LogByte(ChunkID, Value And &HFF&)
'End Sub
'
''Chunk_AddWORD
''======================================================================================
''Value          | The "WORD" (2-bytes) to add to the current Chunk
''======================================================================================
'Private Sub Chunk_AddWORD( _
'    ByRef ChunkID As Long, _
'    ByRef Value As Double _
')
'    With Chunks(ChunkID)
'        'Increase the Chunk accordingly
'        Let .Size = .Size + 2
'
'        'If the Chunk's buffer is full then double the buffer size to compensate. _
'        This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, ...
'        If .Size >= .Buffer Then
'            Let .Buffer = .Buffer * 2
'            ReDim Preserve .Data(1 To .Buffer) As Byte
'        End If
'
'        'Store the bytes, low then high (Little Endian)
'        Let .Data(.Size - 1) = Value And &HFF&
'        Let .Data(.Size) = (Value \ &H100&) And &HFF&
'    End With
'
'    'Add this to the disassembly log
'    Call Chunk_LogByte(ChunkID, Value And &HFF&)
'    Call Chunk_LogByte(ChunkID, (Value \ &H100&) And &HFF&)
'End Sub
'
''Chunk_LogByte
''======================================================================================
''Value          | The byte to add to the current disassembly line
''======================================================================================
'Private Sub Chunk_LogByte( _
'    ByRef ChunkID As Long, _
'    ByRef Value As Byte _
')
'    'If we've just come from Z80 instructions, we'll need to switch to the data _
'     columns in the disassembly log
'    If ChunkMode <> MODE_DATA Then
'        'Start a new disassembly log line
'        Let ChunkDataLog = LOG_DATA
'        'Include the current address on the log line
'        Call W32TLB.MemCopy( _
'            OutputPointer:=StrPtr(ChunkDataLog), _
'            SourcePointer:=StrPtr(HexStr16(Chunks(ChunkID).Size - 1)), _
'              LengthBytes:=8 _
'        )
'        'Bytes will be logged according to the current offset _
'         (so that the data is displayed aligned to 16 bytes)
'        Let ChunkDataIndex = (Chunks(ChunkID).Size - 1) Mod 16
'
'        If ChunkMode <> MODE_UNSET Then
''            Call Chunks(ChunkID).LogASM.Add(LOG_ASM_FOOT)
'        End If
''        Call Chunks(ChunkID).LogASM.Add(LOG_DATA_COLS)
''        Call Chunks(ChunkID).LogASM.Add(LOG_DATA_HEAD)
'        Let ChunkMode = MODE_DATA
'    End If
'
'    'Place the byte into the disassembly line
'    Call W32TLB.MemCopy( _
'        OutputPointer:=StrPtr(ChunkDataLog) + 14 + (ChunkDataIndex * 6), _
'        SourcePointer:=StrPtr(HexStr8(Value)), _
'          LengthBytes:=4 _
'    )
'
'    'TODO: Adding the ASCII disply of bytes will require a Unicode textbox
'
'    'If we've reached the end of a disassembly line then log it and start a new one
'    If ChunkDataIndex = 15 Then
'        'Add the prepared line to the log
''        Call Chunks(ChunkID).LogASM.Add(ChunkDataLog)
'        'Start the next line
'        Let ChunkDataLog = LOG_DATA
'        Let ChunkDataIndex = 0
'        'Include the next address on the log line
'        Call W32TLB.MemCopy( _
'            OutputPointer:=StrPtr(ChunkDataLog), _
'            SourcePointer:=StrPtr(HexStr16(Chunks(ChunkID).Size)), _
'              LengthBytes:=8 _
'        )
'    Else
'        'Move on to the next byte in the line
'        Let ChunkDataIndex = ChunkDataIndex + 1
'    End If
'End Sub

''Z80Assemble : Assemble bytes from a Z80 instruction and parameters
''======================================================================================
'Private Sub Z80Assemble( _
'    ByRef ChunkID As Long _
')
'    '*** TIMER BEGIN ***
'    Call Profiler.EventBegin(EVENT_Z80ASSEMBLE)
'
'    'Remember which Z80 instruction is being processed
'    Dim Instruction As oz80Instruction
'    Dim Mnemonic As OZ80_TOKEN
'    Let Mnemonic = TokenKind
'
'    'Skip over the Z80 instruction token now that we know it. _
'     If the file ends it's an error because the Chunk has not been closed
'    If NextTokenError() Then GoTo Finish
'
'    'Set up the disassembly log template
'    Dim LogStr As String:   Let LogStr = LOG_ASM
'    Dim LogPtr As Long:     Let LogPtr = StrPtr(LogStr)
'
'    'We already know the instruction, so we can add that to the disassembly now
'    Call W32TLB.MemCopy(LogPtr + 54, StrPtr( _
'        ParamText(PARAMTEXT_Z80 + (Mnemonic - [_TOKEN_Z80_BEGIN])) _
'    ), 20)
'
'    'Initialse variables to hold pieces of the assembly
'    Dim Prefix As Long, Offset As Double
'    Let Prefix = 0: Let Offset = INDEFINITE
'
'    'Z80 instructions can be up to 5 bytes long depending on parameters
'    Dim ByteCount As Long
'    Dim Bytes(0 To 4) As Byte
'
'    'A built in macro allows `PUSH` and `POP` Z80 instructions to handle multiple _
'     parameters; each parameter is broken into its own `PUSH`/`POP` instruction
'    '----------------------------------------------------------------------------------
'    If (Mnemonic = TOKEN_Z80_PUSH) _
'    Or (Mnemonic = TOKEN_Z80_POP) _
'    Then
'        'Is there any parameters to begin with?
'        If (TokenAttr And ATTR_Z80_PARAM) = 0 Then Call RaiseError( _
'            OZ80_ERROR_INVALID, "TODO" _
'        ): GoTo Finish
'
'        'We'll need to loop back to continue, so we'll use this as a flag
'        Dim PushPop As Boolean
'        Let PushPop = True
'
'PushPop:
'        'When a non-register is found, the Push/Pop list has ended
'        If (TokenAttr And ATTR_Z80_REGISTER) = 0 Then GoTo Finish
'
'        'Fetch a parameter
'        Call Z80GetParameter(Instruction.Param1, ChunkID)
'        '(If there was an error, the error number will already have been set, _
'          we only need to fold upwards to end the assembling)
'        If Error Then GoTo Finish
'
'        'If it's not a pushable/poppable register, then it's invalid
'        If (Instruction.Param1.Mask And PARAM_REGS_AF_BC_DE_HL_IXY) = 0 Then
'            Call RaiseError(OZ80_ERROR_INVALID, "TODO")
'            GoTo Finish
'        End If
'
'        'Add the param to the log
'        Call W32TLB.MemCopy( _
'            LogPtr + 74, StrPtr(ParamText(Instruction.Param1.LogText)), _
'        20)
'
'        'Is it IX/IY?
'        If (Instruction.Param1.Mask And PARAM_ANY_IXY) <> 0 Then
'            'Is it IX or IY?
'            If (Instruction.Param1.Mask And PARAM_ANY_IX) <> 0 _
'                Then Let Prefix = &HDD& _
'                Else Let Prefix = &HFD&
'            'We can already add this to the bytes to output
'            Let Bytes(0) = Prefix
'            'And include it in the disassembly log
'            Call W32TLB.MemCopy(LogPtr + 14, StrPtr(HexStr8(Prefix)), 4)
'        Else
'            Let Bytes(0) = 0
'        End If
'
'    '----------------------------------------------------------------------------------
'    Else
'        'Up to 3 parameters are supported to allow for the undocumented instructions, _
'         e.g. `SET 0 [IX+$8] A` which sets bit 0 of the address and copies the result _
'         to register A
'        If (TokenAttr And ATTR_Z80_PARAM) <> 0 Then
'            'Parameter Expressions might return an unknown Value _
'             (such as when a Label, which has not yet been addressed, is used)
'
'            'Get the first parameter
'            Call Z80GetParameter(Instruction.Param1, ChunkID)
'            '(If there was an error, the error number will already have been set, _
'              we only need to fold upwards to end the assembling)
'            If Error Then GoTo Finish
'
'            Dim Mask As OZ80_MASK
'
'            'Is Parameter 1 `A`?
'            If (Instruction.Param1.Mask And PARAM_REG_A) <> 0 Then
'                Let Mask = MASK_P1_A
'            'Is Parameter 1 `B|D|E|H|L`?
'            ElseIf (Instruction.Param1.Mask And ( _
'                PARAM_REG_B Or PARAM_REG_D Or PARAM_REG_E Or _
'                PARAM_REG_H Or PARAM_REG_L _
'            )) <> 0 Then
'                Let Mask = MASK_P1_BDEHL
'            'Is Parameter 1 `C` (separate due to being Register or Flag)
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_C) <> 0 Then
'                Let Mask = MASK_P1_C
'            'Is Parameter 1 `IXH|IXL|IYH|IYL`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REGS_IXYHL) <> 0 Then
'                Let Mask = MASK_P1_IXYHL
'            'Is Parameter 1 `I|R`?
'            ElseIf (Instruction.Param1.Mask And (PARAM_REG_I Or PARAM_REG_R)) <> 0 Then
'                Let Mask = MASK_P1_IR
'            'Is Parameter 1 a Value?
'            ElseIf (Instruction.Param1.Mask And PARAM_VAL) <> 0 Then
'                Let Mask = MASK_P1_VAL
'            'Is Parameter 1 `AF`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_AF) <> 0 Then
'                Let Mask = MASK_P1_AF
'            'Is Parameter 1 `BC`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_BC) <> 0 Then
'                Let Mask = MASK_P1_BC
'            'Is Parameter 1 `DE`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_DE) <> 0 Then
'                Let Mask = MASK_P1_DE
'            'Is Parameter 1 `HL`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_HL) <> 0 Then
'                Let Mask = MASK_P1_HL
'            'Is Parameter 1 `IX` or `IY`?
'            ElseIf (Instruction.Param1.Mask And (PARAM_REG_IX Or PARAM_REG_IY)) <> 0 Then
'                Let Mask = MASK_P1_IXY
'            'Is Parameter 1 `SP`?
'            ElseIf (Instruction.Param1.Mask And PARAM_REG_SP) <> 0 Then
'                Let Mask = MASK_P1_SP
'            'Is Parameter 1 `[HL|IX|IY]`? (i.e. `jp`)
'            ElseIf (Instruction.Param1.Mask And PARAM_MEM_HLIXY) <> 0 Then
'                Let Mask = MASK_P1_MEM_HLIXY
'            'Is Parameter 1 `NZ|Z|NC|C|PE|PO|P|M`? (Flags)
'            ElseIf (Instruction.Param1.Mask And PARAM_FLAGS) <> 0 Then
'                Let Mask = MASK_P1_NZZNCPOPEPM
'            Else
''                Stop
'            End If
'
'
'
'            'Use of IX / IY requires a prefix byte ($DD/$FD) before the opcode. _
'             We need to record which parameter is the IX/IY reference
'            Dim ParamIXY As oz80Param
'
'            'Is Param 1 an IX/IY parameter?
'            If (Instruction.Param1.Mask And PARAM_ANY_IXY) <> 0 Then
'                Let ParamIXY = Instruction.Param1
'            End If
'
'            'Is Param 1 an 8/16-bit parameter Value?
'            If (Instruction.Param1.Mask And PARAM_ANY_VAL) <> 0 Then
'                Let Instruction.HasValue = 1
'                Let Instruction.Value = Instruction.Param1.Value
'            End If
'
'            'Add param 1 to the log
'            Call W32TLB.MemCopy( _
'                LogPtr + 74, StrPtr(ParamText(Instruction.Param1.LogText)), _
'            20)
'
'            'Check if a second parameter exists
'            '..........................................................................
'            If (TokenAttr And ATTR_Z80_PARAM) <> 0 Then
'                'Fetch the second parameter
'                Call Z80GetParameter(Instruction.Param2, ChunkID)
'                If Error Then GoTo Finish
'
'
'                'Is Parameter 2 `A`?
'                If (Instruction.Param2.Mask And PARAM_REG_A) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_A
'                'Is Parameter 2 `B|C|D|E|H|L`??
'                ElseIf (Instruction.Param2.Mask And PARAM_REGS_ABCDEHL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_BCDEHL
'                'Is Parameter 2 `IXH|IXL|IYH|IYL`?
'                ElseIf (Instruction.Param2.Mask And PARAM_REGS_IXYHL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_IXYHL
'                'Is Parameter 2 `I|R`?
'                ElseIf (Instruction.Param2.Mask And (PARAM_REG_I Or PARAM_REG_R)) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_IR
'                'Is Parameter 2 a Value?
'                ElseIf (Instruction.Param2.Mask And PARAM_VAL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_VAL
'                'Is Parameter 2 `AF`? (as in `ex AF AF'`)
'                ElseIf (Instruction.Param2.Mask And PARAM_REG_AF) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_AF
'                'Is Parameter 2 `BC|DE|SP`?
'                ElseIf (Instruction.Param2.Mask And PARAM_REGS_BC_DE_SP) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_BC_DE_SP
'                'Is Parameter 2 `HL`?
'                ElseIf (Instruction.Param2.Mask And PARAM_REG_HL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_HL
'                'Is Parameter 2 `IX`?
'                ElseIf (Instruction.Param2.Mask And PARAM_REG_IX) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_IX
'                'Is Parameter 2 `IY`?
'                ElseIf (Instruction.Param2.Mask And PARAM_REG_IY) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_IY
'                'Is Parameter 2 `[BC]`? (also `[C]`) or `[DE]`
'                ElseIf (Instruction.Param2.Mask And PARAM_MEM_BC) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_MEM_BC_DE
'                'Is Parameter 2 `[HL]`?
'                ElseIf (Instruction.Param2.Mask And PARAM_MEM_HL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_MEM_HL
'                'Is Parameter 2 `[IX|IY]`?
'                ElseIf (Instruction.Param2.Mask And PARAM_MEM_IXY) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_MEM_IXY
'                'Is Parameter 2 `[$FF|$FFFF]`?
'                ElseIf (Instruction.Param2.Mask And PARAM_MEM_VAL) <> 0 Then
'                    Let Mask = Mask Or MASK_P2_MEM_VAL
'                Else
''                    Stop
'                End If
'
'
'                'Is Param 2 an IX/IY parameter?
'                If (Instruction.Param2.Mask And PARAM_ANY_IXY) <> 0 Then
'                    Let ParamIXY = Instruction.Param2
'                End If
'
'                'Is Param 2 an 8/16-bit parameter Value?
'                If (Instruction.Param2.Mask And PARAM_ANY_VAL) <> 0 Then
'                    Let Instruction.HasValue = 1
'                    Let Instruction.Value = Instruction.Param2.Value
'                End If
'
'                'Add param 2 to the log
'                Call W32TLB.MemCopy(LogPtr + 94, _
'                    StrPtr(ParamText(Instruction.Param2.LogText)), _
'                20)
'
'                'Check if a third parameter exists
'                '......................................................................
'                If (TokenAttr And ATTR_Z80_PARAM) <> 0 Then
'                    'Fetch the third parameter (very uncommon)
'                    'NOTE: Parameter 3 is only used for some undocumented instructions _
'                     and will not be IX/IY (though IX/IY can be in previous parameters)
'                    Call Z80GetParameter(Instruction.Param3, ChunkID)
'                    If Error Then GoTo Finish
'
'                    'Add param 3 to the log
'                    Call W32TLB.MemCopy(LogPtr + 114, _
'                        StrPtr(ParamText(Instruction.Param3.LogText)), _
'                    20)
'
'                    'Is Param 3 `A|B|C|D|E|H|L`?
'                    If (Instruction.Param3.Mask And PARAM_REGS_ABCDEHL) <> 0 Then
'                        Let Mask = Mask Or MASK_P3_ABCDEHL
'                    End If
'                End If
'            End If
'
'            'If there is any IX/Y parameter, setup the opcode prefix
'            If ParamIXY.Mask <> 0 Then
'                'Is it IX or IY?
'                If (ParamIXY.Mask And PARAM_ANY_IX) <> 0 _
'                    Then Let Prefix = &HDD& _
'                    Else Let Prefix = &HFD&
'                'We can already add this to the bytes to output
'                Let Bytes(0) = Prefix
'                'And include it in the disassembly log
'                Call W32TLB.MemCopy(LogPtr + 14, StrPtr(HexStr8(Prefix)), 4)
'
'                'Does it include an offset byte? i.e. "[IX+$FF]"
'                If (ParamIXY.Mask And PARAM_MEM_IXY) <> 0 Then
'                    Let Instruction.HasOffset = 1
'                    Let Offset = ParamIXY.Value
'
'                    'Include this in the disassembly log
'                    If Offset = INDEFINITE Then
'                        Call W32TLB.MemCopy(LogPtr + 26, StrPtr(HexStr8(-1)), 4)
'                    Else
'                        Let Bytes(2) = Offset
'                        Call W32TLB.MemCopy(LogPtr + 26, StrPtr(HexStr8(Offset)), 4)
'                    End If
'
'                    If Instruction.Param1.Defer = 1 Then
'                        If Prefix = &HDD& _
'                            Then Call W32TLB.MemCopy(LogPtr + 74, StrPtr("[IX+$??]  "), 20) _
'                            Else Call W32TLB.MemCopy(LogPtr + 74, StrPtr("[IY+$??]  "), 20)
'                    ElseIf Instruction.Param2.Defer = 1 Then
'                        If Prefix = &HDD& _
'                            Then Call W32TLB.MemCopy(LogPtr + 94, StrPtr("[IX+$??]  "), 20) _
'                            Else Call W32TLB.MemCopy(LogPtr + 94, StrPtr("[IY+$??]  "), 20)
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    '----------------------------------------------------------------------------------
'
'    With Chunks(ChunkID)
'        'If the disassembly log is not yet set for Z80, do so now
'        If ChunkMode <> MODE_Z80 Then
'            'If we're coming from binary data, we'll need to finalise that
'            If ChunkMode = MODE_DATA Then
'                'Have any binary bytes have been prepared, but not logged yet?
'                If ChunkDataIndex <> 0 Then
''                    Call .LogASM.Add(ChunkDataLog)
''                    Call .LogASM.Add(LOG_ASM_FOOT)
'                End If
'            End If
'            'Add the table header for Z80 disassembly logging
''            Call .LogASM.Add(LOG_ASM_COLS)
''            Call .LogASM.Add(LOG_ASM_HEAD)
'            Let ChunkMode = MODE_Z80
'        End If
'
'        'If the Chunk data hasn't started yet, make sure we're 1-based
'        Dim BinaryOffset As Long
'        If .Size = 0 Then Let BinaryOffset = 1 Else Let BinaryOffset = .Size + 1
'
'        Dim From As Long
'        Let From = BinaryOffset
'
'        If Bytes(0) <> 0 Then Let BinaryOffset = BinaryOffset + 1
'
'        'Write the current Chunk address on the disassembly log
'        Call W32TLB.MemCopy(LogPtr, StrPtr(HexStr16(.Size)), 8)
'    End With
'
'    Dim Opcode As Long
'
'    'See, BASIC does have a good use!
'    On (Mnemonic - [_TOKEN_Z80_BEGIN] + 1) GoTo _
'        iADC, iADD, iAND, iBIT, iCALL, iCCF, iCP, iCPD, iCPDR, iCPI, iCPIR, iCPL, _
'        iDAA, iDEC, iDI, iDJNZ, iEI, iEX, iEXX, iHALT, iIM, iIN, iINC, iIND, iINDR, _
'        iINI, iINIR, iJP, iJR, iLD, iLDD, iLDDR, iLDI, iLDIR, iNEG, iNOP, iOR, iOUT, _
'        iOUTD, iOTDR, iOUTI, iOTIR, iPOP, iPUSH, iRES, iRET, iRETI, iRETN, iRL, iRLA, _
'        iRLC, iRLCA, iRLD, iRR, iRRA, iRRCA, iRRC, iRRD, iRST, iSBC, iSCF, iSET, _
'        iSLA, iSRA, iSLL, iSRL, iSUB, iXOR
'
'    'TODO: ERROR: completely invalid token number!!
'    Stop
'
'InvalidParams:
'    '----------------------------------------------------------------------------------
'    'ERROR: Wrong parameters for Z80 instruction
'    Call RaiseError(OZ80_ERROR_INVALID, "TODO")
'    GoTo Finish
'
'iADC:
'    'Add with Carry _
'     ..................................................................................
'    '`ADC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REG_A, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
''    If (Mask And TEST_P1_A) <> 0 And (Mask Or TEST_P2_ABCDEHL_IXYHL_MEM_HLIXY) <> 0 Then
'        Let Opcode = X2 Or Y1 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_2Param
'    '`ADC A, $8`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_VAL) Then
'        Let Opcode = X3 Or Y1 Or Z6: GoTo Output8_2Param
'    '`ADC HL, BC|DE|HL|SP`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_HL, PARAM_REGS_BC_DE_HL_SP) Then
'        Let Opcode = X1 Or Q1 Or Z2 Or OpcodeP_RegisterPair(Instruction.Param2.Token)
'        GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iADD:
'    'Add: _
'     ..................................................................................
'    '`ADD HL, BC|DE|HL|SP`
'    If Z80TestParams(Instruction, PARAM_REG_HL, PARAM_REGS_BC_DE_HL_SP) Then
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param2.Token) Or Q1 Or Z1
'        GoTo Output0
'    '`ADD IX, BC|DE|IX|SP`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_IX, PARAM_REG_BC Or PARAM_REG_DE Or PARAM_REG_IX Or PARAM_REG_SP) Then
'        'Even though this is the same opcode as above (the prefix denotes "IX"), _
'         we still need to validate the params to ensure HL & IX are not mixed
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param2.Token) Or Q1 Or Z1
'        GoTo Output0
'    '`ADD IY, BC|DE|IY|SP`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_IY, PARAM_REG_BC Or PARAM_REG_DE Or PARAM_REG_IY Or PARAM_REG_SP) Then
'        'Even though this is the same opcode as above (the prefix denotes "IY"), _
'         we still need to validate the params to ensure HL & IY are not mixed
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param2.Token) Or Q1 Or Z1
'        GoTo Output0
'    '`ADD A, A|B|C|D|E|H|IXH|IYH|L|IYH|IYL|[HL|IX+$8|IY+$8]`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_2Param
'    '`ADD A, $8`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_VAL) Then
'        Let Opcode = X3 Or Z6: GoTo Output8_2Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iAND:
'    'Bitwise AND: _
'     ..................................................................................
'    '`AND A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y4 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputOffset_1Param
'    '`AND $8`
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Y4 Or Z6: GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iBIT:
'    'Test a bit: _
'     ..................................................................................
'    '`BIT 0-7, A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_VAL, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = X1 Or OpcodeY_Bit(Instruction.Param1.Value) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_2Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCALL:
'    'Call Procedure: _
'     ..................................................................................
'    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
'    If Z80TestParams(Instruction, PARAM_FLAGS, PARAM_VAL) Then
'        Let Opcode = X3 Or OpcodeY_Condition(Instruction.Param1.Token) Or Z4
'        GoTo Output16_2Param
'    '`CALL $16` - absolute call
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Q1 Or Z5: GoTo Output16_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCCF:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y7 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCP:
'    '..................................................................................
'    '`CP A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y7 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputOffset_1Param
'    '`CP $8`
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Y7 Or Z6: GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCPD:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y5 Or Z1: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCPDR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y7 Or Z1: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCPI:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y4 Or Z1: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCPIR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y6 Or Z1: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iCPL:
'    'Compliment: Flip the bits of the Accumulator _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y5 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iDAA:
'    'Decimal Adjust Accumulator: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y4 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iDEC:
'    'Decrement: _
'     ..................................................................................
'    '`DEC BC|DE|HL|SP|IX|IY`
'    If Z80TestParams(Instruction, PARAM_REGS_BC_DE_HL_SP_IXY) Then
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param1.Token) Or Q1 Or Z3
'        GoTo Output0
'    '`DEC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = OpcodeY_Register(Instruction.Param1.Mask) Or Z5
'        GoTo OutputOffset_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iDI:
'    'Disable Interrupts: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X3 Or Y6 Or Z3: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iDJNZ:
'    'Decrement, Jump if Not Zero: _
'     ..................................................................................
'    If Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = Y2
'        GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iEI:
'    'Enable Interrupts: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X3 Or Y7 Or Z3: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iEX:
'    'Exchange: _
'     ..................................................................................
'    '`EX AF AF'`
'    If Z80TestParams(Instruction, PARAM_REG_AF, PARAM_REG_AF) Then
'        Let Opcode = Y1: GoTo Output0
'    '`EX [SP], HL|IX|IY`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_SP, PARAM_REGS_HL_IXY) Then
'        Let Opcode = X3 Or Y4 Or Z3: GoTo Output0
'    '`EX DE, HL`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_DE, PARAM_REG_HL) Then
'        Let Opcode = X3 Or Y5 Or Z3: GoTo Output0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iEXX:
'    'Exchange All: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X3 Or Q1 Or P1 Or Z1: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iHALT:
'    'Halt: Stop the processor and wait for interrupts _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Y6 Or Z6: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iIM:
'    'Interrupt Mode: _
'     ..................................................................................
'    '`IM 0|1|2`
'    If Z80TestParams(Instruction, PARAM_VAL) Then
'        If Instruction.Param1.Value = 0 Then
'            Let Opcode = X1 Or Z6
'        ElseIf Instruction.Param1.Value = 1 Then
'            Let Opcode = X1 Or Y2 Or Z6
'        ElseIf Instruction.Param1.Value = 2 Then
'            Let Opcode = X1 Or Y3 Or Z6
'        Else
'            'ERROR: Wrong parameters for Z80 instruction
'            GoTo InvalidParams
'        End If
'        'The Value is baked into the Opcode, _
'         so doesn't need to be written out separately
'        GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iIN:
'    'Read from port: _
'     ..................................................................................
'    '`IN A, [$8]`
'    If Z80TestParams(Instruction, PARAM_REG_A, PARAM_MEM_VAL) Then
'        Let Opcode = X3 Or Y3 Or Z3: GoTo Output8_2Param
'    '`IN A|B|C|D|E|H|L, [C]`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_ABCDEHL, PARAM_MEM_BC) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param1.Mask): GoTo OutputED0
'    '`IN [C]`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_BC) Then
'        Let Opcode = X1 Or Y6: GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iINC:
'    'Increment: _
'     ..................................................................................
'    '`INC BC|DE|HL|SP|IX|IY`
'    If Z80TestParams(Instruction, PARAM_REGS_BC_DE_HL_SP_IXY) Then
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param1.Token) Or Z3: GoTo Output0
'    '`INC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = OpcodeY_Register(Instruction.Param1.Mask) Or Z4
'        GoTo OutputOffset_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iIND:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y5 Or Z2: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iINDR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y7 Or Z2: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iINI:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y4 Or Z2: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iINIR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y6 Or Z2: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iJP:
'    'Jump: _
'     ..................................................................................
'    '`JP [HL|IX|IY]` or `JP HL|IX|IY` - indirect absolute jump
'    If Z80TestParams(Instruction, PARAM_MEM_HLIXY Or PARAM_REGS_HL_IXY) Then
'        Let Opcode = X3 Or P2 Or Q1 Or Z1: GoTo Output0
'    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
'    ElseIf Z80TestParams(Instruction, PARAM_FLAGS, PARAM_VAL) Then
'        Let Opcode = X3 Or OpcodeY_Condition(Instruction.Param1.Token) Or Z2
'        GoTo Output16_2Param
'    '`JP $16` - absolute jump
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Z3: GoTo Output16_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iJR:
'    'Jump Relative: _
'     ..................................................................................
'    '`JR $8` - relative immediate jump
'    If Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = Y3: GoTo Output8_1Param
'    '`JR NZ|Z|NC|C, $8` - conditional relative immediate jump
'    ElseIf Z80TestParams(Instruction, PARAM_FLAGS_CZ, PARAM_VAL) Then
'        Let Opcode = OpcodeY_Condition(Instruction.Param1.Token) + Y4
'        GoTo Output8_2Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iLD:
'    'The Mother-Load: _
'     ..................................................................................
'    '`LD BC|DE|HL|SP|IX|IY, $16`
'    If Z80TestParams(Instruction, PARAM_REGS_BC_DE_HL_SP_IXY, PARAM_VAL) Then
'        Let Opcode = OpcodeP_RegisterPair(Instruction.Param1.Token) Or Z1
'        GoTo Output16_2Param
'    '`LD [BC], A`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_BC, PARAM_REG_A) Then
'        Let Opcode = Z2: GoTo Output0
'    '`LD [DE], A`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_DE, PARAM_REG_A) Then
'        Let Opcode = P1 Or Z2: GoTo Output0
'    '`LD A, [BC]`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_MEM_BC) Then
'        Let Opcode = Q1 Or Z2: GoTo Output0
'    '`LD A, [DE]`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_MEM_DE) Then
'        Let Opcode = P1 Or Q1 Or Z2: GoTo Output0
'    '`LD [$16], HL|IX|IY`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_VAL, PARAM_REGS_HL_IXY) Then
'        Let Opcode = P2 Or Z2: GoTo Output16_1Param
'    '`LD [$16], A`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_VAL, PARAM_REG_A) Then
'        Let Opcode = P3 Or Z2: GoTo Output16_1Param
'    '`LD HL|IX|IY, [$16]`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_HL_IXY, PARAM_MEM_VAL) Then
'        Let Opcode = P2 Or Q1 Or Z2: GoTo Output16_2Param
'    '`LD A, [$16]`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_MEM_VAL) Then
'        Let Opcode = P3 Or Q1 Or Z2: GoTo Output16_2Param
'    '`LD A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8], $8`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY, PARAM_VAL) Then
'        Let Opcode = OpcodeY_Register(Instruction.Param1.Mask) Or Z6: GoTo OutputOff8
'    'You can't combine IX and IY use in the same opcode, there's only one prefix, _
'     so these two following checks ensure that IX & IY are not used together
'    '`LD A|B|C|D|E|H|IXH|L|IXL[HL|IX+$8], A|B|C|D|E|H|IXH|L|IXL`
'    ElseIf Z80TestParams(Instruction, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IXHL Or PARAM_MEM_HL Or PARAM_MEM_IX, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IXHL _
'    ) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param1.Mask) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_1Param
'    '`LD A|B|C|D|E|H|IYH|L|IYL[HL|IY+$8], A|B|C|D|E|H|IYH|L|IYL`
'    ElseIf Z80TestParams(Instruction, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IYHL Or PARAM_MEM_HL Or PARAM_MEM_IY, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IYHL _
'    ) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param1.Mask) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_1Param
'    'You can't have `LD [IX+$8], [IX+$8]!
'    '`LD A|B|C|D|E|H|IXH|L|IXL, A|B|C|D|E|H|IXH|L|IXL|[HL|IX+$8]`
'    ElseIf Z80TestParams(Instruction, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IXHL, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IXHL Or PARAM_MEM_HL Or PARAM_MEM_IX _
'    ) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param1.Mask) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_2Param
'    '`LD A|B|C|D|E|H|IYH|L|IYL, A|B|C|D|E|H|IYH|L|IYL|[HL|IY+$8]`
'    ElseIf Z80TestParams(Instruction, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IYHL, _
'        PARAM_REGS_ABCDEHL Or PARAM_REGS_IYHL Or PARAM_MEM_HL Or PARAM_MEM_IY _
'    ) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param1.Mask) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_2Param
'    '`LD SP, HL|IX|IY`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_SP, PARAM_REGS_HL_IXY) Then
'        Let Opcode = X3 Or P3 Or Q1 Or Z1: GoTo Output0
'    '`LD [$16], BC|DE|SP`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_VAL, PARAM_REGS_BC_DE_SP) Then
'        Let Opcode = X1 Or OpcodeP_RegisterPair(Instruction.Param2.Token) Or Z3
'        GoTo OutputED16_1Param
'    '`LD BC|DE|SP, [$16]`
'    ElseIf Z80TestParams(Instruction, PARAM_REGS_BC_DE_SP, PARAM_MEM_VAL) Then
'        Let Opcode = X1 Or OpcodeP_RegisterPair(Instruction.Param1.Token) Or Q1 Or Z3
'        GoTo OutputED16_2Param
'    '`LD I, A`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_I, PARAM_REG_A) Then
'        Let Opcode = X1 Or Z7: GoTo OutputED0
'    '`LD R, A`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_R, PARAM_REG_A) Then
'        Let Opcode = X1 Or Y1 Or Z7: GoTo OutputED0
'    '`LD A, I`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_REG_I) Then
'        Let Opcode = X1 Or Y2 Or Z7: GoTo OutputED0
'    '`LD A, R`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_REG_R) Then
'        Let Opcode = X1 Or Y3 Or Z7: GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iLDD:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y5: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iLDDR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then: Let Opcode = X2 Or Y7: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iLDI:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y4: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iLDIR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y6: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iNEG:
'    'Negate the Accumulator: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Z4: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iNOP:
'    'No Operation: _
'     ..................................................................................
'    '`NOP` - Opcode is 0
'    If Z80TestParams(Instruction) Then: Let Opcode = 0: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOR:
'    'Bitwise OR: _
'     ..................................................................................
'    '`OR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y6 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputOffset_1Param
'    '`OR $8`
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Y6 Or Z6: GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOUT:
'    'Write to port: _
'     ..................................................................................
'    '`OUT [$8], A`
'    If Z80TestParams(Instruction, PARAM_MEM_VAL, PARAM_REG_A) Then
'        Let Opcode = X3 Or Y2 Or Z3: GoTo Output8_1Param
'    '`OUT [C], A|B|C|D|E|H|L`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_BC, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = X1 Or OpcodeY_Register(Instruction.Param2.Mask) Or Z1
'        GoTo OutputED0
'    '`OUT [C], 0`
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_BC, PARAM_VAL) Then
'        Let Opcode = X1 Or Y6 Or Z1: GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOUTD:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y5 Or Z3: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOTDR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y7 Or Z3: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOUTI:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y4 Or Z3: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iOTIR:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X2 Or Y6 Or Z3: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iPOP:
'    'Fetch from stack: _
'     ..................................................................................
'    '`POP AF|BC|DE|HL|IX|IY`
'    If Z80TestParams(Instruction, PARAM_REGS_AF_BC_DE_HL_IXY) Then
'        Let Opcode = X3 Or OpcodeP_RegisterPair(Instruction.Param1.Token) Or Z1
'        GoTo Output0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iPUSH:
'    'Push a register on to the stack: _
'     ..................................................................................
'    '`PUSH AF|BC|DE|HL|IX|IY`
'    If Z80TestParams(Instruction, PARAM_REGS_AF_BC_DE_HL_IXY) Then
'        Let Opcode = X3 Or OpcodeP_RegisterPair(Instruction.Param1.Token) Or Z5
'        GoTo Output0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRES:
'    'Reset a bit: _
'     ..................................................................................
'    '`RES 0-7, A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_VAL, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or OpcodeY_Bit(Instruction.Param1.Value) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_2Param
'    '`RES 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RES memory & register
'    ElseIf Z80TestParams(Instruction, PARAM_VAL, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = X2 Or OpcodeY_Bit(Instruction.Param1.Value) _
'                  Or OpcodeZ_Register(Instruction.Param3.Mask)
'        GoTo OutputCB_2Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRET:
'    'Return: _
'     ..................................................................................
'    '`RET NZ|Z|NC|C|PO|PE|P|M` - conditional return
'    If Z80TestParams(Instruction, PARAM_FLAGS) Then
'        Let Opcode = X3 Or OpcodeY_Condition(Instruction.Param1.Token): GoTo Output0
'    '`RET`
'    ElseIf Z80TestParams(Instruction) Then
'        Let Opcode = X3 Or Q1 Or Z1: GoTo Output0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRETI:
'    'Return from Interrupt: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Y1 Or Z5: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRETN:
'    'Return from NMI: _
'     ..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Z5: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRLA:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y2 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRL:
'    '..................................................................................
'    '`RL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y2 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`RL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RL memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y2 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRLCA:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRLC:
'    '..................................................................................
'    '`RLC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = OpcodeZ_Register(Instruction.Param1.Mask): GoTo OutputCB_1Param
'    '`RLC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RLC memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = OpcodeZ_Register(Instruction.Param2.Mask): GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRLD:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Y5 Or Z7: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRRA:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y3 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRR:
'    '..................................................................................
'    '`RR A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y3 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`RR [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RR memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y3 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRRCA:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y1 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRRC:
'    '..................................................................................
'    '`RRC A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y1 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`RRC [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented RRC memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y1 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRRD:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = X1 Or Y4 Or Z7: GoTo OutputED0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iRST:
'    '"Reset" - Fire specific interrupt vector: _
'     ..................................................................................
'    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
'    If Z80TestParams(Instruction, PARAM_VAL) Then
'        'TODO: Need to validate the number is in range at some point in the chain
'        'TODO: Best way to handle INDEFINITE?
'        If Instruction.Value <> INDEFINITE Then
'            Let Opcode = X3 Or OpcodeY_RST(Instruction.Value) Or Z7
'        Else
'            Call DeferExpression(ASM_RST, _
'                File, ChunkID, Instruction.Param1.Expression, BinaryOffset, 0 _
'            )
'        End If
'        'The Value is baked into the Opcode, _
'         so discard it and don't write a separate byte!
'        GoTo Output0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSBC:
'    '..................................................................................
'    '`SBC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REG_A, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y3 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputOffset_2Param
'    '`SBC A, $8`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_A, PARAM_VAL) Then
'        Let Opcode = X3 Or Y3 Or Z6: GoTo Output8_2Param
'    '`SBC HL, BC|DE|HL|SP`
'    ElseIf Z80TestParams(Instruction, PARAM_REG_HL, PARAM_REGS_BC_DE_HL_SP) Then
'        Let Opcode = X1 Or OpcodeP_RegisterPair(Instruction.Param2.Token) Or Z2
'        GoTo OutputED0
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSCF:
'    '..................................................................................
'    If Z80TestParams(Instruction) Then Let Opcode = Y6 Or Z7: GoTo Output0
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSET:
'    'Set a bit: _
'     ..................................................................................
'    '`SET 0-7, A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_VAL, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = X3 Or OpcodeY_Bit(Instruction.Param1.Value) _
'                  Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_2Param
'    '`SET 0-7, [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SET memory & register
'    ElseIf Z80TestParams(Instruction, PARAM_VAL, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = X3 Or OpcodeY_Bit(Instruction.Param1.Value) _
'                  Or OpcodeZ_Register(Instruction.Param3.Mask)
'        GoTo OutputCB_2Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSLA:
'    '..................................................................................
'    '`SLA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y4 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`SLA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLA memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y4 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSRA:
'    '..................................................................................
'    '`SRA A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y5 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`SRA [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRA memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y5 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSLL:
'    '..................................................................................
'    '`SLL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y6 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`SLL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SLL memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y6 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSRL:
'    '..................................................................................
'    '`SRL A|B|C|D|E|H|L|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEHL_MEM_HLIXY) Then
'        Let Opcode = Y7 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputCB_1Param
'    '`SRL [IX+$8|IY+$8], A|B|C|D|E|H|L` - undocumented SRL memory and copy to register
'    ElseIf Z80TestParams(Instruction, PARAM_MEM_IXY, PARAM_REGS_ABCDEHL) Then
'        Let Opcode = Y7 Or OpcodeZ_Register(Instruction.Param2.Mask)
'        GoTo OutputCB_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iSUB:
'    'Subtract: _
'     ..................................................................................
'    '`SUB A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y2 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputOffset_1Param
'    '`SUB $8`
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Y2 Or Z6: GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'iXOR:
'    'Bitwise XOR: _
'     ..................................................................................
'    '`XOR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|[HL|IX+$8|IY+$8]`
'    If Z80TestParams(Instruction, PARAM_REGS_ABCDEIXYHL_MEM_HLIXY) Then
'        Let Opcode = X2 Or Y5 Or OpcodeZ_Register(Instruction.Param1.Mask)
'        GoTo OutputOffset_1Param
'    '`XOR $8`
'    ElseIf Z80TestParams(Instruction, PARAM_VAL) Then
'        Let Opcode = X3 Or Y5 Or Z6: GoTo Output8_1Param
'    End If
'    'ERROR: Wrong parameters for Z80 instruction
'    GoTo InvalidParams
'
'    '----------------------------------------------------------------------------------
'    'Write the assembled bytes to the Chunk
'    '----------------------------------------------------------------------------------
'    'Output only the Opcode, no Value or Offset required:
'    '..................................................................................
'Output0:
'    Let Bytes(1) = Opcode
'    'Add the opcode to the disassembly log
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    Let ByteCount = 1
'    GoTo AddBytes
'
'    'Output an Opcode with optional Offset, no Value
'    '..................................................................................
'OutputOffset_1Param:
'    Let Bytes(1) = Opcode
'    'Add the opcode to the disassembly log
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.HasOffset = 0 Then
'        Let ByteCount = 1
'    Else
'        'If the offset is indefinite, defer the calculation
'        If Offset = INDEFINITE Then Call DeferExpression( _
'            ASM_JR, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'        )
'        Let ByteCount = 2
'    End If
'    GoTo AddBytes
'
'OutputOffset_2Param:
'    '..................................................................................
'    Let Bytes(1) = Opcode
'    'Add the opcode to the disassembly log
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.HasOffset = 0 Then
'        Let ByteCount = 1
'    Else
'        'If the offset is indefinite, defer the calculation
'        If Offset = INDEFINITE Then Call DeferExpression( _
'            ASM_JR, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'        )
'        Let ByteCount = 2
'    End If
'    GoTo AddBytes
'
'    'Output an Opcode with 8-bit Value
'    '..................................................................................
'Output8_1Param:
'    Let Bytes(1) = Opcode
'    'Add the opcode to the disassembly log
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    'If the Value cannot be calculated yet, we'll need to defer it
'    If Instruction.Value = INDEFINITE Then
'        'Add the indefinite value to the disassembly log and params list
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 74, StrPtr("$??       "), 20)
'        If (Mnemonic = TOKEN_Z80_JR) _
'        Or (Mnemonic = TOKEN_Z80_DJNZ) Then
'            Call DeferExpression( _
'                ASM_JR, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'            )
'        Else
'            Call DeferExpression( _
'                ASM_VAL8, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'            )
'        End If
'    Else
'        'Add the Value to the opcode bytes
'        Let Bytes(2) = Instruction.Value
'        'Add this to the disassembly log
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value)), 4)
'    End If
'
'    Let ByteCount = 2
'    GoTo AddBytes
'
'Output8_2Param:
'    '..................................................................................
'    Let Bytes(1) = Opcode
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    'If the Value cannot be calculated yet, we'll need to defer it
'    If Instruction.Value = INDEFINITE Then
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 94, StrPtr("$??       "), 20)
'        If Mnemonic = TOKEN_Z80_JR Or Mnemonic = TOKEN_Z80_DJNZ Then
'            Call DeferExpression( _
'                ASM_JR, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'            )
'        Else
'            Call DeferExpression( _
'                ASM_VAL8, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'            )
'        End If
'    Else
'        'Add the Value to the opcode bytes
'        Let Bytes(2) = Instruction.Value
'        'Add this to the disassembly log
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value)), 4)
'    End If
'
'    Let ByteCount = 2
'    GoTo AddBytes
'
'    'Output an Opcode with Value and optional Offset:
'    '..................................................................................
'OutputOff8:
'    Let ByteCount = 1
'    Let Bytes(1) = Opcode
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    'Is there an offset Value? (e.g. `[IX+$FF]`)
'    If Instruction.HasOffset = 1 Then
'        If Offset = INDEFINITE Then
'            Call DeferExpression( _
'                ASM_VAL8, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'            )
'        Else
'            Let Bytes(2) = Offset
'        End If
'
'        Let ByteCount = 2
'    End If
'
'    Let ByteCount = ByteCount + 1
'    If Instruction.Value = INDEFINITE Then
'        Call DeferExpression( _
'            ASM_VAL8, File, ChunkID, Instruction.Param2.Expression, _
'            BinaryOffset + (ByteCount - 2), From _
'        )
'    Else
'        Let Bytes(ByteCount) = Instruction.Value
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value)), 4)
'    End If
'
'    GoTo AddBytes
'
'Output16_1Param:
'    '..................................................................................
'    Let Bytes(1) = Opcode
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.Value = INDEFINITE Then
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 74, StrPtr("$????     "), 20)
'
'        Call DeferExpression( _
'            ASM_JP, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'        )
'    Else
'        Let Bytes(2) = (Instruction.Value And &HFF&)
'        Let Bytes(3) = ((Instruction.Value \ &H100&) And &HFF&)
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value And &HFF&)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8((Instruction.Value \ &H100&) And &HFF&)), 4)
'    End If
'
'    Let ByteCount = 3
'    GoTo AddBytes
'
'Output16_2Param:
'    '..................................................................................
'    Let Bytes(1) = Opcode
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.Value = INDEFINITE Then
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 74, StrPtr("$????     "), 20)
'
'        Call DeferExpression( _
'            ASM_JP, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'        )
'    Else
'        Let Bytes(2) = (Instruction.Value And &HFF&)
'        Let Bytes(3) = ((Instruction.Value \ &H100&) And &HFF&)
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value And &HFF&)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8((Instruction.Value \ &H100&) And &HFF&)), 4)
'    End If
'
'    Let ByteCount = 3
'    GoTo AddBytes
'
'OutputED0:
'    '..................................................................................
'    'Push the implied prefix on the list of bytes to output
'    Let Bytes(0) = &HED&: Let BinaryOffset = BinaryOffset + 1
'    Let Bytes(1) = Opcode
'
'    Call W32TLB.MemCopy(LogPtr + 14, StrPtr(HexStr8(&HED&)), 4)
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    Let ByteCount = 1
'    GoTo AddBytes
'
'OutputED16_1Param:
'    '..................................................................................
'    'Push the implied prefix on the list of bytes to output
'    Let Bytes(0) = &HED&: Let BinaryOffset = BinaryOffset + 1
'    Let Bytes(1) = Opcode
'
'    Call W32TLB.MemCopy(LogPtr + 14, StrPtr(HexStr8(&HED&)), 4)
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.Value = INDEFINITE Then
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8(-1)), 4)
'
'        Call DeferExpression( _
'            ASM_JP, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'        )
'    Else
'        Let Bytes(2) = (Instruction.Value And &HFF&)
'        Let Bytes(3) = ((Instruction.Value \ &H100&) And &HFF&)
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value And &HFF&)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8((Instruction.Value \ &H100&) And &HFF&)), 4)
'    End If
'
'    Let ByteCount = 3
'    GoTo AddBytes
'
'OutputED16_2Param:
'    '..................................................................................
'    'Push the implied prefix on the list of bytes to output
'    Let Bytes(0) = &HED&: Let BinaryOffset = BinaryOffset + 1
'    Let Bytes(1) = Opcode
'
'    Call W32TLB.MemCopy(LogPtr + 14, StrPtr(HexStr8(&HED&)), 4)
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.Value = INDEFINITE Then
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(-1)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8(-1)), 4)
'
'        Call DeferExpression( _
'            ASM_JP, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'        )
'    Else
'        Let Bytes(2) = (Instruction.Value And &HFF&)
'        Let Bytes(3) = ((Instruction.Value \ &H100&) And &HFF&)
'        Call W32TLB.MemCopy(LogPtr + 38, StrPtr(HexStr8(Instruction.Value And &HFF&)), 4)
'        Call W32TLB.MemCopy(LogPtr + 44, StrPtr(HexStr8((Instruction.Value \ &H100&) And &HFF&)), 4)
'    End If
'
'    Let ByteCount = 3
'    GoTo AddBytes
'
'OutputCB_1Param:
'    '..................................................................................
'    'The double-opcodes use $CB as the primary opcode and then a second opcode after _
'     the offset byte. It's fastest to handle and log this special case specifically
'    Let Bytes(1) = &HCB&
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(&HCB&)), 4)
'    Call W32TLB.MemCopy(LogPtr + 32, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.HasOffset = 1 Then
'        Let Bytes(3) = Opcode
'        Let ByteCount = 3
'        If Offset = INDEFINITE Then
'            Call DeferExpression( _
'                ASM_VAL8, File, ChunkID, Instruction.Param1.Expression, BinaryOffset, From _
'            )
'        End If
'    Else
'        Let Bytes(2) = Opcode
'        Let ByteCount = 2
'    End If
'    GoTo AddBytes
'
'OutputCB_2Param:
'    '..................................................................................
'    'The double-opcodes use $CB as the primary opcode and then a second opcode after _
'     the offset byte. It's fastest to handle and log this special case specifically
'    Let Bytes(1) = &HCB&
'    Call W32TLB.MemCopy(LogPtr + 20, StrPtr(HexStr8(&HCB&)), 4)
'    Call W32TLB.MemCopy(LogPtr + 32, StrPtr(HexStr8(Opcode)), 4)
'
'    If Instruction.HasOffset = 1 Then
'        Let Bytes(3) = Opcode
'        Let ByteCount = 3
'        If Offset = INDEFINITE Then
'            Call DeferExpression( _
'                ASM_VAL8, File, ChunkID, Instruction.Param2.Expression, BinaryOffset, From _
'            )
'        End If
'    Else
'        Let Bytes(2) = Opcode
'        Let ByteCount = 2
'    End If
'
'AddBytes:
'    '----------------------------------------------------------------------------------
'    'Now the assembled bytes have been queued, add them to the Chunk
'    With Chunks(ChunkID)
'        'Increase the Chunk accordingly
'        If Bytes(0) = 0 _
'            Then Let .Size = .Size + ByteCount _
'            Else Let .Size = .Size + ByteCount + 1
'
'        'If the Chunk's buffer is full then double the buffer size to compensate. _
'         This will cause the buffer to grow exponentially, i.e. 2 KB, 4 KB, 8 KB, ...
'        If .Size >= .Buffer Then
'            Let .Buffer = .Buffer * 2
'            ReDim Preserve .Data(1 To .Buffer) As Byte
'        End If
'
'        'Copy the bytes into the Chunk
'        If Bytes(0) = 0 Then
'            Call W32TLB.MemCopy( _
'                OutputPointer:=VarPtr(.Data(BinaryOffset)), _
'                SourcePointer:=VarPtr(Bytes(1)), _
'                  LengthBytes:=ByteCount _
'            )
'        Else
'            Call W32TLB.MemCopy( _
'                OutputPointer:=VarPtr(.Data(BinaryOffset - 1)), _
'                SourcePointer:=VarPtr(Bytes(0)), _
'                  LengthBytes:=ByteCount + 1 _
'            )
'        End If
'
'        'Add the disassembly to the log
''        Call .LogASM.Add(LogStr)
'    End With
'
'    '----------------------------------------------------------------------------------
'
'    'Don't push me, push a push-pop! _
'     (if doing the Push/Pop macro, loop back for more)
'    If PushPop Then GoTo PushPop
'
'Finish:
'    Call Profiler.EventEnd(EVENT_Z80ASSEMBLE)
'    '*** TIMER END ***
'End Sub
'
''Z80GetParameter : Fetch the type / value of a Z80 instruction parameter
''======================================================================================
''ReturnParam    | An `oz80_Param` structure to populate
''[ChunkID]      | The Chunk ID is used to apply scope to Values that require scope,
''               | such as Label names that assume the Chunk's Section name. If this is
''               | ommitted (such as Expressions outside of Chunks), an error is raised
''               | if a Scope-dependent Value appears
''======================================================================================
'Private Sub Z80GetParameter( _
'             ByRef ReturnParam As oz80Param, _
'    Optional ByRef ChunkID As Long = 0 _
')
'    'Default to "no Value", it could be a Register
'    Let ReturnParam.Value = INDEFINITE
'
'    '----------------------------------------------------------------------------------
'    'A memory reference, such as `[IX+$FF]` _
'     ----------------------------------------------------------------------------------
'    If TokenKind = TOKEN_SQUARE_OPEN Then
'        'ERROR: File ends just as the memory expression is opened
'        If NextTokenError() Then Exit Sub
'
'        'The first item in the memory expression can be a register _
'         ..............................................................................
'        If (TokenAttr And ATTR_Z80_REGISTER) <> 0 Then
'            'Copy the register type to the parameter being built
'            Let ReturnParam.Token = TokenKind
'
'            'The `IN` & `OUT` instructions accept only the C register, _
'             though in reality this is register pair BC
'            If TokenKind = TOKEN_Z80_C Then
'                Let ReturnParam.Mask = PARAM_MEM_BC
'                Let ReturnParam.LogText = PARAMTEXT_MEM_BC
'
'            '`LD` accepts registers BC & DE
'            ElseIf TokenKind = TOKEN_Z80_BC Then
'                Let ReturnParam.Mask = PARAM_MEM_BC
'                Let ReturnParam.LogText = PARAMTEXT_MEM_BC
'            ElseIf TokenKind = TOKEN_Z80_DE Then
'                Let ReturnParam.Mask = PARAM_MEM_DE
'                Let ReturnParam.LogText = PARAMTEXT_MEM_DE
'
'            'HL / IX & IY
'            ElseIf TokenKind = TOKEN_Z80_HL Then
'                Let ReturnParam.Mask = PARAM_MEM_HL
'                Let ReturnParam.LogText = PARAMTEXT_MEM_HL
'
'            'IX/IY can have an additional expression, e.g. `ld A, [IX+$FF]`
'            ElseIf TokenKind = TOKEN_Z80_IX Then
'                Let ReturnParam.Mask = PARAM_MEM_IX
'                Let ReturnParam.LogText = PARAMTEXT_MEM_IX
'                GoTo AddExpr
'            ElseIf TokenKind = TOKEN_Z80_IY Then
'                Let ReturnParam.Mask = PARAM_MEM_IY
'                Let ReturnParam.LogText = PARAMTEXT_MEM_IY
'                GoTo AddExpr
'
'            ElseIf TokenKind = TOKEN_Z80_SP Then
'                Let ReturnParam.Mask = PARAM_MEM_SP
'                Let ReturnParam.LogText = PARAMTEXT_MEM_SP
'            Else
'                'ERROR: register not allowed here
'                Stop
'            End If
'
'            'ERROR: File ends before the memory expression is closed
'            If NextTokenError() Then Exit Sub
'
'            GoTo EndScope
'
'AddExpr:    'Look for the `+` following the register name
'            If NextTokenError() Then Exit Sub
'
'            'If present, skip over it and check for an expression
'            If TokenKind = TOKEN_OPERATOR_ADD Then
'                'ERROR: Something must follow the "+"!
'                If NextTokenError() Then Exit Sub
'
'                'Remember where this Expression begins, should it be deferred
'                Let ReturnParam.Expression = TokenIndex
'
'                'Try and calculate the Expression
'                Call ParseExpression(ReturnParam.Value, ChunkID)
'                '(If there was an error in the Expression, the error event _
'                  will already have been raised, we only need to fold upwards)
'                If Error Then Exit Sub
'
'                'TODO: Error if number is over 16-bit
'                'TODO: Issue warning for rounding
'                'TODO: Validate maximum 8-bit? Where?
'
'                If ReturnParam.Value = INDEFINITE Then
'                    Let ReturnParam.Defer = 1
'                    'TODO: Special text for indefinite?
'                ElseIf ReturnParam.Mask = PARAM_MEM_IX Then
'                    Let ReturnParam.LogText = PARAMTEXT_MEM_IX_VAL + ReturnParam.Value
'                ElseIf ReturnParam.Mask = PARAM_MEM_IY Then
'                    Let ReturnParam.LogText = PARAMTEXT_MEM_IY_VAL + ReturnParam.Value
'                Else
'                    Let ReturnParam.LogText = PARAMTEXT_MEM_VAL16 + ReturnParam.Value
'                End If
'
'                'If this is a value for an IX/IY memory reference, _
'                 select the appropriate cached text for logging
'                If TokenKind = TOKEN_SQUARE_CLOSE Then
'                    Call NextToken
'                    Exit Sub
'                End If
'
'                'ERROR: Unexpected token in memory expression
'                Stop
'            End If
'
'        '..............................................................................
'        Else
'            'Is it an expression at all?
'            'TODO: ERROR: Not a valid expression
'            If (TokenAttr And ATTR_EXPR) = 0 Then Stop
'
'            'Besides registers, memory addresses are allowed
'            Let ReturnParam.Token = TOKEN_NUMBER
'
'            'Remember where this Expression begins, should it be deferred
'            Let ReturnParam.Expression = TokenIndex
'
'            'Try and calculate the expression
'            Call ParseExpression(ReturnParam.Value, ChunkID)
'            '(If there was an error in the Expression, the error event _
'              will already have been raised, we only need to fold upwards)
'            If Error Then Exit Sub
'
'            'TODO: Error if number is over 16-bit
'            'TODO: Issue warning for rounding
'
'            If ReturnParam.Value = INDEFINITE Then
'                Let ReturnParam.Defer = 1
'            Else
'                Let ReturnParam.LogText = PARAMTEXT_MEM_VAL16 + ReturnParam.Value
'            End If
'
'            'Any number will class as 16-bit since it's the maximum allowed
'            Let ReturnParam.Mask = PARAM_MEM_VAL
'        End If
'
'EndScope:
'        '..............................................................................
'        If TokenKind = TOKEN_SQUARE_CLOSE Then
'            Call NextToken
'            Exit Sub
'        End If
'        'ERROR: Unexpected token in memory expression
'        Stop
'
'    '----------------------------------------------------------------------------------
'    'A Register _
'     ----------------------------------------------------------------------------------
'    ElseIf (TokenAttr And ATTR_Z80_REGFLAG) <> 0 _
'    Then
'        'Copy the register type to the parameter being built
'        Let ReturnParam.Token = TokenKind
'
'        'Set the index for the cached text prepared for the register name
'        Let ReturnParam.LogText = PARAMTEXT_REGS + (TokenKind - [_TOKEN_REGS_BEGIN])
'
'        'Map the register token to the mask bit; token numbers are sequential _
'         whilst parameter masks are bit combinations allowing rapid comparisons _
'         of parameters with allowed options
'        Let ReturnParam.Mask = TokenRegister(TokenKind)
'
'        'Now move beyond the register token
'        Call NextToken
'
'    '----------------------------------------------------------------------------------
'    'Anything else classes as a calculable expression _
'     ----------------------------------------------------------------------------------
'    ElseIf (TokenAttr And ATTR_EXPR) <> 0 Then
'        'We manually build a number token, since we don't want to use the first _
'         token of the expression -- it might be a Label for instance and we can _
'         only send registers or numbers to be assembled into opcodes
'        Let ReturnParam.Token = TOKEN_NUMBER
'        Let ReturnParam.Mask = PARAM_VAL
'
'        'Remember where this Expression begins, should it be deferred
'        Let ReturnParam.Expression = TokenIndex
'
'        'Try and calculate the expression
'        Call ParseExpression(ReturnParam.Value, ChunkID)
'        '(If there was an error in the Expression, the error event _
'          will already have been raised, we only need to fold upwards)
'        If Error Then Exit Sub
'
'        'TODO: Error if number is over 16-bit
'        'TODO: Issue warning for rounding
'
'        If ReturnParam.Value = INDEFINITE Then
'            Let ReturnParam.Defer = 1
'        Else
'            Let ReturnParam.LogText = PARAMTEXT_VAL16 + ReturnParam.Value
'        End If
'
'    '----------------------------------------------------------------------------------
'    Else
'        'ERROR: Not a Z80 parameter
'        Call RaiseError(OZ80_ERROR_INVALID, "TODO")
'    End If
'End Sub
'
''Z80TestParams : Check if parameters are of the right type desired
''======================================================================================
''Instruction    | An `oz80Instruction` type containing the parameters
''[Test1]        | A bit-pattern of allowed parameters to check parameter 1 against
''[Test2]        | A bit-pattern of allowed parameters to check parameter 2 against
''[Test3]        | A bit-pattern of allowed parameters to check parameter 3 against
''---------------¦----------------------------------------------------------------------
''Returns        | True if all given parameters fulfill the tests, False otherwise
''======================================================================================
'Private Function Z80TestParams( _
'             ByRef Instruction As oz80Instruction, _
'    Optional ByRef Test1 As OZ80_PARAM, _
'    Optional ByRef Test2 As OZ80_PARAM, _
'    Optional ByRef Test3 As OZ80_PARAM _
') As Boolean
'    'Why call the function with no parameters? It's a consistent way to check when _
'     there *should* be no parameters for an instruction!
'
'    'Are we testing a first parameter?
'    If (Test1 = 0) Then
'        'If a parameter exists, but there shouldn't be, return False
'        If (Instruction.Param1.Mask <> 0) Then Exit Function
'        'If there is no first parameter, pass True
'        If (Instruction.Param1.Mask = 0) Then GoTo Finish
'    End If
'
'    '1st Parameter Test: _
'     ----------------------------------------------------------------------------------
'
'    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
'    If (Instruction.Param1.Mask And Test1) = 0 Then Exit Function
'
'    'Are we testing a second parameter?
'    If (Test2 = 0) Then
'        'If there shouldn't be a second parameter, return False
'        If (Instruction.Param2.Mask <> 0) Then Exit Function
'        'If there is no second parameter, the test has passed
'        If (Instruction.Param2.Mask = 0) Then GoTo Finish
'    End If
'
'    '2nd Parameter Test: _
'     ----------------------------------------------------------------------------------
'
'    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
'    If (Instruction.Param2.Mask And Test2) = 0 Then Exit Function
'
'    'If there shouldn't be a third parameter, return False
'    If (Test3 = 0) And (Instruction.Param3.Mask <> 0) Then Exit Function
'    'If there is no third parameter, the test has passed
'    If (Test3 = 0) And (Instruction.Param3.Mask = 0) Then GoTo Finish
'
'    '3rd Parameter Test: _
'     ----------------------------------------------------------------------------------
'
'    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
'    If (Instruction.Param3.Mask And Test3) = 0 Then Exit Function
'
'Finish:
'    Let Z80TestParams = True
'End Function
'
''GetListAsDictionary : Reads a whole list into a Dictionary
''======================================================================================
'Private Function GetListAsDictionary( _
'    Optional ByRef ChunkID As Long = 0 _
') As Scripting.Dictionary
'    'Instantiate the return object
'    Set GetListAsDictionary = New Scripting.Dictionary
'
'    'As long as the list continues, keep fetching Values
'    Do While (TokenAttr And ATTR_EXPR) <> 0
'        'Parse the Expression; if it's indefinite, _
'         leave that up to the caller to handle
'        Dim Value As Double
'        Call ParseExpression(Value, ChunkID)
'        '(If there was an error in the expression, the error event _
'         will already have been raised, we only need to fold upwards)
'        If Error Then Exit Function
'
'        Call GetListAsDictionary.Add( _
'            GetListAsDictionary.Count, Value _
'        )
'    Loop
'End Function
'
''GetText : Reads a list, interpretting Values as Unicode points
''======================================================================================
''ReturnText     | A bluString object that will be populated with the text
''---------------¦----------------------------------------------------------------------
''Returns        | The error number
''======================================================================================
'Private Function GetText( _
'             ByRef ReturnText As bluString, _
'    Optional ByRef ChunkID As Long = 0 _
') As OZ80_ERROR
'    'If the bluString is not instantiated, do so now
'    If ReturnText Is Nothing Then Set ReturnText = New bluString
'
'    'If the current token is not a list item, then the list has come to an end
'    Do While (TokenAttr And ATTR_LIST) <> 0
'
'        'If the current token is text, append it to what we have
'        If TokenKind = TOKEN_TEXT Then
'            Call ReturnText.Append(TokenText)
'
'        'Otherwise it'll be a calculable Expression, _
'         hopefully providing an ASCII value
'        ElseIf (TokenAttr And ATTR_EXPR) <> 0 Then
'            'Get the Expression Value
'            Dim Value As Double
'            'If there were an error in parsing the Expression, fold up
'            'TODO: Does this need a Section/Label reference?
'            Call ParseExpression(Value, ChunkID)
'            If Error Then GoTo ErrOut
'
'            'TODO: Indefnite values? Allowed or not in Text? _
'                   If we allow for text mapping with different sized strings then _
'                   the string size cannot be determinate
'
'            'Any UCS-2 value is allowed, it's up to the caller _
'             to do text-mapping and validation
'            If (Value < 0) Or (Value > 65535) Then
'                Call RaiseError(OZ80_ERROR_INVALID, "TOOD")
'                GoTo ErrOut
'            End If
'        End If
'
'        'Continue reading the list. Note that when the file ends, so does the list _
'         and this is not an error as far as this function is concerned. If the parent _
'         is within a scope and expecting the closing bracket, then it's up to them _
'         to raise the error
'        Call NextToken
'    Loop
'
'    'The resultant string won't be changing any more
'    Call ReturnText.TrimBuffer
'ErrOut:
'    'Return the error number
'    Let GetText = Error
'End Function
'
''ParseExpression : Evaluate a calculation into a Value
''======================================================================================
''ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
''               | Note that this is a 64-bit floating-pointer number, so it's up to
''               | the caller to validate the number according to their needs
''[ChunkID]      | The Chunk ID is used to apply scope to Values that require scope,
''               | such as Label names that assume the Chunk's Section name. If this is
''               | ommitted (such as Expressions outside of Chunks), an error is raised
''               | if a Scope-dependent Value appears
''======================================================================================
'Private Sub ParseExpression( _
'             ByRef ReturnValue As Double, _
'    Optional ByRef ChunkID As Long = 0 _
')
'    Dim Operand1 As Double              'Left-hand value in a calculation
'    Dim Operand2 As Double              'Right-hand value in a calculation
'    Dim Operator As OZ80_TOKEN          'Operator to use
'
'    'Begin with fetching the first operand
'    Call ParseValue(Operand1, ChunkID)
'    If Error Then Exit Sub
'    'For the moment the Expression's calculated Value will be this unless we find _
'     an Operator to extend the Expression
'    Let ReturnValue = Operand1
'
'    'The file could legitimately end with an Expression, so don't error out if the _
'     token stream comes to an end since the Expression is valid. The parent scopes _
'     will handle an unexpected end error
'    If EOF Then Exit Sub
'
'    'If next token is not an Operator assume the Expression has finished, _
'      the parent will handle whatever the next token is used for
'    If (TokenAttr And ATTR_OPERTATOR) = 0 Then Exit Sub
'
'Continue:
'    'Remember the Operator for after we have operand 2 to work with
'    Let Operator = TokenKind
'
'    'Get the second operand
'    'TODO: Error - unexpected end
'    If EOF Then Stop
'    Call NextToken
'    'Cannot have an operator hanging there without an operand following
'    Call ParseValue(Operand2, ChunkID)
'    If Error Then Stop
'
'    'If either operand were a yet-unknowable value (such as a label or property), _
'     then don't bother applying the operator, but do continue the expression so that _
'     we may reach the end and continue processing the tokens that follow
'    If (Operand1 = INDEFINITE) _
'    Or (Operand2 = INDEFINITE) Then
'        'Mark expression result as indefinite and skip applying the operator
'        Let ReturnValue = INDEFINITE
'
'    ElseIf Operator = TOKEN_OPERATOR_ADD Then
'        'TODO: Check for overflow beforehand
'        Let ReturnValue = Operand1 + Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_SUB Then
'        'TODO: Check for overflow beforehand
'        Let ReturnValue = Operand1 - Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_MUL Then
'        'TODO: Check for overflow beforehand
'        Let ReturnValue = Operand1 * Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_DIV Then
'        'TODO: Check for divide by 0
'        Let ReturnValue = Operand1 / Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_POW Then
'        'TODO: Check for overflow
'        Let ReturnValue = Operand1 ^ Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_MOD Then
'        Let ReturnValue = Operand1 Mod Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_AND Then
'        Let ReturnValue = Operand1 And Operand2
'
'    ElseIf Operator = TOKEN_OPERATOR_OR Then
'        Let ReturnValue = Operand1 Or Operand2
'
'    Else
'        Stop
'    End If
'
'    'Does the expression continue?
'    If (TokenAttr And ATTR_OPERTATOR) <> 0 Then
'        'Set the current result as the first operand and go around again
'        Let Operand1 = ReturnValue
'        GoTo Continue
'    End If
'End Sub
'
''ParseValue : Get a single Value as part of an Expression
''======================================================================================
''ReturnValue    | Gets set to the calculated Value, or `INDEFINITE`.
''               | Note that this is a 64-bit floating-pointer number, so it's up to
''               | the caller to validate the number according to their needs
''[ChunkID]      | The Chunk ID is used to apply scope to Values that require scope,
''               | such as Label names that assume the Chunk's Section name. If this is
''               | ommitted (such as Expressions outside of Chunks), an error is raised
''               | if a Scope-dependent Value appears
''======================================================================================
'Private Sub ParseValue( _
'             ByRef ReturnValue As Double, _
'    Optional ByRef ChunkID As Long = 0 _
')
'    '----------------------------------------------------------------------------------
'    'A nested Expression?
'    '----------------------------------------------------------------------------------
'    If TokenKind = TOKEN_PARENS_OPEN Then
'        'Move over the open parentheses. If the file ends here, it's an error
'        If NextTokenError() Then Exit Sub
'
'        'Parse the nested Expression and return its Value as our own
'        Call ParseExpression(ReturnValue, ChunkID)
'        If Error Then Exit Sub
'
'        'The nested Expression has to be closed
'        If TokenKind <> TOKEN_PARENS_CLOSE Then Call RaiseError( _
'            OZ80_ERROR_EXPECTED, "TODO" _
'        ): Exit Sub
'
'        'Move over the closing parentehses. The Expression can end here, so it's not _
'         necessarily an error if the file suddenly ends. The parent can check for _
'         any scopes remaining open
'        Call NextToken
'
'        'The token pointer will already be at the next token to read, _
'         so we exit without moving it forward any more
'        Exit Sub
'
'    '----------------------------------------------------------------------------------
'    'The `HI` function -- returns the high-byte of a 16-bit Value
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_HI Then
'        'Move over the `HI` keyword. If the file ends here, it's an error
'        If NextTokenError() Then Exit Sub
'
'        'Fetch the parameter to the `HI` function
'        Call ParseExpression(ReturnValue, ChunkID)
'        If Error Then Exit Sub
'        'If that is indefinite, so will our result
'        If ReturnValue = INDEFINITE Then Exit Sub
'
'        'Check that this is no more than a 16-bit number
'        'TODO: Z80/16-bit specific error message
'        If ReturnValue > 65535 Then Call RaiseError( _
'            OZ80_ERROR_OVERFLOW, STR_ERR_OVERFLOW _
'        ): Exit Sub
'
'        'Extract the high-byte
'        Let ReturnValue = (ReturnValue And &HFF00&) \ &H100&
'        'The token pointer will already be at the next token to read, _
'         so we exit without moving it forward any more
'        Exit Sub
'
'    '----------------------------------------------------------------------------------
'    'The `LO` function -- returns the low-byte of a 16-bit Value
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_LO Then
'        'Move over the `LO` keyword. If the file ends here, it's an error
'        If NextTokenError() Then Exit Sub
'
'        'Fetch the parameter to the `LO` function
'        Call ParseExpression(ReturnValue, ChunkID)
'        If Error Then Exit Sub
'        'If that is indefinite, so will be our result
'        If ReturnValue = INDEFINITE Then Exit Sub
'
'        'Check that this is no more than a 16-bit number _
'         (despite this not being an issue to the low 8-bits, we prevent numbers _
'          higher than 16-bit to protect against unintended behaviour)
'        'TODO: Z80/16-bit specific error message
'        If ReturnValue > 65535 Then Call RaiseError( _
'            OZ80_ERROR_OVERFLOW, STR_ERR_OVERFLOW _
'        ): Exit Sub
'
'        'Extract the low-byte
'        Let ReturnValue = ReturnValue And &HFF&
'        'The token pointer will already be at the next token to read, _
'         so we exit without moving it forward any more
'        Exit Sub
'
'    '----------------------------------------------------------------------------------
'    'Negative Numbers
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_OPERATOR_SUB Then
'        'If the negative sign preceeds, remember this until the full number has _
'         been calculated. ONLY decimal numbers can be negative, you cannot negate _
'         hexadecimal/binary numbers or named-items (Constants, RAM names, Labels etc.)
'        Dim IsNegative As Boolean: Let IsNegative = True
'        'Move forward past the negative sign. If the file ends here, it's an error
'        If NextTokenError() Then Exit Sub
'        'Continue parsing the number...
'        GoTo GetNumber
'
'    '----------------------------------------------------------------------------------
'    'A decimal number:
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_NUMBER Then
'GetNumber:
'        'Begin with getting the given value
'        Let ReturnValue = TokenValue
'        'Was there a negative sign preceeding this number?
'        If IsNegative Then Let ReturnValue = -ReturnValue
'
'        'Move over the number. If the file ends here, the Value is still valid, _
'         so no error. The parent can check if any scope has been left open. We skip _
'         checking for a number prefix, but we do need to apply any negative sign
'        If EOF Then
'            Exit Sub
'        Else
'            Call NextToken
'
'            'Check for a number suffixes, e.g. `KB`
'            '..........................................................................
'            If TokenKind = TOKEN_SUFFIX_K Then
'                'Check for min/max number you can multiply by 1'000 without overflowing
'                If (ReturnValue < NUMBER_MIN_K) Or (ReturnValue > NUMBER_MAX_K) Then
'                    Call RaiseError(OZ80_ERROR_OVERFLOW, STR_ERR_OVERFLOW)
'                    Exit Sub
'                End If
'                'Multiply by 1'000
'                Let ReturnValue = ReturnValue * 1000
'                'Move over the prefix
'                Call NextToken
'
'            '..........................................................................
'            ElseIf TokenKind = TOKEN_SUFFIX_KB Then
'                'Check for min/max number you can multiply by 1'024 without overflowing
'                If (ReturnValue < NUMBER_MIN_KB) Or (ReturnValue > NUMBER_MAX_KB) Then
'                    Call RaiseError(OZ80_ERROR_OVERFLOW, STR_ERR_OVERFLOW)
'                    Exit Sub
'                End If
'                'Multiply by 1'024
'                Let ReturnValue = ReturnValue * 1024
'            Else
'                'If the token is not a prefix, exit this function without moving the _
'                 token pointer forward (as we've already done that)
'                Exit Sub
'            End If
'        End If
'
'    '----------------------------------------------------------------------------------
'    'Defined Name (Constants, Sections, Labels, RAM Names &c.)
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_NAME Then
'        'Constants:
'        If (TokenAttr And ATTR_NAMES) = ATTR_NAME_CONST Then
'            'ERROR: Constant not yet defined. Constants cannot be indefinite
'            If Not Values.Exists(TokenText) Then Call RaiseError(OZ80_ERROR_INDEFINITE, _
'                "Constant cannot be indefinite" _
'            ): Exit Sub
'
'            'Retrieve the Constant value
'            Let ReturnValue = Values.Item(TokenText)
'
'        ElseIf (TokenAttr And ATTR_NAMES) = ATTR_NAME_SECTION Then
'            'Already fully-qualified names are used to reference between Sections
'            If Not Values.Exists(TokenText) Then
'                'Return Indefinite
'                Let ReturnValue = INDEFINITE
'            Else
'                'Retrieve the Label address
'                Let ReturnValue = Values.Item(TokenText)
'                If ReturnValue < 0 Then ReturnValue = INDEFINITE
'            End If
'
'        ElseIf (TokenAttr And ATTR_NAME_LABEL) <> 0 Then
'            'If no scope is provided, we cannot qualify local names to global names
'            If ChunkID = 0 Then
'                'In this case you can only refer to Sectionless Labels
'                If Values.Exists(TokenText) _
'                    Then Let ReturnValue = Values.Item(TokenText) _
'                    Else Let ReturnValue = INDEFINITE
'                If ReturnValue < 0 Then ReturnValue = INDEFINITE
'
'            'The "self" Label (`:`) refers to the Chunk's own address
'            ElseIf TokenValue = 0 Then
'                'TODO: ChunkID = 0?
'                'Get the Chunk's address (will be indefinite prior to ROM layout)
'                Let ReturnValue = Chunks(ChunkID).AddrSlot
'                If ReturnValue < 0 Then ReturnValue = INDEFINITE
'
'            'Check for Labels within the current Section scope
'            ElseIf Values.Exists(Chunks(ChunkID).SectionName & TokenText) Then
'                Let ReturnValue = Values.Item(Chunks(ChunkID).SectionName & TokenText)
'                If ReturnValue < 0 Then ReturnValue = INDEFINITE
'
'            'Check for global Labels
'            ElseIf Values.Exists(TokenText) Then
'                Let ReturnValue = Values.Item(TokenText)
'                If ReturnValue < 0 Then ReturnValue = INDEFINITE
'            Else
'                Let ReturnValue = INDEFINITE
'            End If
'
'        ElseIf (TokenAttr And ATTR_NAMES) = ATTR_NAME_PROPERTY Then
'            'Property names have to be expanded to their fully-qualified form
'            If Not Values.Exists(Chunks(ChunkID).FullName & TokenText) Then
'                'Return Indefinite
'                Let ReturnValue = INDEFINITE
'            Else
'                'Retrieve the sub-routine address
'                Let ReturnValue = Values.Item(Chunks(ChunkID).FullName & TokenText)
'            End If
'
'        Else
'            If Not Values.Exists(TokenText) _
'                Then Let ReturnValue = INDEFINITE _
'                Else Let ReturnValue = Values.Item(TokenText)
'        End If
'
'    '----------------------------------------------------------------------------------
'    'The `START` keyword refers to the System's boot vector
'    '----------------------------------------------------------------------------------
'    ElseIf TokenKind = TOKEN_KEYWORD_START Then
'        'TODO: This is determined by the System, so we can forgo having to reference it
'        If Not Values.Exists(STR_START) Then
'            'Return Indefinite
'            Let ReturnValue = INDEFINITE
'        Else
'            'Retrieve the START address
'            Let ReturnValue = Values.Item(STR_START)
'        End If
'
'    '----------------------------------------------------------------------------------
'    Else
'        Debug.Print TokenNames(TokenKind)
'
'        'If there's a Z80 register in an Value, it's probably because the user has _
'         used parentheses for a Z80 memory reference instead of square brackets
'        If (TokenAttr And ATTR_Z80_REGISTER) <> 0 Then
'            Call RaiseError(OZ80_ERROR_INVALID, _
'                "Z80 registers are not valid within expressions. Remeber that Z80 " & _
'                "memory references use square brackets, not parentheses." _
'            ): Exit Sub
'        End If
'
'        Call RaiseError(OZ80_ERROR_INTERNAL, "NOPE")
'
'        Stop
'        Exit Sub
'    End If
'
'    'Move over the processed token, the parent expects the token pointer _
'     to be on the next unread token
'    Call NextToken
'End Sub

''Chunk_Begin : Begin a new Chunk to hold assembled Z80 code / data
''======================================================================================
''LabelName      | String of the fully-qualified Label name, i.e. "::section:label"
''[SectionIndex] | Which Section this belongs to, index of `Sections` array
''[Address]      | Optional address to place the Chunk in ROM
''---------------¦----------------------------------------------------------------------
''Returns        | The index number of the new Chunk created
''======================================================================================
'Private Function Chunk_Begin( _
'             ByRef LabelName As String, _
'    Optional ByRef SectionID As Long = -1, _
'    Optional ByRef Address As Double = -1 _
') As Long
'    'Get the Section name if a SectionID is provided
'    Dim SectionName As String
'    If SectionID <> -1 Then Let SectionName = Sections(SectionID).Name
'
'    'Register the Label name in the database
'    Dim FullName As String
'    Let FullName = SectionName & LabelName
'    Call AddUpdateLabel(FullName, Address)
'
'    Let ChunkCount = ChunkCount + 1
'    ReDim Preserve Chunks(1 To ChunkCount) As oz80Chunk
'    Let Chunk_Begin = ChunkCount
'
'    'Apply the details we have
'    With Chunks(Chunk_Begin)
'        Let .FullName = FullName
'        Let .LabelName = LabelName
'        Let .SectionName = SectionName
'        Let .SectionID = SectionID
'
'        Let .AddrReal = CLng(Address)
'        Let .AddrSlot = CLng(Address)
'
'        If SectionID >= 0 Then
'            Let .Slot = Sections(SectionID).Slot
'
'            Dim SectionChunks As Long
'            Let SectionChunks = Sections(SectionID).ChunkCount + 1
'            ReDim Preserve Sections(SectionID).Chunks(1 To SectionChunks) As Long
'            Let Sections(SectionID).Chunks(SectionChunks) = ChunkCount
'            Let Sections(SectionID).ChunkCount = SectionChunks
'        End If
'
'        'Create a buffer of 1K to hold the assembled data
'        Let .Buffer = 1024
'        ReDim .Data(1 To 1024) As Byte
'        'As we assemble, we'll also provide a disassembly listing
'        Set .LogASM = New bluArrayStrings
'
'        Set .Properties = New Scripting.Dictionary
'    End With
'End Function

''AddUpdateLabel: Blindly adds to or updates the Label database
''======================================================================================
''LabelName      | The fully-qualified Label name, i.e. including the Section name
''[Value]        | The new address to set for the Label
''======================================================================================
'Private Sub AddUpdateLabel( _
'    ByRef LabelName As String, _
'    ByRef Value As Double _
')
'    'Add a new label or update an existing label value?
'    If Not Values.Exists(LabelName) _
'        Then Call Values.Add(LabelName, Value) _
'        Else Let Values.Item(LabelName) = Value
'End Sub

''ROMLayout : Calculate the placement of the code/data Chunks in the output ROM
''======================================================================================
'Private Sub ROMLayout()
'    '*** TIMER BEGIN ***
'    Call Profiler.EventBegin(EVENT_ROM_LAYOUT)
'
'    'How many bytes do we have?
''    Call RaiseMessage(OZ80_LOG_INFO, _
''        "Total Chunks Size: " & Format$(ChunksSize, "#,0") & " bytes" _
''    )
'
'    'Assign the default ROM space
'    ReDim ROM(0 To ROMSize - 1) As Byte
'
'    'Log the Sections defined
'    Dim s As Long, c As Long
''    For s = 1 To SectionCount
''        Call RaiseMessage(OZ80_LOG_INFO, _
''            "Section " & Sections(s).Name & ": " & _
''            Format$(Sections(s).Size, "#,0") & " bytes" _
''        )
''    Next s
'
''    Call RaiseMessage(OZ80_LOG_INFO, "ROM Layout:")
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
''    Call RaiseMessage(OZ80_LOG_INFO, "ADDRESS  |   SIZE | LABEL")
''    Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'    '----------------------------------------------------------------------------------
'    '[1]: Place the Interrupt Procedures
'    '----------------------------------------------------------------------------------
'    If Chunk_Interrupt00 <> 0 Then
'        'TODO: Check for clash with START routine on Master System
'        Call ROMLayoutChunk(Chunk_Interrupt00, 0)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt08 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt08, &H8&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt10 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt10, &H10&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt18 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt18, &H18&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt20 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt20, &H20&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt28 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt28, &H28&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt30 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt30, &H30&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt38 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt38, &H38&)
'        If Error Then GoTo FinishLayout
'    End If
'    If Chunk_Interrupt66 <> 0 Then
'        Call ROMLayoutChunk(Chunk_Interrupt66, &H66&)
'        If Error Then GoTo FinishLayout
'    End If
'
'    '----------------------------------------------------------------------------------
'    '[2]: Place the ROM Header
'    '----------------------------------------------------------------------------------
'
'    '----------------------------------------------------------------------------------
'    '[3]: Place the START Procedure
'    '----------------------------------------------------------------------------------
'    'The `START` Procedure is where the System begins running code first, _
'     usually this is $0000, but can be elsewhere on other Systems
'    If Chunk_Start <> 0 Then
''        Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'        'TODO: Check for clash with INTERRUPT 0 routine on Master System
'        Call ROMLayoutChunk(Chunk_Start, 0)
'        If Error Then GoTo FinishLayout
'    End If
'
'    '----------------------------------------------------------------------------------
'    '[4]: Place the 1 KB SEGA Slot
'    '----------------------------------------------------------------------------------
'    'The special SEGA 1 KB Slot ($0000-$03FF) stays paged-in at all times regardless _
'     of which Bank is in Slot 0 ($0000-$3FFF). Therefore we have to place these Chunks _
'     before other Un/Sectioned code as all data MUST appear within the first 1 KB of _
'     the ROM
''    Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'    For s = 1 To SectionCount
'        If Sections(s).Slot = SLOT1KB Then
'            'Walk through the Chunks in this Slot
'            For c = 1 To Sections(s).ChunkCount
'                'Place this Chunk
'                Call ROMLayoutChunk(Sections(s).Chunks(c))
'                If Error Then GoTo FinishLayout
'            Next c
'            'Record the base address & bank of the Section
'            Call AddUpdateLabel(Sections(s).Name, 0)
'            Call AddUpdateLabel(Sections(s).Name & ".bank", 0)
'        End If
'    Next s
'
'    '----------------------------------------------------------------------------------
'    '[5]: Place un-Sectioned Chunks
'    '----------------------------------------------------------------------------------
''    Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'    For c = 1 To Sections(0).ChunkCount
'        Call ROMLayoutChunk(Sections(0).Chunks(c))
'        If Error Then GoTo FinishLayout
'    Next c
'
'    '----------------------------------------------------------------------------------
'    '[6]: Place un-paged Sections
'    '----------------------------------------------------------------------------------
'    'Chunks that are assigned to Sections, but which Sections do not assign a Slot _
'     pattern (for paging) are treated the same as un-Sectioned code and lumped into _
'     the first 48 KB of ROM
''    Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'    'Go through the defined Sections looking for Sections with no Slots assigned
'    For s = 1 To SectionCount
'        'Is this Section Slotless?
'        If Sections(s).Slot = 0 Then
'            'Walk through the Chunks in this Slot
'            For c = 1 To Sections(s).ChunkCount
'                'Place this Chunk
'                Call ROMLayoutChunk(Sections(s).Chunks(c))
'                If Error Then GoTo FinishLayout
'            Next c
'            If Sections(s).ChunkCount <> 0 Then
'                'Record the base address & bank of the Section
'                Call AddUpdateLabel(Sections(s).Name, _
'                    CDbl(Chunks(Sections(s).Chunks(1)).AddrSlot) _
'                )
'                Call AddUpdateLabel(Sections(s).Name & ".bank", _
'                    CDbl(Chunks(Sections(s).Chunks(1)).AddrReal \ &H4000&) _
'                )
'            End If
'        End If
'    Next s
'
'    '----------------------------------------------------------------------------------
'    '[7]: Place paged Sections
'    '----------------------------------------------------------------------------------
'    'TODO: Obviously implement correct rules for Slots...
''    Call RaiseMessage(OZ80_LOG_INFO, "---------¦--------¦------------------------------------------------------")
'
'    For s = 1 To SectionCount
'        'Is this Section assigned to a Slot pattern other than 'SEGA 1 KB'?
'        If (Sections(s).Slot And (SLOT0 Or SLOT1 Or SLOT2)) <> 0 Then
'            'Walk through the Chunks in this Section
'            For c = 1 To Sections(s).ChunkCount
'                'Place this Chunk
'                Call ROMLayoutChunk(Sections(s).Chunks(c))
'                If Error Then GoTo FinishLayout
'            Next c
'            If Sections(s).ChunkCount <> 0 Then
'                'Record the base address & bank of the Section
'                Call AddUpdateLabel(Sections(s).Name, _
'                    CDbl(Chunks(Sections(s).Chunks(1)).AddrSlot) _
'                )
'                Call AddUpdateLabel(Sections(s).Name & ".bank", _
'                    CDbl(Chunks(Sections(s).Chunks(1)).AddrReal \ &H4000&) _
'                )
'            End If
'        End If
'    Next s
'
''    Call RaiseMessage(OZ80_LOG_INFO, LOG_LINEDOUBLE)
'    GoTo FinishLayout
'
'FinishLayout:
'    '----------------------------------------------------------------------------------
'    Call Profiler.EventEnd(EVENT_ROM_LAYOUT)
'    '*** TIMER END ***
'End Sub

''ROMLayoutChunk : Add a Chunk to the ROM, checking for collisions with other Chunks
''======================================================================================
''ChunkRef       | Which Chunk to add to the ROM
''[ChunkAddr]    | Where specifically to put the Chunk, otherwise find a place
''======================================================================================
'Private Sub ROMLayoutChunk( _
'             ByRef ChunkRef As Long, _
'    Optional ByRef ChunkAddr As Long = -1 _
')
'    '*** TIMER BEGIN ***
'    Call Profiler.EventBegin(EVENT_ROM_ADDCHUNK)
'
'    'TODO: If the ChunkSize is zero, discard it? _
'           (We need to handle zero-sized Chunks earlier though -- how do we handle _
'            references to empty Chunks? -- fixed and non-fixed)
'
'    'How big is the Chunk we're handling?
'    Dim ChunkSize As Long: Let ChunkSize = Chunks(ChunkRef).Size
'    Dim ChunkEnd As Long:  Let ChunkEnd = ChunkAddr + (ChunkSize - 1)
'
'    'The Layout is a 'linked-list' where each link contains a reference to the next _
'     link in the chain. This allows us to re-order the items in the chain quickly _
'     without having to shift up/down the entire list. The Layout list needs to always _
'     be in order so that we can find the gaps between Chunks in the ROM Layout
'    Dim i As Long
'    Dim PrevLink As Long, NextLink As Long
'
'    'If there are no layout items yet, start the linked list
'    '----------------------------------------------------------------------------------
'    If LayoutCount = 0 Then
'        'TODO: Check the address is within the ROM space / overflow
'        'Define the first Chunk in the layout _
'         (fall through to `AddLink`)
'
'    '----------------------------------------------------------------------------------
'    'Do we have an address for the Chunk already?
'    '----------------------------------------------------------------------------------
'    ElseIf ChunkAddr <> -1 Then
'        'Search through the layout chain and find the point where our Chunk goes
'        Let i = 1
'        Do
'            'Remember this link, should we reach a successful conclusion and need _
'             to insert a link between this one and the next
'            Let PrevLink = i
'
'            'Check if the current link comes after our Chunk's desired location
'            '..........................................................................
'            If Layout(i).AddrBegin > ChunkAddr Then
'                'Check that our Chunk won't overlap the following one
'                'ERROR: Our Chunk will not fit in the desired location
'                'TODO: What Line/Col number to use?
'                If ChunkEnd >= Layout(i).AddrBegin Then Call RaiseError( _
'                    OZ80_ERROR_LAYOUT, 0, 0, _
'                    "TODO" _
'                ): GoTo Finish
'
'                'TODO: Check our Chunk won't overlap the end of the ROM
'
'                Let NextLink = Layout(i).NextIndex
'                GoTo AddLink
'
'            'If any Chunk is in the same place as our desired location, error
'            '..........................................................................
'            ElseIf Layout(i).AddrBegin = ChunkAddr Then
'                'ERROR: Two Chunks want the same address
'                'TODO: What Line/Col number to use?
'                Call RaiseError( _
'                    OZ80_ERROR_LAYOUT, 0, 0, _
'                    "TODO" _
'                ): GoTo Finish
'
'            'If a Chunk comes before our desired location, _
'             but its size precludes our desired location, then error
'            '..........................................................................
'            ElseIf Layout(i).AddrEnd >= ChunkAddr Then
'                'ERROR: A Chunk is trying to be placed within another
'                'TODO: What Line/Col number to use?
'                Call RaiseError( _
'                    OZ80_ERROR_LAYOUT, 0, 0, _
'                    "TODO" _
'                ): GoTo Finish
'            End If
'
'            'Follow the chain to the next item in the list
'            Let i = Layout(i).NextIndex
'            'If this is the end of the list, do not continue
'            If i = 0 Then Exit Do
'
'        Loop
'
'        'If all links in the chain have been tested then we have come to the _
'         conclusion that there is no Chunk after our desired address, and no Chunk _
'         before that would overlap. We can go ahead and use the address we want
'        '(fall through to `AddLink` below)
'
'    '----------------------------------------------------------------------------------
'    'Find the first available space to put the Chunk:
'    '----------------------------------------------------------------------------------
'    Else
'        'We will begin with attempting to place our Chunk from $67 onwards
'        'TODO: Refine this definition -- System dependent
'        Let ChunkAddr = &H67&
'
'        'Search through the layout chain and find the point where our Chunk goes
'        Let i = 1
'        Do
'            'Remember this link, should we reach a successful conclusion and need _
'             to insert a link between this one and the next
'            Let PrevLink = i
'            Let NextLink = Layout(i).NextIndex
'
'            'If the beginning of our Chunk clips the current Chunk:
'            '..........................................................................
'            If (ChunkAddr >= Layout(i).AddrBegin) _
'            And (ChunkAddr <= Layout(i).AddrEnd) _
'            Then
'                'Push our location forward
'                Let ChunkAddr = Layout(i).AddrEnd + 1
'
'            'If the end of our Chunk clips the current Chunk:
'            '..........................................................................
'            ElseIf (ChunkAddr + ChunkSize >= Layout(i).AddrBegin) _
'               And (ChunkAddr + ChunkSize < Layout(i).AddrEnd) _
'            Then
'                'Push our location forward
'                Let ChunkAddr = Layout(i).AddrEnd + 1
'
'            'If the current Chunk lies entirely within our Chunk:
'            '..........................................................................
'            ElseIf (Layout(i).AddrBegin >= ChunkAddr) _
'               And (Layout(i).AddrEnd < ChunkAddr + ChunkSize) _
'            Then
'                'Push our location forward
'                Let ChunkAddr = Layout(i).AddrEnd + 1
'            End If
'
'            'Follow the chain to the next item in the list
'            Let i = Layout(i).NextIndex
'            'If this is the end of the list, do not continue
'            If i = 0 Then Exit Do
'        Loop
'
'        Let ChunkEnd = ChunkAddr + (ChunkSize - 1)
'    End If
'
'    '----------------------------------------------------------------------------------
'AddLink:
'    'Increase the number of links in the chain. _
'     Note that they won't necessarily link in array order
'    Let LayoutCount = LayoutCount + 1
'    ReDim Preserve Layout(1 To LayoutCount) As oz80Layout
'
'    'Set the Chunk information on the newly crerated link
'    Let Layout(LayoutCount).AddrBegin = ChunkAddr
'    Let Layout(LayoutCount).AddrSize = ChunkSize
'    Let Layout(LayoutCount).AddrEnd = ChunkEnd
'
'    If PrevLink <> 0 Then
'        Let Layout(LayoutCount).PrevIndex = PrevLink
'        Let Layout(PrevLink).NextIndex = LayoutCount
'    End If
'    If NextLink <> 0 Then
'        Let Layout(LayoutCount).NextIndex = NextLink
'        Let Layout(NextLink).PrevIndex = LayoutCount
'    End If
'
'    '----------------------------------------------------------------------------------
'    With Chunks(ChunkRef)
'        Dim BankAddr As Long
'        Let BankAddr = ChunkAddr Mod &H4000&
'
'        'Set the absolute ROM address on the Chunk
'        Let .AddrReal = ChunkAddr
'        'Add the bank property
'        Call AddUpdateLabel(.FullName & ".bank", CDbl(.AddrReal \ &H4000&))
'
'        'Calculate the banked address (based on the Slot pattern)
'        'TODO: This is `SYSTEM`-specific
'        If (.Slot And SLOT1) <> 0 Then
'            Let .AddrSlot = &H4000& + BankAddr
'
'        ElseIf (.Slot And SLOT2) <> 0 Then
'            Let .AddrSlot = &H8000& + BankAddr
'
'        Else
'            Let .AddrSlot = ChunkAddr
'        End If
'
'        'Update the Label database with the new address
'        Call AddUpdateLabel(.FullName, CDbl(.AddrSlot))
'
'        If ChunkSize > 0 Then
'            If (ChunkEnd + 1) > ROMUsed Then Let ROMUsed = ChunkEnd + 1
'            'Is this going to fit in the ROM?
'            If ROMUsed > ROMSize Then
'                'Keep doubling the ROM size until the contents fits
'                Do While ROMUsed > ROMSize: Let ROMSize = ROMSize * 2: Loop
'                'TODO: Check for ROM too big (hard limit -- 32 MB -- and user limit)
'                ReDim Preserve ROM(0 To ROMSize - 1) As Byte
'            End If
'
'            'Copy the Chunk bytes into the ROM
'            Call W32TLB.MemCopy( _
'                OutputPointer:=VarPtr(ROM(ChunkAddr)), _
'                SourcePointer:=VarPtr(.Data(1)), _
'                  LengthBytes:=ChunkSize _
'            )
'
''            Call RaiseMessage(OZ80_LOG_INFO, _
''                "${$}:{$} | {$} | {$}", _
''                HexStr8((ChunkAddr And &HFF0000) \ &H10000), _
''                HexStr16(ChunkAddr And &HFFFF&), _
''                Right$("      " & Format$(ChunkSize, "#,0"), 6), _
''                .FullName _
''            )
'
'            'Walk through the local properties in the Chunk and update their addresses too
'            'TODO: Replace this Dictionary with something much faster
'            Dim Key As Variant
'            For Each Key In .Properties
'                Call AddUpdateLabel( _
'                    .FullName & Key, _
'                    .AddrSlot + .Properties(Key) _
'                )
''                Call RaiseMessage(OZ80_LOG_INFO, _
''                    "${$}:{$} | {$} | {$}", _
''                    HexStr8((.AddrReal + .Properties(Key) And &HFF0000) \ &H10000), _
''                    HexStr16(.AddrReal + .Properties(Key) And &HFFFF&), _
''                    "     .", _
''                    .FullName & Key _
''                )
'            Next
'        End If
'
'    End With
'
'Finish:
'    Call Profiler.EventEnd(EVENT_ROM_ADDCHUNK)
'    '*** TIMER END ***
'End Sub

''ChunkExists : Test if a Chunk is in the array by the Label name
''======================================================================================
''LabelName      | Fully-qualified Label Name for the Chunk, e.g. "::Section:Label"
''---------------¦----------------------------------------------------------------------
''Returns        | True if the Chunk exists, False otherwise
''======================================================================================
'Private Property Get ChunkExists( _
'    ByRef LabelName As String _
') As Boolean
'    'If there are no Chunks, just return False
'    If ChunkCount = 0 Then Exit Property
'    'Go over each Chunk and check the name
'    Dim i As Long
'    For i = 1 To ChunkCount
'        If Chunks(i).LabelName = LabelName Then
'            Let ChunkExists = True
'            Exit Property
'        End If
'    Next i
'End Property

''GetChunk : Retrieves a Chunk in the array by its Label Name
''======================================================================================
'Private Function GetChunk( _
'    ByRef LabelName As String _
') As Long
'    'Go over each Chunk and check the name
'    Dim i As Long
'    For i = 1 To ChunkCount
'        If Chunks(i).LabelName = LabelName Then
'            Let GetChunk = i
'            Exit Function
'        End If
'    Next i
'End Function

''DeferExpression
''======================================================================================
''Kind           | Where the deferred Expression needs to be applied, ie. ROM / CONST ...
''File           | Which file number the Expression is located within
''TokenIndex     | Which Token within the file the Expression begins
''ROMOffset      | For Z80 use, where in the ROM to write the calculated Value
''======================================================================================
'Private Sub DeferExpression( _
'    ByRef Kind As OZ80_DEFERRED, _
'    ByRef File As Long, _
'    ByRef ChunkID As Long, _
'    ByRef TokenIndex As Long, _
'    ByRef ROMOffset As Long, _
'    ByRef From As Long _
')
'    Let DeferredCount = DeferredCount + 1
'    ReDim Preserve Deferred(1 To DeferredCount) As oz80Deferred
'
'    With Deferred(DeferredCount)
'        Let .Kind = Kind
'        Let .File = File
'        Let .TokenIndex = TokenIndex
'        Let .ChunkID = ChunkID
'        Let .ROMOffset = ROMOffset
'        Let .From = From
'    End With
'End Sub

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
    Let LogStr.Text = FormatText

    'Are there any values to format?
    If UBound(FormatValues) < 0 Then
        'No, raise just the text given
        RaiseEvent Message(LogLevel, LogStr)
        Call Log.Log(LogStr.Text)
    Else
        'Format the given log message before sending it up
        RaiseEvent Message(LogLevel, LogStr.Format(FormatValues))
        Call Log.Log(LogStr.Format(FormatValues).Text)
    End If
End Sub

'RiaseError : Provide error details through an event
'======================================================================================
'Number         | Error number to set on the Class and raise via Event
'Message        | A context-specific error message
'---------------¦----------------------------------------------------------------------
'Returns        | The same error number, for simple tail recursion
'======================================================================================
Private Function RaiseError( _
         ByRef Number As OZ80_ERROR, _
         ByRef Line As Long, ByRef Col As Long, _
         ByRef Message As String _
) As OZ80_ERROR
    'Set the error number on this class
    Let Error = Number
    
    'Get the generic title for an error number
    Dim Title As String
    Select Case Number
        Case OZ80_ERROR_DUPLICATE:      Let Title = "Duplicate Error"
        Case OZ80_ERROR_EXPECTED:       Let Title = "Expected Else Error"
        Case OZ80_ERROR_FILE_END:       Let Title = "Unexected End Of File Error"
        Case OZ80_ERROR_FILE_NOTFOUND:  Let Title = "File Not Found Error"
        Case OZ80_ERROR_FILE_READ:      Let Title = "File Read Error"
        Case OZ80_ERROR_INDEFINITE:     Let Title = "Indefinite Value Error"
        Case OZ80_ERROR_OVERFLOW:       Let Title = "Overflow Error"
        Case OZ80_ERROR_UNDEFINED:      Let Title = "Undefined Error"
        Case OZ80_ERROR_INTERNAL:       Let Title = "Internal Error"
    End Select
    
    'If no file has yet been processed, we won't be returning a file name
    If File = 0 Then
        'Pass the error up to the controller, _
         no filename or line/col location are provided
        RaiseEvent Error(vbNullString, Number, Title, Message, 0, 0)
    Else
        'Pass the error up to the controller
        RaiseEvent Error( _
            Files(File).Path, _
            Number, Title, Message, _
            Line, Col _
        )
    End If
    
    'We return the error number, purely for nicer looking tail-recursion
    Let RaiseError = Number
End Function
