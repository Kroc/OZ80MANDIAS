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

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

'For timing this process
Private Stopwatch As PerformanceCounter

'Error number!
Private My_Error As OZ80_ERROR

'Parsing: _
 --------------------------------------------------------------------------------------

'Each source code file is broken down into a machine-readable representation _
 ('oz80_TokenStream') which we store here. When parsing one file and we come across _
 an `INCLUDE` statement, another token stream gets added here
Private Files As Scripting.Dictionary

'The current source file being processed, attached such that we can get the _
 message and error events from it
Private WithEvents Token As oz80_TokenStream
Attribute Token.VB_VarHelpID = -1

'The numerical values of variables / labels / properties &c. will be stored here. _
 The key for a name is the CRC32 of the string

'Variables and Labels are kept separately as the first assembling pass determines _
 the final location of the labels in the ROM and the second pass then fills in the _
 expressions that could not be evaluated, but we need to clear out the variable _
 values produced during the first pass so that they don't change the behaviour of _
 the second pass
Private Labels As Scripting.Dictionary

'Sections need to be defined first before they can be used
Private SectionNames As Scripting.Dictionary
Private SectionSlots As Scripting.Dictionary

'Assembling / Output: _
 --------------------------------------------------------------------------------------

'Each chunk of code / data is assembled separately from the others and with no final _
 position in the ROM in mind until all source has been parsed. Once all source is _
 parsed, the banks are chosen for the chunks and the assembly copied into the ROM
Private Chunks As Scripting.Dictionary
'The current chunk being assembled into
Private Chunk As oz80_Chunk
Attribute Chunk.VB_VarHelpID = -1

Private ROM() As Byte                   'The output!
Private ROMSize As Long                 'Size of the ROM (32/64/128/256/512/1024 KB)

Private Const BANKSIZE As Long = 16 * 1024

'A list of start & length entries that determine the layout of the Chunks in the ROM
Private ROMMap As Scripting.Dictionary

'Opcodes: _
 --------------------------------------------------------------------------------------

'Z80 opcodes are made up of various bit patterns that can be broken down into the _
 following layout: (this information from <z80.info/decoding.htm>)

'             .---------------.
' Bit Number: |7|6|5|4|3|2|1|0|
'             '---------------'
' Component:  | X |  Y  |  Z  |
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
Private TokenRegister([_TOKEN_REGISTERS_BEGIN] To [_TOKEN_REGISTERS_END]) As Long

'Look-up tables for converting certain z80 parameters into opcode-components
Private OpcodeP_RegisterPair([_TOKEN_REGISTERS_BEGIN] To [_TOKEN_REGISTERS_END]) As Long
Private OpcodeY_Bit(0 To 7) As Long
Private OpcodeY_Condition([_TOKEN_REGISTERS_BEGIN] To [_TOKEN_REGISTERS_END]) As Long
Private OpcodeY_Register(1 To 8192) As Long
Private OpcodeY_RST(0 To &H38) As Long
Private OpcodeZ_Register(1 To 8192) As Long

'Our routine for comparing parameters is going to be used hundreds of times, _
 so for clarity, we'll use some class-level copies to share the input parameters _
 with the comparison routine
Private Param1Test As oz80Param
Private Param2Test As oz80Param
Private Param3Test As oz80Param

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output
Event Message( _
    ByRef LogLevel As OZ80_LOG, ByRef Text As String _
)

'A class event is used to communicate error details beyond a simple error number
Event Error( _
    ByVal Number As OZ80_ERROR, ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'If too many possible tokens are created, things will screw up
    Debug.Assert OZ80_TOKEN.[_TOKEN_LAST] < 256
    
    Set Stopwatch = New PerformanceCounter
    Set Labels = New Scripting.Dictionary
    Set Chunks = New Scripting.Dictionary
    Set SectionNames = New Scripting.Dictionary
    Set SectionSlots = New Scripting.Dictionary
    Set ROMMap = New Scripting.Dictionary
    
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
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Stopwatch = Nothing
    Set Labels = Nothing
    Set Chunk = Nothing: Set Chunks = Nothing
    Set Token = Nothing: Set Files = Nothing
    Set SectionNames = Nothing
    Set SectionSlots = Nothing
    Set ROMMap = Nothing
    Erase ROM
    
    Erase TokenRegister
    Erase OpcodeP_RegisterPair
    Erase OpcodeY_Bit
    Erase OpcodeY_Condition
    Erase OpcodeY_Register
    Erase OpcodeY_RST
    Erase OpcodeZ_Register
End Sub

'EVENT <Token> Error : Detailed error message from a TokenStream object _
 ======================================================================================
Private Sub Token_Error( _
    ByVal Number As OZ80_ERROR, ByVal Line As Long, ByVal Col As Long _
)   'Pass this up to the controller, this class will shortly end processing. _
     The `RaiseError` function will select the standard error title from the number
    Call RaiseError(Number, Line, Col)
End Sub

'EVENT <Token> Message : Basic progress message from a TokenStream object _
 ======================================================================================
Private Sub Token_Message( _
    ByRef LogLevel As OZ80_LOG, ByRef Text As String _
)   'It's up to the controller to decide to display (or `Debug.Print` any messages), _
     we 're not going to unnecessarily spam their console
    RaiseEvent Message(LogLevel, Text)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'PROPERTY Error : Expose the internal error number _
 ======================================================================================
Public Property Get Error() As OZ80_ERROR: Let Error = My_Error: End Property

'Assemble : Compile a Z80 binary from a given starting source code file _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    'Reset the error number
    Let My_Error = OZ80_ERROR_NONE
    
    'Announce ourselves
    RaiseEvent Message(OZ80_LOG_INFO, _
        "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision)
    
    'We'll measure the time this takes
    Call Stopwatch.Reset
    
    'Initialise the storage for the TokenStream representations of the source files
    Set Files = New Scripting.Dictionary
    
    'Phase 1: First Pass & Assemble _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(OZ80_LOG_ACTION, "Assembling " & Chr(34) & FilePath & Chr(34))
    RaiseEvent Message(OZ80_LOG_ACTION, "Phase 1")
    
    'Start with parsing the given file into a TokenStream; _
     the assembler does not work with the original text directly
    Call ParseFile(FilePath)
    'If there was a problem parsing the file, an error message will already have been _
     raised by an event, we only need to exit
    If My_Error Then GoTo Finish
    
    'Default to 32 KB ROM (the ROM size will double each time it overflows)
    Let ROMSize = 2 ^ 15
    ReDim ROM(0 To ROMSize - 1) As Byte
    
    Call Token.Rewind
    Call ScopeRoot
    If My_Error Then GoTo Finish
    
    'Phase 2: Layout the ROM _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(OZ80_LOG_ACTION, "Phase 2")
    
    Dim Bank As Long
    'Loop through the Chunks of code/data collected
    
'    Dim i As Long
'    For i = 0 To Chunks.Count - 1
'        'Dereference the Chunk from the dictionary object
'        Set Chunk = Chunks(i)
'
'        'Lookup which Section the Chunk is assigned to
'        Dim Section As Long
'        Let Section = Chunk.Section
'
'        'Look up the Slot pattern the Section uses
'        Dim Slot As Long
'        If Section = 0 Then
'            Let Slot = OZ80_SLOT.SLOT0 Or OZ80_SLOT.SLOT1 Or OZ80_SLOT.SLOT2
'        Else
'            Let Slot = SectionSlots(Section)
'        End If
'    Next i
    
    
    'Phase 3: Second Parse & Assemble (now with ROM Addresses) _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(OZ80_LOG_ACTION, "Phase 3")
    'TODO
    
    'Phase 4: Post Processing, e.g. Checksum _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(OZ80_LOG_ACTION, "Phase 4")
    'TODO
    
    RaiseEvent Message(OZ80_LOG_ACTION, "Assembly Complete")
    RaiseEvent Message(OZ80_LOG_INFO, "ROM size: " & (ROMSize \ 1024) & " KB")
    RaiseEvent Message(OZ80_LOG_INFO, "Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    
Finish:
    Let Assemble = My_Error
    Call Class_Terminate
End Function

'/// SCOPE HANDLERS ///////////////////////////////////////////////////////////////////

'This is a structured map of the OZ80 syntax. _
 The first column on the left is the name of a scope. For each scope, the columns to _
 the right represent the heirarchy of allowable items in that scope. Whatever is in _
 the third column is only allowed to follow what is in the second column. Items below _
 one another within a column represent more than one item being allowed within a scope.
'Another scope name within a column represents a nested scope. Items within brackets _
 show optional items

'NOTES: *   You can take `{ASCII}` to mean any ASCII code 0-127 _
        *   `{operator}` can be any one of `+ - * / ^ \ & | ! << >>`

'Scope:     1           2           3           4           5           6 _
 --------------------------------------------------------------------------------------
'<text>     "           {ASCII}     "
'
'<expr>     <number>    [{operator} <expr>]
'           #var        [{operator} <expr>]
'           :label      [{operator} <expr>]
'           $.ram       [{operator} <expr>]
'           {           <expr>      }
'
'
'<list>     <text>      [,]         [<list>]
'           <expr>      [,]         [<list>]
'
'......................................................................................
'
'<root>     INCLUDE     <text>
'
'           OBJECT      #object     {           <object>
'                                   }
'
'           PROC        :label      [PARAMS     <list>]
'                                   [RETURN     <list>]
'                                   [INTERRUPT  <expr>]
'                                   {           <code>
'                                   }
'
'           SECTION     ::section   [BANKS      <expr>]
'                                   [BANK       <expr>]
'                                   [SLOT       <expr>]
'
'           STRUCT      :label      AS          #object     <struct>
'
'           TABLE       :label      <code>
'
'           VAR         #variable   [<list>]
'
'<code>     .label
'           DATA        <list>
'           BYTE        <list>
'           WORD        <list>
'           TRIP        <list>
'           LONG        <list>
'
'           <asm>

'ScopeRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Sub ScopeRoot()
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
    'Used for evaluating expressions in parameters
    Dim Value As Long
    'For remembering item names (labels, variables &c.)
    Dim StringKey As Long
    
Continue:
    If My_Error Then Exit Sub
    'As we loop, other functions will move the Token pointer forward (such as when _
     evaluating expressions). Check if we've hit the end of the file
    If Token.EOF Then Exit Sub
    
    'ERROR: Not a keyword on root scope
    If Not Token.IsKeyword Then Call RaiseError(OZ80_ERROR_EXPECTED_ROOT): Exit Sub
        
    '----------------------------------------------------------------------------------
    '`PROC :<label> { [...] }` _
     ----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_KEYWORD_PROC Then
        'Hand off to the function that will process this
        Call ProcessProc
        
    '----------------------------------------------------------------------------------
    '`SECTION ::<section>` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_SECTION Then
        'Hand off to the function that will process this
        Call ProcessSection
        
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
        Call RaiseError(OZ80_ERROR_FILE_END)
        Exit Sub
    End If
    
    'As long as there is content left to process in the file, keep going
    GoTo Continue
End Sub

'ScopeChunk: Within a code/data chunk, i.e. `PROC`, `TABLE` _
 ======================================================================================
Private Sub ScopeChunk()
    'Within a code/data chunk we can have local labels, data statements and Z80 code
    
Continue:
    'Z80 code?
    If Token.IsInstruction Then
        Call ProcessZ80
        If My_Error Then Exit Sub
        GoTo Continue
    End If
    
    'The chunk might be empty `{}`, in which case we don't want to walk it
    If Token.Kind = TOKEN_CHUNKCLOSE Then GoTo Finish
    
    '----------------------------------------------------------------------------------
    
    'Move to the next token. If the file ends unexpectedly, the chunk has not been _
     closed properly (closing brace is expected)
    If Not Token.Forward() Then
        Call RaiseError(OZ80_ERROR_FILE_END)
    
    'When the closing brace is encountered, we'll drop out of the scope
    ElseIf Token.Kind <> TOKEN_CHUNKCLOSE Then
        GoTo Continue
    End If
    
Finish:
    'Move onto the next token ready for the root scope to process. _
     The file could end just after the chunk, so no error is necessary
    Call Token.Forward
End Sub

'/// STATEMENT PROCESSORS /////////////////////////////////////////////////////////////
'Somewhat like a scope, these functions handle the main statements, e.g. `PROC`

'ProcessProc : Process a procedure chunk _
 ======================================================================================
'PROC   :<label>
'       [SECTION    ::<section>]
'       [PARAMS     <list>]
'       [RETURN     <list>]
'       [INTERRUPT  <expr>]
'{ [...] }
'======================================================================================
Private Sub ProcessProc()
    'ERROR: The file ended before the first parameter
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'ERROR: No label name for procedure
    If Token.Kind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_PROC_NAME _
    ): Exit Sub
    
    'Remember the String ID of the Label name, _
     we can't create the Label until the parameters are read
    Dim LabelKey As Long
    Let LabelKey = Token.Value
    
    'ERROR: The file cannot end until the chunk has been opened and closed
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'The `INTERRUPT` parameter is optional, but can be zero, so use -1 for 'undefined'
    Dim Interrupt As Double
    Let Interrupt = -1
    
ReadParams:
    '----------------------------------------------------------------------------------
    '`SECTION` Parameter : Set which Section the procedure belongs to _
     ----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_KEYWORD_SECTION Then
        Dim SectionKey As Long
        Dim IsSectionDefined As Boolean
    
        'ERROR: The section parameter has already been specified
        If IsSectionDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_SECTION _
        ): Exit Sub
        
        'ERROR: File ended before the section name?
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'ERROR: Not a section name following the section keyword
        If Token.Kind <> TOKEN_SECTION Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_SECTION_NAME _
        ): Exit Sub
        
        'ERROR: Section has not been defined
        If Not SectionNames.Exists(Token.Value) Then Call RaiseError( _
            OZ80_ERROR_INVALID_SECTION _
        ): Exit Sub
        
        'Record the section name
        Let SectionKey = Token.Value
        Let IsSectionDefined = True
        
        'Pass over the Section name
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
     
    '----------------------------------------------------------------------------------
    '`PARAMS` Parameter : Document which Z80 parameters the procedure uses _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_PARAMS Then
        'The `PARAMS` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM names
         Dim IsParamsDefined As Boolean
         
        'ERROR: The params parameter has already been specified
        If IsParamsDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_PARAMS _
        ): Exit Sub
         
SkipParam:
        'ERROR: File cannot end in the midst of the list
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'If a Keyword, the list has come to the end
        If Token.IsKeyword() Then
            Let IsParamsDefined = True
            GoTo ReadParams
        End If
        
        'ERROR: Only registers and RAM names are allowed at the moment
        If (Not Token.IsRegister) And _
           (Token.Kind <> TOKEN_RAM) Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_PROC_PARAMS _
        ): Exit Sub
        
        GoTo SkipParam
        
    '----------------------------------------------------------------------------------
    '`RETURN` Parameter : Document what the procedure sets / returns _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_RETURN Then
        'The `RETURN` parameter is not used by the assembler at this time, _
         neither is it a standard list as it consists of Registers and/or RAM names
         Dim IsReturnDefined As Boolean
         
        'ERROR: The return parameter has already been specified
        If IsReturnDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_RETURN _
        ): Exit Sub
         
SkipReturn:
        'ERROR: File cannot end in the midst of the list
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'If a Keyword, the list has come to the end
        If Token.IsKeyword() Then
            Let IsReturnDefined = True
            GoTo ReadParams
        End If
        
        'ERROR: Only registers and RAM names are allowed at the moment
        If (Not Token.IsRegister) _
        And (Token.Kind <> TOKEN_RAM) Then Call RaiseError( _
            OZ80_ERROR_EXPECTED_PROC_RETURN _
        ): Exit Sub
        
        GoTo SkipReturn
        
    '----------------------------------------------------------------------------------
    '`INTERRUPT` Parameter : Define a fixed-location Interrupt Procedure _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_INTERRUPT Then
        'The `INTERRUPT` parameter fixes a Procedure to a specific location in the _
         ROM intended for hardware interrupts. Whilst you shouldn't specify a Section _
         for an Interrupt Procedure, you can do so but the Section must use Slot 0 or _
         Slot 1 KB. This is not validated here! Sections Slots are validated after _
         the first assembly pass when we try and layout the Chunks in the ROM
        Dim IsInterruptDefined As Boolean
        
        'ERROR: The interrupt parameter has already been specified
        If IsInterruptDefined Then Call RaiseError( _
            OZ80_ERROR_DUPLICATE_PROC_INTERRUPT _
        ): Exit Sub
        
        'ERROR: File cannot end before the parameter Expression!
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'Is this an Expression at all?
        If Not Token.IsExpression Then Call RaiseError(OZ80_ERROR_EXPECTED): Exit Sub
        
        'Calculate the Expression:
        Let Interrupt = ParseExpression()
        'If there was an error in the expression, the error event _
         will already have been raised, we only need to fold upwards
        If My_Error Then Exit Sub
        
        'An indefinite value cannot be used for an Interrupt address, we will only be _
         able to define the indefinite values (i.e. Label locations in ROM) once _
         the whole code is parsed and we use the Sections to layout the ROM
        If Interrupt = OZ80_INDEFINITE Then Call RaiseError(OZ80_ERROR_INDEFINITE)
        
        'Validate the Value, only certain Interrupt addresses are allowed
        If Not ((Interrupt = &H0) Or (Interrupt = &H8) Or (Interrupt = &H10) _
            Or (Interrupt = &H18) Or (Interrupt = &H20) Or (Interrupt = &H28) _
            Or (Interrupt = &H30) Or (Interrupt = &H38) Or (Interrupt = &H66)) _
        Then Call RaiseError(OZ80_ERROR_INVALID_INTERRUPT): Exit Sub
        
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the chunk
    If Token.Kind <> TOKEN_CHUNKOPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): Exit Sub
    
    'The chunk cannot be left open
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'Now define the label name; if a section was given, it needs to be prepended
    Dim LabelName As String
    If IsSectionDefined Then _
        Let LabelName = Token.Strings(SectionKey) & Token.Strings(LabelKey) Else _
        Let LabelName = Token.Strings(LabelKey)
    
    'TODO: Is this chunk already defined?
    
    'Define the current label location as indefinite, it will be placed in the ROM _
     after the first assembly phase (once the chunk sizes are known)
    Call DefineLabel(LabelName)
    
    'Create the chunk
    'TODO: messy use of CRC
    Call AddChunk(oz80.CRC.Calculate(LabelName))
    
    'Apply the Interrupt parameter to the new Chunk
    Let Chunk.Interrupt = Interrupt
    
    Dim LogMsg As String
    Let LogMsg = "PROC " & LabelName
    If Interrupt >= 0 _
        Then Let LogMsg = LogMsg & " INTERRUPT $" & oz80.HexStr8(CLng(Interrupt))
    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    
    'Enter the procedure chunk to process its innards
    Call ScopeChunk
End Sub

'ProcessSection : Define a Section _
 ======================================================================================
'SECTION    ::<section>    [SLOT   <list>]
'======================================================================================
Private Sub ProcessSection()
    'ERROR: The file ended before the first parameter
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'ERROR: Not a section name following the section keyword
    If Token.Kind <> TOKEN_SECTION Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_SECTION_NAME _
    ): Exit Sub
    
    'Keep ahold of the Section identifier
    Dim SectionKey As Long
    Let SectionKey = Token.Value
    
    'Has this Section already been defined?
    If SectionNames.Exists(SectionKey) Then Call RaiseError( _
        OZ80_ERROR_DUPLICATE_SECTION _
    ): Exit Sub
    
    'Define the Section name - fetch the string from the TokenStream's _
     string database (identified by the CRC value on the token)
    Call SectionNames.Add(SectionKey, Token.Strings(SectionKey))
    Call SectionSlots.Add(SectionKey, 0)
    
    '----------------------------------------------------------------------------------
    
    'If the file ends here it's valid, as the params are optional
    If Not Token.Forward() Then GoTo Finish
    
    'The only allowed Slot lists are: 0 / 1 / 2 / 0, 1 / 0, 1, 2 / 1, 2. _
     These are stored as a bit-pattern allowing packing into one value, _
     which when zero infers the default "let the assembler choose for me"
    Dim Slot As OZ80_SLOT
    
    'The only parameter is the Slot list
    If Token.Kind = TOKEN_KEYWORD_SLOT Then
        'ERROR: File ending before we have the value
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
        
        'ERROR: There needs to be at least one item in the list
        If Not Token.IsExpression Then Call RaiseError( _
            OZ80_ERROR_EXPECTED _
        ): Exit Sub
        
        'Get the list of Slots assigned to the Section
        Dim List As Scripting.Dictionary
        Set List = GetListAsDictionary()
        
        'Can't be more than three -- "0, 1, 2"
        If List.Count > 3 Then Call RaiseError(OZ80_ERROR_INVALID_SLOT): Exit Sub
        
        'An indefinite value cannot be used for the Slot
        If List.Item(0) = OZ80_INDEFINITE Then GoTo ErrIndefinite
        
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
            If List.Item(1) = OZ80_INDEFINITE Then GoTo ErrIndefinite
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
    Let SectionSlots(SectionKey) = Slot
    RaiseEvent Message(OZ80_LOG_STATUS, _
        "SECTION " & Token.Strings(SectionKey) & _
        " SLOT " & IIf( _
            Slot = 0, "?", _
                IIf(Slot And OZ80_SLOT.SLOT0, "0 ", "") & _
                IIf(Slot And OZ80_SLOT.SLOT1, "1 ", "") & _
                IIf(Slot And OZ80_SLOT.SLOT2, "2 ", "") _
            ) _
    )
    Exit Sub

    '----------------------------------------------------------------------------------
ErrInvalid:
    'Only a specific Slot pattern can be used
    Call RaiseError(OZ80_ERROR_INVALID_SLOT)
    Exit Sub

ErrIndefinite:
    'An indefinite value cannot be used for a Slot number, we will only be _
     able to define the indefinite values (i.e. Label locations in ROM) once _
     the whole code is parsed and we use the Sections to layout the ROM
    Call RaiseError(OZ80_ERROR_INDEFINITE)
End Sub

'ProcessZ80 : Process a Z80 instruction (and parameters) _
 ======================================================================================
'<z80> [<param>] [<param>] [<param>]
'======================================================================================
Private Sub ProcessZ80()
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN
    Let Instruction = Token.Kind
    'Skip over the Z80 instruction token now that we know it. _
     If the file ends it's an error because the chunk has not been closed
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Sub
    
    'TODO: Handle POP / PUSH
    
    'Up to three parameters are supported to allow for the undocumented instructions, _
     e.g. `SET 0 (IX+$8) A` which sets bit 0 of the address and copies the result to A
    Dim Param1 As oz80Param, Param2 As oz80Param, Param3 As oz80Param
    
    'Is there a first parameter?
    If Token.IsParameter Then
        'Parameter expressions might return an unknown value _
         (such as when a label, which has not yet been addressed, is used)
        
        'Get the first parameter
        Let Param1 = GetZ80Parameter()
        'If there was an error, the error number will already have been set, _
         we only need to fold upwards to end the assembling
        If My_Error Then Exit Sub
        
        'Check if a second parameter exists
        If Token.IsParameter Then
            'Fetch the second parameter
            Let Param2 = GetZ80Parameter()
            If My_Error Then Exit Sub
            
            'Check if a third parameter exists
            If Token.IsParameter Then
                'Fetch the third parameter (very uncommon)
                Let Param2 = GetZ80Parameter()
                If My_Error Then Exit Sub
            End If
        End If
    Else
        'When an instruction has parameters, the parameter-reading always leaves _
         the token stream pointing at the next Token to process. If there are no _
         parameters, we do not want to leave the token stream pointing at the z80 _
         instruction otherwise it'll get parsed again!
        'Call Token.Forward
    End If
    
    'OPCODE MAP: _
     Here's a complete list of the Z80 instruction set and the opcodes used for each _
     addressing type.
     
    'http://www.smspower.org/Development/InstructionSet
    
    'NOTE:
    '   +rr     BC = +00
    '           DE = +10
    '           HL = +20 (Also IX, IY)
    '           SP = +30 (Also AF)
    '   rr      IX = DD
    '           IY = FD
    '           HL = rr and ?? is omitted
    '   +f      nz = +00
    '           z  = +08
    '           nc = +10
    '           c  = +18
    '           po = +20
    '           pe = +28
    '           p  = +30
    '           m  = +38
    '   +b      bit * 8
    '   +im     0  = +00
    '           1  = +10
    '           2  = +18
    
    '   ixyh    ixh|iyh
    '   ixyl    ixl|iyl
    '   ixyhl   ixh|iyh|ixl|iyl
        
    'Z80    Param 1                 Param 2                 Opcode         Cycles _
     ----------------------------------------------------------------------------------
    'ADC    a                       a|b|c|d|e|h|ixyh|l|ixyl rr 88+r             4
    '                               (hl|ix+$8|iy+$8)        rr 8E ??            7
    '                               $8                      CE ??               7
    '       hl                      bc|de|hl|sp             ED 4A+rr            15
                                
    'ADD    a                       a|b|c|d|e|h|ixyh|l|ixyl rr 80+r             4
    '                               (hl|ix+$8|iy+$8)        rr 86 ??            7
    '                               $8                         C6 ??            7
    '       hl                      bc|de|hl|sp                09+rr            11
    '       ix                      bc|de|ix|sp             rr 09+rr            11
    '       iy                      bc|de|iy|sp             rr 09+rr            11
    
    'AND    a|b|c|d|e|h|ixyh|l|ixyl                         rr A0+r
    '       (hl|ix+$8|iy+$8)                                rr A6 ??
    '       $8                                                 E6 ??
    
    'BIT    0-7                     a|b|c|d|e|h|l              CB 40+r+b
    '                               (hl|ix+$8|iy+$8)        rr CB ?? 46+b
    '       0-7     (ix+$8|iy+$8)   a|b|c|d|e|h|l           rr CB ?? 40+r+b
    
    'CALL   $16                                                CD ?? ??
    '       c|nc|m|p|z|nz|pe|po     $16                        C4+f
    
    'CCF                                                       3F
    
    'CP     a|b|c|d|e|h|ixyh|l|ixyl                         rr B8+r
    '       (hl|ix+$8|iy+$8)                                rr BE ??
    '       $8                                                 FE ??
    
    'CPD                                                    ED A9
    'CPDR                                                   ED B9
    'CPI                                                    ED A1
    'CPIR                                                   ED B1
    'CPL                                                       2F
    'DAA                                                       27
    
    'DEC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 05+r*8
    '       (hl|ix+$8|iy+$8)                                rr 35 ??
    '       bc|de|hl|sp                                        0B+rr
    '       ix|iy                                           rr 2B
    
    'DI                                                        F3
    
    'DJNZ   $8                                                 10 ??
    
    'EI                                                        FB
    
    'EX     (sp)                    hl|ix|iy                rr E3
    '       af                      af                         08
    '       de                      hl                         EB
    
    'EXX                                                       D9
    'HALT                                                      76
    
    'IM     0-2                                                46+im
    
    'IN     a                       ($8)                       DB ??
    '       (c)                                             ED 70
    '       a|b|c|d|e|h|l           (c)                     ED 40+r
    
    'INC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 04+r*8
    '       bc|de|hl|sp|ix|iy                               rr 03+rr
    '       (hl|ix+$8|iy+$8)                                rr 34 ??
    
    'IND                                                    ED AA
    'INDR                                                   ED BA
    'INI                                                    ED A2
    'INIR                                                   ED B2
    
    'JP     $16                                                C3 ?? ??
    '       (hl|ix|iy)                                      rr E9
    '       c|nc|m|p|z|nz|pe|po     $16                        C2+f ?? ??
    
    'JR     $8                                                 18 ??
    '       c|nc|z|nz               $8                         20+f
    
    'LD     i                       a                       ED 47
    '       r                       a                       ED 4F
    '       a                       i                       ED 57
    '                               r                       ED 5F
    '                               a|b|c|d|e|h|ixyh|l|ixyl rr 78+r
    '                               (hl|ix+$8|iy+$8)        rr 7E ??
    '                               $8                         3E ??
    '                               (bc|de)                    0A+rr
    '                               ($16)                      3A ?? ??
    '       b                       a|b|c|d|e|h|ixyh|l|ixyl rr 40+r
    '                               (hl|ix+$8|iy+$8)        rr 46 ??
    '                               $8                         06 ??
    '       c                       a|b|c|d|e|h|ixyh|l|ixyl rr 48+r
    '                               (hl|ix+$8|iy+$8)        rr 4E ??
    '                               $8                         0E ??
    '       d                       a|b|c|d|e|h|ixyh|l|ixyl rr 50+r
    '                               (hl|ix+$8|iy+$8)        rr 56 ??
    '                               $8                         16 ??
    '       e                       a|b|c|d|e|h|ixyh|l|ixyl    58+r
    '                               (hl|ix+$8|iy+$8)        rr 5E ??
    '                               $8                         1E ??
    '       h                       a|b|c|d|e|h|l              60+r
    '                               (hl|ix+$8|iy+$8)        rr 66 ??
    '                               $8                         26 ??
    '       ixh                     a|b|c|d|e|ixh|ixl       rr 60+r
    '                               $8                      rr 26 ??
    '       iyh                     a|b|c|d|e|iyh|iyl       rr 60+r
    '                               $8                      rr 26 ??
    '       l                       a|b|c|d|e|h|l              68+r
    '                               (hl|ix+$8|iy+$8)        rr 6E ??
    '                               $8                         2E ??
    '       ixl                     a|b|c|d|e|ixh|ixl       rr 68+r
    '                               $8                      rr 2E ??
    '       iyl                     a|b|c|d|e|iyh|iyl       rr 68+r
    '                               $8                      rr 2E ??
    '       bc                      ($16)                   ED 4B ?? ??
    '                               $16                        01+rr ?? ??
    '       de                      ($16)                   ED 5B ?? ??
    '                               $16                        01+rr ?? ??
    '       hl                      ($16)                      2A ?? ??
    '                               $16                        01+rr ?? ??
    '       sp                      ($16)                   ED 7B ?? ??
    '                               hl|ix|iy                rr F9
    '                               $16                        01+rr ?? ??
    '       ix|iy                   ($16)                   rr 2A ?? ??
    '                               $16                     rr 21 ?? ??
    '       (hl|ix+$8|iy+$8)        a|b|c|d|e|h|l           rr 70+r
    '                               $8                      rr 36 ??
    '       (bc)                    a                          02+rr
    '       (de)                    a                          02+rr
    '       ($16)                   a                          32 ?? ??
    '                               bc                      ED 43 ?? ??
    '                               de                      ED 53 ?? ??
    '                               hl|ix|iy                rr 22 ?? ??
    '                               sp                      ED 73 ?? ??
    
    'LDD                                                    ED A8
    'LDDR                                                   ED B8
    'LDI                                                    ED A0
    'LDIR                                                   ED B0
    'NEG                                                    ED 44
    'NOP                                                       00
    
    'OR     a|b|c|d|e|h|ixyh|l|ixyl                         rr B0+r
    '       (hl|ix+$8|iy+$8)                                rr B6 ??
    '       $8                                                 F6 ??
    
    'OUT    ($8)                    a                          D3 ??
    '       (C)                     0                       ED 71
    '                               a|b|c|d|e|h|l           ED 41+r*8
    
    'OUTD                                                   ED AB
    'OTDR                                                   ED BB
    'OUTI                                                   ED A3
    'OTIR                                                   ED B3
    
    'POP    af|bc|de|hl|ix|iy                               rr C1+rr
    
    'PUSH   af|bc|de|hl|ix|iy                               rr C5+rr
    
    'RES    0-7                     a|b|c|d|e|h|l              CB 80+r+b
    '                               (hl|ix+$8|iy+$8)        rr CB ?? 86+b
    '       0-7     (ix+$8|iy+$8)   a|b|c|d|e|h|l           rr CB ?? 80+r+b
    
    'RET                                                       C9
    '       c|nc|m|p|z|nz|pe|po                                C0+f
    
    'RETI                                                   ED 4D
    'RETN                                                   ED 45
    'RLA                                                       17
    
    'RL     a|b|c|d|e|h|l                                      CB 10+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 16
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 10+r
    
    'RLCA                                                      07
    
    'RLC    a|b|c|d|e|h|l                                      CB 00+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 06
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 00+r
    
    'RLD                                                    ED 6F
    'RRA                                                       1F
    
    'RR     a|b|c|d|e|h|l                                      CB 18+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 1E
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 18+r
    
    'RRCA                                                      0F
    
    'RRC    a|b|c|d|e|h|l                                      CB 08+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 0E
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 08+r
    
    'RRD                                                    ED 67
    
    'RST    0|08|10|18|20|28|30|38                             C7+num
    
    'SBC    a|b|c|d|e|h|ixyh|l|ixyl                         rr 98+r
    '       (hl|ix+$8|iy+S8)                                rr 9E ??
    '       a                       $8                         DE ??
    '       hl                      bc|de|hl|sp             ED+rr
    
    'SCF                                                       37
    
    'SET    0-7                     a|b|c|d|e|h|l              CB C0+b+r
    '                               (hl|ix+$8|iy+$8)        rr CB ?? C6+b+r
    '       0-7     (ix+$8|iy+$8)   a|b|c|d|e|h|l           rr CB ?? C0+r+b
    
    'SLA    a|b|c|d|e|h|l                                      CB 20+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 26
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 20+r
    
    'SLL    a|b|c|d|e|h|l                                      CB 30+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 36
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 30+r
    
    'SRA    a|b|c|d|e|h|l                                      CB 28+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 2E
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB 28+r
    
    'SRL    a|b|c|d|e|h|l                                      CB 38+r
    '       (hl|ix+$8|iy+$8)                                rr CB ?? 3E
    '       (ix+$8|iy+$8)           a|b|c|d|e|h|l           rr CB ?? 38+r
    
    'SUB    a|b|c|d|e|h|ixyh|l|ixyl                         rr 90+r
    '       (hl|ix+$8|iy+$8)                                rr 96 ??
    '       $8                                                 D6 ??
    
    'XOR    a|b|c|d|e|h|ixyh|l|ixyl                         rr A8+r
    '       (hl|ix+$8|iy+$8)                                rr AE ??
    '       $8                                                 EE ??
    
    '----------------------------------------------------------------------------------
    Dim Prefix As Long, Opcode As Long, Offset As Long
    Let Offset = -1
    
    'TODO: All this string concatenation is *SLOW* _
           Need to find a way to do logging fast
    Dim LogMsg As String
    Let LogMsg = "$" & oz80.HexStr16(Chunk.Size) & ": "
    
    'Since the numerical value that accompanies some instructions could be in _
     param 1 or 2, we need to record which to use at the end. This excludes _
     the extra value that goes with IX/IY which is handled separately, _
     e.g. when using `LD (IX+$8), $8`
    Dim ValueParam As oz80Param
    
    'Use of IX / IY requires a prefix byte ($DD/$FD) before the opcode. _
     We need to record which parameter is the IX/IY reference, along with its value
    Dim ParamIXY As oz80Param
    'NOTE: Parameter three is only used for some undocumented instructions and _
     will not be IX/IY (though IX/IY can be in previous parameters)
    If (Param1.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Param1
    ElseIf (Param2.Mask And MASK_ANY_IXY) <> 0 Then
        Let ParamIXY = Param2
    End If
    'If there is any IX/Y parameter, add the opcode prefix
    If ParamIXY.Mask <> 0 Then
        'Is it IX or IY?
        If (ParamIXY.Mask And MASK_ANY_IX) <> 0 _
            Then Let Prefix = &HDD _
            Else Let Prefix = &HFD
        'Does it include an offset byte? i.e. "(IX+$8)"
        If (ParamIXY.Mask And MASK_MEM_IXY) <> 0 Then Let Offset = ParamIXY.Value
    End If
    
    'Our routine for comparing parameters is going to be used hundreds of times, _
     so for clarity, we'll use some class-level references to share the input _
     parameters with the comparison routine
    Let Param1Test = Param1
    Let Param2Test = Param2
    Let Param3Test = Param3
    
    'See, BASIC does have a good use!
    On (Instruction - [_TOKEN_INSTRUCTIONS_BEGIN]) GoTo _
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
    '`ADC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y1 Or OpcodeZ_Register(Param2.Mask), Param2.Value)
    '`ADC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit( _
            X3 Or Y1 Or Z6, Param2.Value)
    '`ADC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED( _
            X1 Or Q1 Or Z2 Or OpcodeP_RegisterPair(Param2.Token))
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iADD:
    'Add: _
     ..................................................................................
    '`ADD HL, BC|DE|HL|SP`
    If TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode( _
            OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1)
    '`ADD IX, BC|DE|IX|SP`
    ElseIf TestParams(MASK_REG_IX, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IX Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IX"), _
         we still need to validate the params to ensure HL & IX are not mixed
        Let LogMsg = LogMsg & Chunk.AddOpcodePrefix(Prefix, _
            OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1)
    '`ADD IY, BC|DE|IY|SP`
    ElseIf TestParams(MASK_REG_IY, MASK_REG_BC Or MASK_REG_DE Or MASK_REG_IY Or MASK_REG_SP) Then
        'Even though this is the same opcode as above (the prefix denotes "IY"), _
         we still need to validate the params to ensure HL & IY are not mixed
        Let LogMsg = LogMsg & Chunk.AddOpcodePrefix(Prefix, _
            OpcodeP_RegisterPair(Param2.Token) Or Q1 Or Z1)
    '`ADD A, A|B|C|D|E|H|IXH|IYH|L|IYH|IYL|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or OpcodeZ_Register(Param2.Mask), Param2.Value)
    '`ADD A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Z6, Param2.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iAND:
    'Bitwise AND: _
     ..................................................................................
    '`AND A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y4 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`AND $8`
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y4 Or Z6, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iBIT:
    'Test a bit: _
     ..................................................................................
    '`BIT 0-7, A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            X1 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask), _
            Param2.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iCALL:
    'Call Procedure: _
     ..................................................................................
    '`CALL NZ|Z|NC|C|PO|PE|P|M, $16` - conditional call
    If TestParams(MASK_FLAGS, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            X3 Or OpcodeY_Condition(Param1.Token) Or Z4, Param2.Value)
    '`CALL $16` - absolute call
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            X3 Or Q1 Or Z5, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iCCF:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y7 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCP:
    '..................................................................................
    '`CP A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y7 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`CP $8`
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y7 Or Z6, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iCPD:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y5 Or Z1): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPDR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y7 Or Z1): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iCPI:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y4 Or Z1): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPIR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y6 Or Z1): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iCPL:
    'Compliment: Flip the bits of the Accumulator _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y5 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDAA:
    'Decimal Adjust Accumulator: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y4 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iDEC:
    'Decrement: _
     ..................................................................................
    '`DEC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode( _
            OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z3)
    '`DEC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            OpcodeY_Register(Param1.Mask) Or Z5, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iDI:
    'Disable Interrupts: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Y6 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iDJNZ:
    'Decrement, Jump if Not Zero: _
     ..................................................................................
    If TestParams(MASK_VAL) Then Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(Y2, Param1.Value): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
     
iEI:
    'Enable Interrupts: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Y7 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iEX:
    'Exchange: _
     ..................................................................................
    '`EX AF AF'`
    If TestParams(MASK_REG_AF, MASK_REG_AF) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(Y1)
    '`EX (SP), HL|IX|IY`
    ElseIf TestParams(MASK_MEM_SP, MASK_REGS_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Y4 Or Z3)
    '`EX DE, HL`
    ElseIf TestParams(MASK_REG_DE, MASK_REG_HL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Y5 Or Z3)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iEXX:
    'Exchange All: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Q1 Or P1 Or Z1): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iHALT:
    'Halt: Stop the processor and wait for interrupts _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(X1 Or Y6 Or Z6): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iIM:
    'Interrupt Mode: _
     ..................................................................................
    '`IM 0|1|2`
    If TestParams(MASK_VAL) Then
        'An indefinite value is assumed to be zero. _
         The assembler will ensure that there are no indefinite values remaining _
         on the second parsing of the source code
        If Param1.Value = OZ80_INDEFINITE Then
            Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Z6)
        ElseIf Param1.Value = 0 Then
            Let LogMsg = LogMsg & Chunk.AddOpcode(X1 Or Z6)
        ElseIf Param1.Value = 1 Then
            Let LogMsg = LogMsg & Chunk.AddOpcode(X1 Or Y2 Or Z6)
        ElseIf Param1.Value = 2 Then
            Let LogMsg = LogMsg & Chunk.AddOpcode(X1 Or Y3 Or Z6)
        Else
            'ERROR: Wrong parameters for Z80 instruction
            GoTo InvalidParams
        End If
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iIN:
    'Read from port: _
     ..................................................................................
    '`IN A, ($8)`
    If TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y3 Or Z3, Param2.Value)
    '`IN A|B|C|D|E|H|L, (C)`
    ElseIf TestParams(MASK_REGS_ABCDEHL, MASK_MEM_BC) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or OpcodeY_Register(Param1.Mask))
    '`IN (C)`
    ElseIf TestParams(MASK_MEM_BC) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y6)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iINC:
    'Increment: _
     ..................................................................................
    '`INC BC|DE|HL|SP|IX|IY`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(OpcodeP_RegisterPair(Param1.Token) Or Z3)
    '`INC A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            OpcodeY_Register(Param1.Mask) Or Z4, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iIND:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y5 Or Z2): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINDR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y7 Or Z2): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iINI:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y4 Or Z2): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iINIR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y6 Or Z2): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iJP:
    'Jump: _
     ..................................................................................
    '`JP (HL|IX|IY)` or `JP HL|IX|IY` - indirect absolute jump
    If TestParams(MASK_MEM_HLIXY Or MASK_REGS_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or P2 Or Q1 Or Z1)
    '`JP NZ|Z|NC|C|PO|PE|P|M, $16` - conditional absolute jump
    ElseIf TestParams(MASK_FLAGS, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            X3 Or OpcodeY_Condition(Param1.Token) Or Z2, Param2.Value)
    '`JP $16` - absolute jump
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            X3 Or Z3, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iJR:
    'Jump Relative: _
     ..................................................................................
    '`JR $8` - relative immediate jump
    If TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(Y3, Param1.Value)
    '`JR NZ|Z|NC|C, $8` - conditional relative immediate jump
    ElseIf TestParams(MASK_FLAGS_CZ, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit( _
            OpcodeY_Condition(Param1.Token) + Y4, Param2.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iLD:
    'The Mother-Load: _
     ..................................................................................
    '`LD BC|DE|HL|SP|IX|IY, $16`
    If TestParams(MASK_REGS_BC_DE_HL_SP_IXY, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            OpcodeP_RegisterPair(Param1.Token) Or Z1, Param2.Value)
    '`LD (BC), A`
    ElseIf TestParams(MASK_MEM_BC, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(Z2)
    '`LD (DE), A`
    ElseIf TestParams(MASK_MEM_DE, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(P1 Or Z2)
    '`LD A, (BC)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_BC) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(Q1 Or Z2)
    '`LD A, (DE)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_DE) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(P1 Or Q1 Or Z2)
    '`LD ($16), HL|IX|IY`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            P2 Or Z2, Param1.Value)
    '`LD ($16), A`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            P3 Or Z2, Param1.Value)
    '`LD HL|IX|IY, ($16)`
    ElseIf TestParams(MASK_REGS_HL_IXY, MASK_MEM_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            P2 Or Q1 Or Z2, Param2.Value)
    '`LD A, ($16)`
    ElseIf TestParams(MASK_REG_A, MASK_MEM_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value16Bit(Prefix, _
            P3 Or Q1 Or Z2, Param2.Value)
    '`LD A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8), $8`
    ElseIf TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit_Offset(Prefix, _
            OpcodeY_Register(Param1.Mask) Or Z6, Param2.Value, Param1.Value)
    'You can't combine IX and IY use in the same opcode, there's only one prefix, _
     so these two following checks ensure that IX & IY are not used together
    '`LD A|B|C|D|E|H|IXH|L|IXL(HL|IX+$8), A|B|C|D|E|H|IXH|L|IXL`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL _
    ) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask), _
            Param1.Value)
    '`LD A|B|C|D|E|H|IYH|L|IYL(HL|IY+$8), A|B|C|D|E|H|IYH|L|IYL`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL _
    ) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask), _
            Param1.Value)
    'You can't have `LD (IX+$8), (IX+$8)!
    '`LD A|B|C|D|E|H|IXH|L|IXL, A|B|C|D|E|H|IXH|L|IXL|(HL|IX+$8)`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IXHL Or MASK_MEM_HL Or MASK_MEM_IX _
    ) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask), _
            Param2.Value)
    '`LD A|B|C|D|E|H|IYH|L|IYL, A|B|C|D|E|H|IYH|L|IYL|(HL|IY+$8)`
    ElseIf TestParams( _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL, _
        MASK_REGS_ABCDEHL Or MASK_REGS_IYHL Or MASK_MEM_HL Or MASK_MEM_IY _
    ) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X1 Or OpcodeY_Register(Param1.Mask) Or OpcodeZ_Register(Param2.Mask), _
            Param2.Value)
    '`LD SP, HL|IX|IY`
    ElseIf TestParams(MASK_REG_SP, MASK_REGS_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or P3 Or Q1 Or Z1)
    '`LD ($16), BC|DE|SP`
    ElseIf TestParams(MASK_MEM_VAL, MASK_REGS_BC_DE_SP) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED_Value16Bit( _
            X1 Or OpcodeP_RegisterPair(Param2.Token) Or Z3, _
            Param1.Value)
    '`LD BC|DE|SP, ($16)`
    ElseIf TestParams(MASK_REGS_BC_DE_SP, MASK_MEM_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED_Value16Bit( _
            X1 Or OpcodeP_RegisterPair(Param1.Token) Or Q1 Or Z3, _
            Param2.Value)
    '`LD I, A`
    ElseIf TestParams(MASK_REG_I, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Z7)
    '`LD R, A`
    ElseIf TestParams(MASK_REG_R, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y1 Or Z7)
    '`LD A, I`
    ElseIf TestParams(MASK_REG_A, MASK_REG_I) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y2 Or Z7)
    '`LD A, R`
    ElseIf TestParams(MASK_REG_A, MASK_REG_R) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y3 Or Z7)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iLDD:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y5): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iLDDR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iLDI:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y4): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
        
iLDIR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y6): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iNEG:
    'Negate the Accumulator: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Z4): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iNOP:
    'No Operation: _
     ..................................................................................
    '`NOP` - Opcode is 0
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(0): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOR:
    'Bitwise OR: _
     ..................................................................................
    '`OR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X2 Or Y6 Or OpcodeZ_Register(Param1.Mask))
    '`OR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y6 Or Z6, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iOUT:
    'Write to port: _
     ..................................................................................
    '`OUT ($8), A`
    If TestParams(MASK_MEM_VAL, MASK_REG_A) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y2 Or Z3, Param1.Value)
    '`OUT (C), A|B|C|D|E|H|L`
    ElseIf TestParams(MASK_MEM_BC, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED( _
            X1 Or OpcodeY_Register(Param2.Mask) Or Z1)
    '`OUT (C), 0`
    ElseIf TestParams(MASK_MEM_BC, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y6 Or Z1)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iOUTD:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y5 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTDR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y7 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iOUTI:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y4 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iOTIR:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X2 Or Y6 Or Z3): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iPOP:
    'Fetch from stack: _
     ..................................................................................
    '`POP AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode( _
            X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z1)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iPUSH:
    'Push a register on to the stack: _
     ..................................................................................
    '`PUSH AF|BC|DE|HL|IX|IY`
    If TestParams(MASK_REGS_AF_BC_DE_HL_IXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode( _
            X3 Or OpcodeP_RegisterPair(Param1.Token) Or Z5)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iRES:
    'Reset a bit: _
     ..................................................................................
    '`RES 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            X2 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask), _
            Param2.Value)
    '`RES 0-7, (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RES memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            X2 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param3.Mask), _
            Param2.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iRET:
    'Return: _
     ..................................................................................
    '`RET NZ|Z|NC|C|PO|PE|P|M` - conditional return
    If TestParams(MASK_FLAGS) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or OpcodeY_Condition(Param1.Token))
    '`RET`
    ElseIf TestParams() Then
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or Q1 Or Z1)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iRETI:
    'Return from Interrupt: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y1 Or Z5): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRETN:
    'Return from NMI: _
     ..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Z5): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLA:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y2 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRL:
    '..................................................................................
    '`RL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y2 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`RL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y2 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iRLCA:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRLC:
    '..................................................................................
    '`RLC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`RLC (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RLC memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iRLD:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y5 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRRA:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y3 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iRR:
    '..................................................................................
    '`RR A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y3 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`RR (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RR memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y3 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iRRCA:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y1 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRRC:
    '..................................................................................
    '`RRC A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y1 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`RRC (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented RRC memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y1 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iRRD:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or Y4 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams

iRST:
    '"Reset" - Fire specific interrupt vector: _
     ..................................................................................
    '`RST $00|$08|$10|$18|$20|$28|$30|$38`
    If TestParams(MASK_VAL) Then
        'TODO: Need to validate the number is in range at some point in the chain
        Let LogMsg = LogMsg & Chunk.AddOpcode(X3 Or OpcodeY_RST(Param1.Value) Or Z7)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iSBC:
    '..................................................................................
    '`SBC A, A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REG_A, MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y3 Or OpcodeZ_Register(Param2.Mask), Param2.Value)
    '`SBC A, $8`
    ElseIf TestParams(MASK_REG_A, MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y3 Or Z6, Param2.Value)
    '`SBC HL, BC|DE|HL|SP`
    ElseIf TestParams(MASK_REG_HL, MASK_REGS_BC_DE_HL_SP) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeED(X1 Or OpcodeP_RegisterPair(Param2.Token) Or Z2)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iSCF:
    '..................................................................................
    If TestParams() Then Let LogMsg = LogMsg & Chunk.AddOpcode(Y6 Or Z7): GoTo LogEnd
    'ERROR: Wrong parameters for Z80 instruction
    GoTo InvalidParams
    
iSET:
    'Set a bit: _
     ..................................................................................
    '`SET 0-7 A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_VAL, MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            X3 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param2.Mask), _
            Param2.Value)
    '`SET 0-7, (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SET memory & register
    ElseIf TestParams(MASK_VAL, MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            X3 Or OpcodeY_Bit(Param1.Value) Or OpcodeZ_Register(Param3.Mask), _
            Param2.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iSLA:
    '..................................................................................
    '`SLA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y4 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`SLA (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SLA memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y4 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iSRA:
    '..................................................................................
    '`SRA A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y5 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`SRA (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SRA memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y5 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iSLL:
    '..................................................................................
    '`SLL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y6 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`SLL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SLL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y6 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iSRL:
    '..................................................................................
    '`SRL A|B|C|D|E|H|L|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y7 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`SRL (IX+$8|IY+$8), A|B|C|D|E|H|L` - undocumented SRL memory and copy to register
    ElseIf TestParams(MASK_MEM_IXY, MASK_REGS_ABCDEHL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcodeCB(Prefix, _
            Y7 Or OpcodeZ_Register(Param2.Mask), Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

iSUB:
    'Subtract: _
     ..................................................................................
    '`SUB A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y2 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`SUB $8`
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y2 Or Z6, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd
    
iXOR:
    'Bitwise XOR: _
     ..................................................................................
    '`XOR A|B|C|D|E|H|IXH|IYH|L|IXL|IYL|(HL|IX+$8|IY+$8)`
    If TestParams(MASK_REGS_ABCDEIXYHL_MEM_HLIXY) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Offset(Prefix, _
            X2 Or Y5 Or OpcodeZ_Register(Param1.Mask), Param1.Value)
    '`XOR $8`
    ElseIf TestParams(MASK_VAL) Then
        Let LogMsg = LogMsg & Chunk.AddOpcode_Value8Bit(X3 Or Y5 Or Z6, Param1.Value)
    'ERROR: Wrong parameters for Z80 instruction
    Else: GoTo InvalidParams: End If
    GoTo LogEnd

LogEnd:
    Let LogMsg = LogMsg & " | " & oz80.TokenName(Instruction)
    If Param1.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param1)
    If Param2.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param2)
    If Param3.Mask <> 0 Then Let LogMsg = LogMsg & vbTab & oz80.ParamToString(Param3)
    RaiseEvent Message(OZ80_LOG_STATUS, LogMsg)
    Exit Sub
    
InvalidParams:
    '----------------------------------------------------------------------------------
    'ERROR: Wrong parameters for Z80 instruction
    Call RaiseError(OZ80_ERROR_INVALID_Z80PARAMS)
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'AddChunk : Begin a new Chunk to hold assembled Z80 code / data _
 ======================================================================================
Private Sub AddChunk( _
    ByRef StringID As Long _
)
    'Create and attach the new chunk
    Set Chunk = New oz80_Chunk
    'Add it to the collection
    Call Chunks.Add(StringID, Chunk)
End Sub

'AddUpdateLabel: Blindly adds to or updates the label names array _
 ======================================================================================
Private Sub AddUpdateLabel( _
             ByRef LabelName As String, _
    Optional ByRef Value As Double = OZ80_INDEFINITE _
)   'Hash the name to get the index/key
    Dim CRC As Long
    Let CRC = oz80.CRC.Calculate(LabelName)
    
    'Add a new label or update an existing label value?
    If Not Labels.Exists(CRC) Then _
        Call Labels.Add(CRC, Value) Else
        Let Labels.Item(CRC) = Value
End Sub

'DefineLabel : Add a label name / value definition (and generate its properties) _
 ======================================================================================
Private Sub DefineLabel( _
             ByRef LabelName As String, _
    Optional ByRef Value As Double = OZ80_INDEFINITE _
)
    'Firstly define the base label name
    Call AddUpdateLabel(LabelName, Value)
    
    'TODO: .size, .bank, .slot properties?
    
    'If the value is indefinite, the properties will be also
    If Value = OZ80_INDEFINITE Then
        'Concatenate the properties onto the label
        Call AddUpdateLabel(LabelName & ".lo")
        Call AddUpdateLabel(LabelName & ".hi")
    
    Else
        'Extract the high / low bytes from the address
        'TODO: Warning if fractions?
        Call AddUpdateLabel(LabelName & ".lo", CLng(Value) And &HFF)
        Call AddUpdateLabel(LabelName & ".hi", (CLng(Value) And &HFF00&) \ &H100)
    End If
End Sub

'GetListAsDictionary : Reads a whole list into a Dictionary _
 ======================================================================================
Private Function GetListAsDictionary() As Scripting.Dictionary
    'Instantiate the return object
    Set GetListAsDictionary = New Scripting.Dictionary
    
    'As long as the list continues, keep fetching Values
    Do While Token.IsExpression
        'Parse the Expression; if it's indefinite, _
         leave that up to the caller to handle
        Call GetListAsDictionary.Add(GetListAsDictionary.Count, ParseExpression())
        
        'If there was an error in the expression, the error event _
         will already have been raised, we only need to fold upwards
        If My_Error Then Exit Function
    Loop
End Function

'GetListItem : Read in an Expression from a list _
 ======================================================================================
Private Function GetListItem( _
             ByRef Value As Double, _
    Optional ByRef Length As Long = 0 _
) As Boolean
    'If the current token is not an expression, the list has ended; return False
    If Not Token.IsExpression Then Exit Function
    
    'Parse the Expression; if it's indefinite, _
     leave that up to the caller to handle
    Let Value = ParseExpression()
    
    'If there was an error in the expression, the error event _
     will already have been raised, we only need to fold upwards
    If My_Error Then Exit Function
    
    'Increases the number of items read from the list thus far
    Let Length = Length + 1
    
    'Return True as the list has not ended yet
    Let GetListItem = True
End Function

'GetZ80Parameter : Fetch the type / value of a Z80 instruction parameter _
 ======================================================================================
Private Function GetZ80Parameter() As oz80Param
    Dim Value As Double
    
    'If the parameter doesn't contain a value, default to -1. This is so that when _
     building the opcode we can immediately tell when to include the offset byte
    Let GetZ80Parameter.Value = -1
    
    'A memory reference, such as `(ix+$8)` _
     ----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_PARENOPEN Then
        'ERROR: File ends just as the memory expression is opened
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Function
    
        'The first item in the memory expression can be a register _
         ..............................................................................
        If Token.IsRegister Then
            'Copy the register type to the parameter being built
            Let GetZ80Parameter.Token = Token.Kind
            
            'The `IN` & `OUT` instructions accept only the C register, _
             though in reality this is register pair BC
            If Token.Kind = TOKEN_Z80_C Then
                Let GetZ80Parameter.Mask = MASK_MEM_BC
            
            '`LD` accepts registers BC & DE
            ElseIf Token.Kind = TOKEN_Z80_BC Then
                Let GetZ80Parameter.Mask = MASK_MEM_BC
            ElseIf Token.Kind = TOKEN_Z80_DE Then
                Let GetZ80Parameter.Mask = MASK_MEM_DE
            
            'HL / IX & IY
            ElseIf Token.Kind = TOKEN_Z80_HL Then
                Let GetZ80Parameter.Mask = MASK_MEM_HL
            'IX/IY can have an additional expression, e.g. `ld a, (ix+$8)`
            ElseIf Token.Kind = TOKEN_Z80_IX Then
                Let GetZ80Parameter.Mask = MASK_MEM_IX
                GoTo AddExpr
            ElseIf Token.Kind = TOKEN_Z80_IY Then
                Let GetZ80Parameter.Mask = MASK_MEM_IY
                GoTo AddExpr
                
            ElseIf Token.Kind = TOKEN_Z80_SP Then
                Let GetZ80Parameter.Mask = MASK_MEM_SP
            Else
                'ERROR: register not allowed here
                Stop
            End If
            
            'ERROR: File ends before the memory expression is closed
            If Not Token.Forward() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Function
            
            GoTo EndScope
            
AddExpr:    'Look for the `+` following the register name
            If Not Token.Forward() _
                Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Function
            
            'If present, skip over it and check for an expression
            If Token.Kind = TOKEN_OPERATOR_ADD Then
                'ERROR: Something must follow the "+"!
                If Not Token.Forward() _
                    Then Call RaiseError(OZ80_ERROR_FILE_END): Exit Function
                
                'Try and calculate the expression
                Let Value = ParseExpression()
                
                'If there was an error in the expression, the error event _
                 will already have been raised, we only need to fold upwards
                If My_Error Then Exit Function
                
                'TODO: Handle indefinite value
                'TODO: Error if number is over 16-bit
                'TODO: Issue warning for rounding
                
                Let GetZ80Parameter.Value = Value
                
                If Token.Kind = TOKEN_PARENCLOSE Then
                    Call Token.Forward
                    Exit Function
                End If
                'ERROR: Unexpected token in memory expression
                Stop
            End If
        
        '..............................................................................
        Else
            'Is it an expression at all?
            'TODO: ERROR: Not a valid expression
            If Not Token.IsExpression Then Stop
        
            'Besides registers, memory addresses are allowed
            Let GetZ80Parameter.Token = TOKEN_NUMBER
            
            'Try and calculate the expression
            Let Value = ParseExpression()
            
            'If there was an error in the expression, the error event _
             will already have been raised, we only need to fold upwards
            If My_Error Then Exit Function
            
            'TODO: Handle indefinite value
            'TODO: Error if number is over 16-bit
            'TODO: Issue warning for rounding
            
            Let GetZ80Parameter.Value = Value
            
            'Any number will class as 16-bit since it's the maximum allowed
            Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_MEM_VAL
        End If
    
EndScope:
        '..............................................................................
        If Token.Kind = TOKEN_PARENCLOSE Then
            Call Token.Forward
            Exit Function
        End If
        'ERROR: Unexpected token in memory expression
        Stop
    
    'A Register _
     ----------------------------------------------------------------------------------
    ElseIf Token.IsRegister Or Token.IsFlag Then
        'Copy the register type to the parameter being built
        Let GetZ80Parameter.Token = Token.Kind
        
        'Map the register token to the mask bit; token numbers are sequential _
         whilst parameter masks are bit combinations allowing rapid comparisons _
         of parameters with allowed options
        Let GetZ80Parameter.Mask = TokenRegister(Token.Kind)
        
        'Now move beyond the register token
        Call Token.Forward
    
    'Anything else classes as a calculable expression _
     ----------------------------------------------------------------------------------
    ElseIf Token.IsExpression Then
        'We manually build a number token, since we don't want to use the first _
         token of the expression -- it might be a label for instance and we can _
         only send registers or numbers to be assembled into opcodes
        Let GetZ80Parameter.Token = TOKEN_NUMBER

GetValue:
        'Try and calculate the expression
        Let Value = ParseExpression()
        
        'If there was an error in the expression, the error event _
         will already have been raised, we only need to fold upwards
        If My_Error Then Exit Function
        
        'TODO: Handle indefinite value
        'TODO: Error if number is over 16-bit
        'TODO: Issue warning for rounding
        
        Let GetZ80Parameter.Value = Value
        
        'IX/IY values are always 8-bit, so don't use the MASK_VAL bits
        If GetZ80Parameter.Token = TOKEN_NUMBER Then
            Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_VAL
        End If
    
    '----------------------------------------------------------------------------------
    Else
        'ERROR: Not a Z80 parameter
        'TODO: parameter specific error number
        Call RaiseError(OZ80_ERROR_EXPRESSION_Z80)
    End If
End Function

'ParseExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ParseExpression() As Double
    'TODO: Handle prefixing, e.g. `-1-!-3`
    'TODO: Handle functions like LO, HI, NOT8/16
    
    Dim Operand1 As Long                'Left-hand value in a calculation
    Dim Operand2 As Long                'Right-hand value in a calculation
    Dim Operator As OZ80_TOKEN          'Operator to use
    
    'Begin with fetching the first operand
    Let Operand1 = ParseValue()
    If My_Error Then Stop
    Let ParseExpression = Operand1
    
    'The file could legitimately end with an expression, so don't error out if the _
     token stream comes to an end since the expression is valid. The parent scopes _
     will handle an unexpected end error
    If Token.EOF Then Exit Function
    
    'If next token is not an operator assume the expression has finished, _
      the parent will handle whatever the next token is used for
    If Not Token.IsOperator Then Exit Function
    
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = Token.Kind
    
    'Get the second operand
    'TODO: Error - unexpected end
    If Not Token.Forward() Then Stop
    Let Operand2 = ParseValue()
    'Cannot have an operator hanging there without an operand following
    If My_Error Then
        'TODO: Error - invalid use of operator
        Stop
    End If
    
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If Operand1 = OZ80_INDEFINITE _
    Or Operand2 = OZ80_INDEFINITE Then
        'Mark expression result as indefinite and skip applying the operator
        Let ParseExpression = OZ80_INDEFINITE
    
    ElseIf Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let ParseExpression = Operand1 + Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let ParseExpression = Operand1 - Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let ParseExpression = Operand1 * Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        Let ParseExpression = Operand1 / Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let ParseExpression = Operand1 ^ Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        Let ParseExpression = Operand1 Mod Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        Let ParseExpression = Operand1 And Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        Let ParseExpression = Operand1 Or Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_XOR Then
        Let ParseExpression = Operand1 Xor Operand2
        
    Else
        Stop
    End If
    
'    'The file could legitimately end with an expression, so don't error out if the _
'     token stream comes to an end since the expression is valid. The parent scopes _
'     will handle an unexpected end error
'    If NextToken() = False Then Exit Function
    
    'Does the expression continue?
    If Token.IsOperator Then
        'Set the current result as the first operand and go around again
        Let Operand1 = ParseExpression
        GoTo Continue
    End If
End Function

'ParseFile : Parse a source code file into a TokenStream object _
 ======================================================================================
Private Sub ParseFile( _
    ByRef FilePath As String _
)
    'Since we have a "scrrun.dll" reference for Dictionary object, this also gives _
     us the `FileSystemObject` to simplify a lot of mucky folder / file checking
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    '[1] Normalise the file path: _
     ----------------------------------------------------------------------------------
    'The file path handed to us might be just a file name and/or contain a relative _
     path (i.e. "..\"). We need to normalise this given path into an absolute path _
     (e.g. "C:\Folder\File.ext")
    
    Dim TestPath As String
    
    'If the given path is already absolute, it does not need normalising. _
     NOTE: Upper/lower case will be corrected and could throw a false negative
    If LCase$(FSO.GetAbsolutePathName(FilePath)) = LCase$(FilePath) Then
        'We'll skip ahead to checking the file exists
        Let TestPath = FilePath
    Else
        'Path is relative! If no other file has been parsed yet _
         we will need to use the current directory as a base
        If Token Is Nothing Then
            Let TestPath = FSO.BuildPath(CurDir(), FilePath)
        Else
            'When another file has been parsed before, _
             we work relatively from that file's directory
            Let TestPath = FSO.GetFile(Token.SourceFile).Path
            'TODO: Error?
            Let TestPath = FSO.BuildPath(TestPath, FilePath)
        End If
    End If
    
    '[2] Check file exists: _
     ----------------------------------------------------------------------------------
    'ERROR: File Not Found
    If Not FSO.FileExists(TestPath) Then Call RaiseError( _
        OZ80_ERROR_FILE_NOTFOUND _
    ): GoTo Finish
    
    '[3] Parse the file into a TokenStream: _
     ----------------------------------------------------------------------------------
    'We'll use the CRC of the filepath string as its key/ID
    Dim CRC As Long
    Let CRC = oz80.CRC.Calculate(TestPath)
    
    'If the file has already been parsed, re-use it
    If Files.Exists(CRC) Then
        'Reference the existing TokenStream object
        Set Token = Files.Item(CRC)
    Else
        'Create a tokeniser object to hold the machine representation of the text file; _
         the assembler doesn't work with the original text directly
        Set Token = New oz80_TokenStream
        'Explode the source code file into tokens
        Call Token.Tokenise(TestPath)
        'If the file failed to parse, an error code / message will already have been _
         raised via an event, we only need to exit
        If My_Error Then GoTo Finish
        'Add the TokenStream to the files collection
        Call Files.Add(CRC, Token)
    End If
    
Finish:
    Set FSO = Nothing
    
    'If there were any error, an error event and message would already have been _
     raised, we only need to fold upwards and the parent will handle terminating the _
     assembling process early
End Sub

'ParseValue : Get a single Value as part of an Expression _
 ======================================================================================
Private Function ParseValue() As Double
    'Is it a named value?
    'TODO: Handle scope expansion for names
    If Token.Kind = TOKEN_LABEL _
    Or Token.Kind = TOKEN_PROPERTY_USE Then '..........................................
        'TODO: get the full label from the property
        
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Not Labels.Exists(Token.Value) Then _
            Let ParseValue = OZ80_INDEFINITE Else _
            Let ParseValue = Labels(Token.Value)
        
    'This is an easy one
    ElseIf Token.Kind = TOKEN_NUMBER Then '............................................
        Let ParseValue = Token.Value
        
    Else '.............................................................................
        'ERROR: Not a valid operand. We won't produce an error message here as the _
         error might be exepected (i.e. an optional parameter). We fold upwards and _
         whomever knows best will deal with the error
        Let ParseValue = OZ80_INDEFINITE
        Let My_Error = OZ80_ERROR_EXPRESSION
    End If
    
    'Move beyond the current token now it's been processed
    Call Token.Forward
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Sub RaiseError( _
             ByRef Number As OZ80_ERROR, _
    Optional ByRef Line As Long = -1, Optional ByRef Col As Long = -1 _
)
    'Set the error number on this class
    Let My_Error = Number
    
    If Line < 0 Then Let Line = Token.Line
    If Col < 0 Then Let Col = Token.Col
    
    'Fetch the detailed error text
    Dim Title As String, Description As String
    Call oz80.GetOZ80Error(Number, Title, Description)
    
    'Pass the error up to the controller
    RaiseEvent Error(Number, Title, Description, Line, Col)
End Sub

'TestParams : Check if parameters are of the right type desired _
 ======================================================================================
Private Function TestParams( _
    Optional ByRef Test1 As OZ80_MASK, _
    Optional ByRef Test2 As OZ80_MASK, _
    Optional ByRef Test3 As OZ80_MASK _
) As Boolean
    'Why call the function with no parameters? It's a consistent way to check when _
     there *should* be no parameters for an instruction!
    
    'Are we testing a first parameter?
    If (Test1 = 0) Then
        'If a parameter exists, but there shouldn't be, return False
        If (Param1Test.Mask <> 0) Then Exit Function
        'If there is no first parameter, pass True
        If (Param1Test.Mask = 0) Then GoTo Finish
    End If
    
    '1st Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param1Test.Mask And Test1) = 0 Then Exit Function
    
    'Are we testing a second parameter?
    If (Test2 = 0) Then
        'If there shouldn't be a second parameter, return False
        If (Param2Test.Mask <> 0) Then Exit Function
        'If there is no second parameter, the test has passed
        If (Param2Test.Mask = 0) Then GoTo Finish
    End If
    
    '2nd Parameter Test: _
     ----------------------------------------------------------------------------------
        
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param2Test.Mask And Test2) = 0 Then Exit Function
    
    'If there shouldn't be a third parameter, return False
    If (Test3 = 0) And (Param3Test.Mask <> 0) Then Exit Function
    'If there is no third parameter, the test has passed
    If (Test3 = 0) And (Param3Test.Mask = 0) Then GoTo Finish
    
    '3rd Parameter Test: _
     ----------------------------------------------------------------------------------
    
    'Does the parameter appear in the test? E.g. Param = A, Test = A/B/C/D/E/H/L
    If (Param3Test.Mask And Test3) = 0 Then Exit Function
    
Finish:
    Let TestParams = True
End Function
