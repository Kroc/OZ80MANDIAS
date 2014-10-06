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
Private WithEvents Chunk As oz80_Chunk
Attribute Chunk.VB_VarHelpID = -1

Private ROM() As Byte                   'The output!
Private ROMSize As Long                 'Size of the ROM (32/64/128/256/512/1024 KB)

Private Const BANKSIZE As Long = 16 * 1024

'A list of start & length entries that determine the layout of the Chunks in the ROM
Private ROMMap As Scripting.Dictionary

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

'EVENT <Chunk> Message : Basic progress message from a Chunk _
 ======================================================================================
Private Sub Chunk_Message(LogLevel As OZ80_LOG, Text As String)
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
    
    'TODO: Error handling return
    Let My_Error = Chunk.AddInstruction(Instruction, Param1, Param2, Param3)
    If My_Error Then
        Stop
    End If
    
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
            
            'HL / IX & IY are the same, a special prefix byte determines IX/IY
            ElseIf Token.Kind = TOKEN_Z80_HL Then
                Let GetZ80Parameter.Mask = MASK_MEM_HLIXY
            ElseIf Token.Kind = TOKEN_Z80_IX _
                Or Token.Kind = TOKEN_Z80_IY _
            Then
                Let GetZ80Parameter.Mask = MASK_MEM_HLIXY
                'IX/IY can have an additional expression, e.g. `ld a, (ix+$8)`
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
            Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_MEM_VAL16
            'Is the number within 8-bits? i.e. $00-$FF
            If Value < 256 _
                Then Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_MEM_VAL8
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
        If Token.Kind = TOKEN_Z80_A Then
            Let GetZ80Parameter.Mask = MASK_REG_A
        ElseIf Token.Kind = TOKEN_Z80_B Then Let GetZ80Parameter.Mask = MASK_REG_B
        ElseIf Token.Kind = TOKEN_Z80_C Then Let GetZ80Parameter.Mask = MASK_REG_C
        ElseIf Token.Kind = TOKEN_Z80_D Then Let GetZ80Parameter.Mask = MASK_REG_D
        ElseIf Token.Kind = TOKEN_Z80_E Then Let GetZ80Parameter.Mask = MASK_REG_E
        ElseIf Token.Kind = TOKEN_Z80_H Then Let GetZ80Parameter.Mask = MASK_REG_H
        ElseIf Token.Kind = TOKEN_Z80_L Then Let GetZ80Parameter.Mask = MASK_REG_L
        ElseIf Token.Kind = TOKEN_Z80_I Then Let GetZ80Parameter.Mask = MASK_REG_I
        ElseIf Token.Kind = TOKEN_Z80_R Then Let GetZ80Parameter.Mask = MASK_REG_R
        ElseIf Token.Kind = TOKEN_Z80_AF Then Let GetZ80Parameter.Mask = MASK_REG_AF
        ElseIf Token.Kind = TOKEN_Z80_BC Then Let GetZ80Parameter.Mask = MASK_REG_BC
        ElseIf Token.Kind = TOKEN_Z80_DE Then Let GetZ80Parameter.Mask = MASK_REG_DE
        ElseIf Token.Kind = TOKEN_Z80_HL Then Let GetZ80Parameter.Mask = MASK_REG_HL
        ElseIf Token.Kind = TOKEN_Z80_SP Then Let GetZ80Parameter.Mask = MASK_REG_SP
        ElseIf Token.Kind = TOKEN_Z80_IX Then Let GetZ80Parameter.Mask = MASK_REG_IX
        ElseIf Token.Kind = TOKEN_Z80_IXL Then Let GetZ80Parameter.Mask = MASK_REG_IXL
        ElseIf Token.Kind = TOKEN_Z80_IXH Then Let GetZ80Parameter.Mask = MASK_REG_IXH
        ElseIf Token.Kind = TOKEN_Z80_IY Then Let GetZ80Parameter.Mask = MASK_REG_IY
        ElseIf Token.Kind = TOKEN_Z80_IYL Then Let GetZ80Parameter.Mask = MASK_REG_IYL
        ElseIf Token.Kind = TOKEN_Z80_IYH Then Let GetZ80Parameter.Mask = MASK_REG_IYH
        ElseIf Token.Kind = TOKEN_Z80_NC Then Let GetZ80Parameter.Mask = MASK_FLAGS_CZ
        ElseIf Token.Kind = TOKEN_Z80_Z Then Let GetZ80Parameter.Mask = MASK_FLAGS_CZ
        ElseIf Token.Kind = TOKEN_Z80_NZ Then Let GetZ80Parameter.Mask = MASK_FLAGS_CZ
        ElseIf Token.Kind = TOKEN_Z80_P Then Let GetZ80Parameter.Mask = MASK_FLAGS_MP
        ElseIf Token.Kind = TOKEN_Z80_PE Then Let GetZ80Parameter.Mask = MASK_FLAGS_MP
        ElseIf Token.Kind = TOKEN_Z80_PO Then Let GetZ80Parameter.Mask = MASK_FLAGS_MP
        ElseIf Token.Kind = TOKEN_Z80_M Then Let GetZ80Parameter.Mask = MASK_FLAGS_MP
        End If
        
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
            'Set the mask bits for 8 and 16-bit number ranges: _
             (this is used to check for overflow when assembling opcodes)
            'Any number will class as 16-bit since it's the maximum allowed
            Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_VAL16
            'Is the number within 8-bits? i.e. $00-$FF
            If Value < 256 _
                Then Let GetZ80Parameter.Mask = GetZ80Parameter.Mask Or MASK_VAL8
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
    Let Operand1 = ScopeValue()
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
    Let Operand2 = ScopeValue()
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
Private Function ScopeValue() As Double
    'Is it a named value?
    'TODO: Handle scope expansion for names
    If Token.Kind = TOKEN_LABEL _
    Or Token.Kind = TOKEN_PROPERTY_USE Then '..........................................
        'TODO: get the full label from the property
        
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Not Labels.Exists(Token.Value) Then _
            Let ScopeValue = OZ80_INDEFINITE Else _
            Let ScopeValue = Labels(Token.Value)
        
    'This is an easy one
    ElseIf Token.Kind = TOKEN_NUMBER Then '............................................
        Let ScopeValue = Token.Value
        
    Else '.............................................................................
        'ERROR: Not a valid operand. We won't produce an error message here as the _
         error might be exepected (i.e. an optional parameter). We fold upwards and _
         whomever knows best will deal with the error
        Let ScopeValue = OZ80_INDEFINITE
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
