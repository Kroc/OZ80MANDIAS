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

'Each block of code / data is assembled separately from the others and with no final _
 position in the ROM in mind until all source has been parsed. Once all source is _
 parsed, the banks are chosen for the blocks and the assembly copied into the ROM
Private Blocks As Scripting.Dictionary
'The current block being assembled into
Private Block As oz80_Block

Private ROM() As Byte                   'The output!
Private ROMSize As Long                 'Size of the ROM (32/64/128/256/512/1024 KB)

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not, and in what format. I don't want to spam unwanted `Debug.Print` output
Event Message( _
    ByRef Depth As Long, ByRef LogLevel As OZ80_LOG, ByRef Text As String _
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
    Set Blocks = New Scripting.Dictionary
    Set SectionNames = New Scripting.Dictionary
    Set SectionSlots = New Scripting.Dictionary
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Stopwatch = Nothing
    Set Labels = Nothing
    Set Block = Nothing: Set Blocks = Nothing
    Set Token = Nothing: Set Files = Nothing
    Set SectionNames = Nothing
    Set SectionSlots = Nothing
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
    ByRef Depth As Long, ByRef LogLevel As OZ80_LOG, ByRef Text As String _
)   'It's up to the controller to decide to display (or `Debug.Print` any messages), _
     we 're not going to unnecessarily spam their console
    RaiseEvent Message(Depth, LogLevel, Text)
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
    RaiseEvent Message(0, OZ80_LOG_INFO, _
        "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision)
    
    'We'll measure the time this takes
    Call Stopwatch.Reset
    
    'Initialise the storage for the TokenStream representations of the source files
    Set Files = New Scripting.Dictionary
    
    'Phase 1: First Pass & Assemble _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Assembling " & Chr(34) & FilePath & Chr(34))
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 1")
    
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
    
    'Phase 2: Layout the ROM banks _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 2")
    'TODO
    
    'Phase 3: Second Parse & Assemble (now with ROM Addresses) _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 3")
    'TODO
    
    'Phase 4: Post Processing, e.g. Checksum _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 4")
    'TODO
    
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Assembly Complete")
    RaiseEvent Message(1, OZ80_LOG_INFO, "ROM size: " & (ROMSize \ 1024) & " KB")
    RaiseEvent Message(1, OZ80_LOG_INFO, "Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    
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
'           PROC        :label      [PARAMS     register
'                                               number]
'                                   [RETURN     register
'                                               number]
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
    '`PROC :<name> [SECTION ::<name>] [PARAMS <list>] [RETURN <list>] { [...] }` _
     ----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_KEYWORD_PROC Then
        'Hand off to the function that will process this
        Call ProcessProc
        
    '----------------------------------------------------------------------------------
    '`SECTION ::<sectionName> [BANK <expr>] [SLOT <expr>]` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_SECTION Then
        'Hand off to the function that will process this
        Call ProcessSection
        
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
        Call RaiseError(OZ80_ERROR_ENDOFFILE)
        Exit Sub
    End If
    
    'As long as there is content left to process in the file, keep going
    GoTo Continue
End Sub

'ScopeBlock: Within a code/data block, i.e. `PROC`, `TABLE` _
 ======================================================================================
Private Sub ScopeBlock()
    'Within a code/data block we can have local labels, data statements and Z80 code
    
    'The block might be empty `{}`, in which case we don't want to walk it
    If Token.Kind = TOKEN_BLOCKCLOSE Then GoTo Finish
    
Continue:
    
    'Z80 code?
    If Token.IsInstruction Then
        Call ProcessZ80
        
    End If
    
    '----------------------------------------------------------------------------------
    
    'Move to the next token. If the file ends unexpectedly, the block has not been _
     closed properly (closing brace is expected)
    If Not Token.Forward() Then
        Call RaiseError(OZ80_ERROR_ENDOFFILE)
    
    'When the closing brace is encountered, we'll drop out of the scope
    ElseIf Token.Kind <> TOKEN_BLOCKCLOSE Then
        GoTo Continue
    End If
    
Finish:
    'Move onto the next token ready for the root scope to process. _
     The file could end just after the block, so no error is necessary
    Call Token.Forward
End Sub

'ScopeExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ScopeExpression() As Double
    'TODO: Handle prefixing, e.g. `-1-!-3`
    
    Dim Operand1 As Long                'Left-hand value in a calculation
    Dim Operand2 As Long                'Right-hand value in a calculation
    Dim Operator As OZ80_TOKEN          'Operator to use
    
    'Begin with fetching the first operand
    Let Operand1 = ScopeValue()
    If My_Error Then Stop
    Let ScopeExpression = Operand1
    
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
        Let ScopeExpression = OZ80_INDEFINITE
    
    ElseIf Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let ScopeExpression = Operand1 + Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let ScopeExpression = Operand1 - Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let ScopeExpression = Operand1 * Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        Let ScopeExpression = Operand1 / Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let ScopeExpression = Operand1 ^ Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        Let ScopeExpression = Operand1 Mod Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        Let ScopeExpression = Operand1 And Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        Let ScopeExpression = Operand1 Or Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_XOR Then
        Let ScopeExpression = Operand1 Xor Operand2
        
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
        Let Operand1 = ScopeExpression
        GoTo Continue
    End If
End Function

'ScopeValue : Get a single value as part of an expression _
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

'/// STATEMENT PROCESSORS /////////////////////////////////////////////////////////////
'Somewhat like a scope, these functions handle the main statements, e.g. `PROC`

'ProcessProc : Process a procedure block _
 ======================================================================================
'PROC   :<name>
'       [SECTION    ::<name>]
'       [PARAMS     <list>]
'       [RETURN     <list>]
'{ [...] }
'======================================================================================
Private Sub ProcessProc()
    'ERROR: The file ended before the first parameter
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
    
    'ERROR: No label name for procedure
    If Token.Kind <> TOKEN_LABEL Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_PROC_NAME _
    ): Exit Sub
    
    'Remember the String ID of the label name
    Dim LabelKey As Long
    Let LabelKey = Token.Value
    
    'ERROR: The file cannot end until the block has been opened and closed
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
    
ReadParams:
    
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
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
        
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
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
        
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
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
        
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
        
    End If
    
    '----------------------------------------------------------------------------------
    
    'With the parameters read, the only allowed token is the opening of the block
    If Token.Kind <> TOKEN_BLOCKOPEN Then Call RaiseError( _
        OZ80_ERROR_EXPECTED _
    ): Exit Sub
    
    'The block cannot be left open
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
    
    'Now define the label name; if a section was given, it needs to be prepended
    Dim LabelName As String
    If IsSectionDefined Then _
        Let LabelName = Token.Strings(SectionKey) & Token.Strings(LabelKey) Else _
        Let LabelName = Token.Strings(LabelKey)
    
    'TODO: Is this block already defined?
    
    'Define the current label location as indefinite, it will be placed in the ROM _
     after the first assembly phase (once the block sizes are known)
    Call DefineLabel(LabelName)
    
    'Create the block
    'TODO: bad use of CRC
    Call AddBlock(oz80.CRC.Calculate(LabelName))
    
    'Enter the procedure block to process its innards
    Call ScopeBlock
    
    'TODO: Update `.size`, `.cycles` properties?
End Sub

'ProcessSection : Define a Section _
 ======================================================================================
'SECTION    ::<name>    [SLOT   <list>]
'======================================================================================
Private Sub ProcessSection()
    'ERROR: The file ended before the first parameter
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
    
    'ERROR: Not a section name following the section keyword
    If Token.Kind <> TOKEN_SECTION Then Call RaiseError( _
        OZ80_ERROR_EXPECTED_SECTION_NAME _
    ): Exit Sub
    
    'Keep ahold of the Section identifier
    Dim SectionKey As Long
    Let SectionKey = Token.Value
    
    'Has this section already been defined?
    If SectionNames.Exists(SectionKey) Then Call RaiseError( _
        OZ80_ERROR_DUPLICATE_SECTION _
    ): Exit Sub
    
    'Define the section name - fetch the string from the TokenStream's _
     string database (identified by the CRC value on the token)
    Call SectionNames.Add(SectionKey, Token.Strings(SectionKey))
    Call SectionSlots.Add(SectionKey, 0)
    
    '----------------------------------------------------------------------------------
    
    'If the file ends here it's valid, as the params are optional
    If Not Token.Forward() Then GoTo Finish
    
    'The only parameter is the Slot list
    If Token.Kind = TOKEN_KEYWORD_SLOT Then
        'ERROR: File ending before we have the value
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
        
        'ERROR: If not an expression, it can't be accepted as a list
        If Not Token.IsExpression Then Call RaiseError( _
            OZ80_ERROR_EXPECTED _
        ): Exit Sub
        
        'TODO: Parse list
        
        'The only allowed Slot lists are: 0 / 1 / 2 / 0, 1 / 0, 1, 2 / 1, 2
        
'        'Evaluate the expression
'        Dim Value As Double
'        Let Value = ScopeExpression()
'        'Was there an error evaluating the expression?
'        If My_Error Then Exit Sub
        
'        'You can't use an indefinite value for a slot number
'        If Value = OZ80_INDEFINITE Then Call RaiseError( _
'            OZ80_ERROR_INDEFINITE _
'        ): Exit Sub
        
        'TODO: Is the value within range?
        
'        Let SectionSlots(SectionKey) = Value
    End If

Finish:
    RaiseEvent Message(2, OZ80_LOG_STATUS, _
        "SECTION " & Token.Strings(SectionKey) & _
        " SLOT " & IIf(SectionSlots(SectionKey) = 0, "?", SectionSlots(SectionKey)) _
    )
End Sub

''ProcessVar : Process a variable definition statement _
' ======================================================================================
''VAR #<name> [<expr>] [AS #<name>]
''======================================================================================
'Private Sub ProcessVar()
'    'ERROR: The file ended before the first parameter
'    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
'
'    'ERROR: Not a variable name for `VAR`
'    If Token.Kind <> TOKEN_VARIABLE Then Call RaiseError( _
'        OZ80_ERROR_EXPECTED_VAR_NAME _
'    ): Exit Sub
'
'    'Fetch the key for the variable name in the string database, _
'     (i.e. remember the variable name before moving past the token)
'    Dim StringKey As Long, Value As Double
'    Let StringKey = Token.Value
'
'    'Move forward to the value parameter
'    'NOTE: the value parameter is optional, and the file could end right here -- _
'           if so, it'll skip over retreiving the value
'    If Token.Forward() And Token.IsExpression Then
'        'Calculate the value of the expression
'        Let Value = ScopeExpression()
'        'If the expression is faulty, an error will already have been raised, _
'         we only need to fold upward to terminate the assembling
'        If My_Error Then Exit Sub
'        'ERROR: A value must be known in order to define a variable, i.e. _
'         yet-unknown values, such as label addresses, cannot be used
'        If Value = OZ80_INDEFINITE Then Call RaiseError( _
'            OZ80_ERROR_INDEFINITE _
'        ): Exit Sub
'    Else
'        'There is no given value for the label, re-use the existing value, _
'         otherwise use zero
'        If Variables.Exists(StringKey) Then _
'            Let Value = Variables.Item(StringKey) Else _
'            Let Value = 0
'    End If
'
'    'TODO: `AS #object`
'
'    'Is the variable already defined?
'    If Variables.Exists(StringKey) Then _
'        Let Variables(StringKey) = Value Else _
'        Call Variables.Add(StringKey, Value)
'
'    'Log the variable
'    RaiseEvent Message(2, OZ80_LOG_STATUS, _
'        "VAR " & Token.Strings(StringKey) & " " & _
'        Right$(Space(40) & Value, 40 - Len(Token.Strings(StringKey))) & _
'        " $" & Right$("000" & Hex(Value), 4) _
'    )
'End Sub

'ProcessZ80 : Process a Z80 instruction (and parameters) _
 ======================================================================================
'<z80> [<param>] [<param>]
'======================================================================================
Private Sub ProcessZ80()
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN
    Let Instruction = Token.Kind
    'Skip over the Z80 instruction token now that we know it. _
     If the file ends it's an error because the block has not been closed
    If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Sub
    
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
    End If
    
    'TODO: Error handling return
    Call Block.AddInstruction(Instruction, Param1, Param2, Param3)
    
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'AddBlock : Begin a new block to hold assembled Z80 code / data _
 ======================================================================================
Private Sub AddBlock(ByRef StringID As Long)
    'Create and attach the new block
    Set Block = New oz80_Block
    'Add it to the collection
    Call Blocks.Add(StringID, Block)
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

'GetList : Keeps reading tokens into a list as long as they're valid expressions _
=======================================================================================
Private Function GetList() As Scripting.Dictionary
    '
End Function

'GetZ80Parameter : Fetch the type / value of a Z80 instruction parameter _
 ======================================================================================
Private Function GetZ80Parameter() As oz80Param
    'If the token is a memory reference, such as `(ix+$8)`, _
     then get the associated register and value
    '----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_PARENOPEN Then
        'ERROR: File ends just as the memory expression is opened
        If Not Token.Forward() Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Function
    
        'The first item in the memory expression can be a register
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
                Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Function
            
            GoTo EndScope
            
AddExpr:    'Look for the `+` following the register name
            If Not Token.Forward() _
                Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Function
            
            'If present, skip over it and check for an expression
            If Token.Kind = TOKEN_OPERATOR_ADD Then
                'ERROR: Something must follow the "+"!
                If Not Token.Forward() _
                    Then Call RaiseError(OZ80_ERROR_ENDOFFILE): Exit Function
                
            'If it's just a register in parens, then finish
            ElseIf Token.Kind = TOKEN_PARENCLOSE Then
                Exit Function
            
            Else
                'TODO: ERROR: Something invalid in the expression
                Stop
            End If
        End If
        
        'Is it an expression at all?
        'TODO: ERROR: Not a valid expression
        If Not Token.IsExpression Then Stop
        
        'Memory references with IX & IY will have already set the register, but for _
         direct memory references, e.g. `jp ($1234)`, we declare this as a pure number
        If GetZ80Parameter.Token = TOKEN_NONE _
            Then Let GetZ80Parameter.Token = TOKEN_NUMBER
        
        GoTo GetValue
    
EndScope:
        If Token.Kind = TOKEN_PARENCLOSE Then
            Call Token.Forward
            Exit Function
        End If
        'ERROR: Unexpected token in memory expression
        Stop
    
    'A Register _
     ----------------------------------------------------------------------------------
    ElseIf Token.IsRegister Then
        Let GetZ80Parameter.Token = Token.Kind
        
        'Map the register token to the mask bit; token numbers are sequential _
         whilst parameter masks are bit combinations allowing rapid comparisons of _
         parameters with allowed options
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
        
    'Anything else classes as a calculable expression
    '----------------------------------------------------------------------------------
    ElseIf Token.IsExpression Then
        'We manually build a number token, since we don't want to use the first _
         token of the expression -- it might be a label for instance and we can _
         only send tokens or numbers to be assembled into opcodes
        Let GetZ80Parameter.Token = TOKEN_NUMBER

GetValue:
        'Try and calculate the expression
        Dim Value As Double
        Let Value = ScopeExpression()
        
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
        OZ80_ERROR_FILENOTFOUND _
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
