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

'The CRC function of this class is from "Calculating CRC32 With VB" by Steve McMahon _
 <www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp> which was _
 derived from from the VB.NET CRC32 algorithm by Paul Caton

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

'For timing this process
Private Stopwatch As PerformanceCounter

'Parsing: _
 --------------------------------------------------------------------------------------

'Each source code file is broken down into a machine-readable representation which _
 we store here. When parsing one file and we come across an `INCLUDE` statement, _
 another token stream gets added here
Private Files As Scripting.Dictionary

'This structure is used when parsing parameters of Z80 instructions
Private Type oz80Param
    Kind As OZ80_PARAM                  'The kind of parameter, e.g. register, value
    Value As Long                       'The numerical value for the parameter
    ValueIndefinite As Boolean          'If the parameter has unknown values (labels)
End Type

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

Private Variables As Scripting.Dictionary
Private Labels As Scripting.Dictionary

'Sections need to be defined first before they can be used, _
 i.e. `SECTION ::stuff BANK 0 SLOT 0`
Private SectionNames As Scripting.Dictionary
Private SectionBanks As Scripting.Dictionary
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
    Set Stopwatch = New PerformanceCounter
    Set Variables = New Scripting.Dictionary
    Set Labels = New Scripting.Dictionary
    Set Blocks = New Scripting.Dictionary
    Set SectionNames = New Scripting.Dictionary
    Set SectionBanks = New Scripting.Dictionary
    Set SectionSlots = New Scripting.Dictionary
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Stopwatch = Nothing
    Set Variables = Nothing
    Set Labels = Nothing
    Set Block = Nothing: Set Blocks = Nothing
    Set Token = Nothing: Set Files = Nothing
    Set SectionNames = Nothing
    Set SectionBanks = Nothing
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

'Assemble : Compile a Z80 binary from a given starting source code file _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    'Announce ourselves
    RaiseEvent Message(0, OZ80_LOG_INFO, _
        "OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision)
    
    'We'll measure the time this takes
    Call Stopwatch.Reset
    
    'Initialise the storage for the TokenStream representations of the source files
    Set Files = New Scripting.Dictionary
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Assembling " & Chr(34) & FilePath & Chr(34))
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 1")
    
    'Start with parsing the given file into a TokenStream; the assembler does not _
     work with the original text directly
    Let Assemble = ParseFile(FilePath)
    'If there was a problem parsing the file, an error message will already have been _
     raised by an event, we only need to exit
    If Assemble <> OZ80_ERROR_NONE Then GoTo ErrorOut
    
    'Stage 2: Assemble _
     ----------------------------------------------------------------------------------
    'Default to 32 KB ROM (the ROM size will double each time it overflows)
    Let ROMSize = 2 ^ 15
    ReDim ROM(0 To ROMSize - 1) As Byte
    
    'Phase 1: _
     ..................................................................................
    Call Token.Rewind
    Let Assemble = ScopeRoot()
    If Assemble <> OZ80_ERROR_NONE Then GoTo ErrorOut
    
    'Phase 2: _
     ..................................................................................
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Phase 2")
    'TODO
    
    RaiseEvent Message(0, OZ80_LOG_ACTION, "Assembly Complete")
    RaiseEvent Message(1, OZ80_LOG_INFO, "ROM size: " & (ROMSize \ 1024) & " KB")
    RaiseEvent Message(1, OZ80_LOG_INFO, "Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    
ErrorOut:
    'Clean up
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
        *   `{operator}` can be any one of `+ - * / ^ \ | &`

'Scope:     1           2           3       4       5       6       7       8 _
 --------------------------------------------------------------------------------------
'<text>     "           {ASCII}     "
'
'<expr>     <text>
'           <number>    [{operator} <expr>]
'           #var        [{operator} <expr>]
'           :label      [{operator} <expr>]
'           {           <expr>      }
'
'
'<list>     <text>      [,          <list>]
'           <expr>      [,          <list>]
'
'......................................................................................
'
'<root>     INCLUDE     <text>
'
'           OBJECT      #object     <object>
'
'           PROC        :label      [PARAMS register
'                                           number]
'                                   [RETURN register
'                                           number]
'                                   {       <code>
'                                   }
'
'           SECTION     [BANK       <expr>]
'                       [SLOT       <expr>]
'
'           <data>
'
'<code>     .label
'           DATA        <list>
'           BYTE        <list>
'           WORD        <list>
'
'           <data>
'
'           <asm>
'
'<data>     VAR         #variable   [<list>]
'
'           TABLE       :label      <code>
'
'           STRUCT      :label      AS      #object         <struct>

'ScopeRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Function ScopeRoot() As OZ80_ERROR
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
    'These are used for evaluating expressions in parameters
    Dim ReturnValue As Long, ValueIsIndefinite As Boolean
    'For remembering item names (labels, variables &c.)
    Dim StringKey As Long
    
Continue:
    'As we loop, other functions will move the Token pointer forward (such as when _
     evaluating expressions). Check if we've hit the end of the file
    If Token.EOF Then Exit Function
    
    'ERROR: Not a keyword on root scope
    If Not Token.IsKeyword Then Let ScopeRoot = RaiseError( _
        OZ80_ERROR_EXPECTED_ROOT, Token.Line, Token.Col _
    ): Exit Function
    
    '----------------------------------------------------------------------------------
    '`INCLUDE <text>` _
     ----------------------------------------------------------------------------------
    If Token.Kind = TOKEN_KEYWORD_INCLUDE Then
        'Move forward past the keyword into the first parameter
        'ERROR: The file ended before the first parameter
        If Not Token.Forward() Then GoTo ErrEndOfFile
    
        'TODO: Capture the filename
        'TODO: Parse the file
        'TODO: Return to the previous file
        Stop
        
    '----------------------------------------------------------------------------------
    '`OBJECT #<objectName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_OBJECT Then
        'Move forward past the keyword into the first parameter
        'ERROR: The file ended before the first parameter
        If Not Token.Forward() Then GoTo ErrEndOfFile
        
        'TODO
        Stop
        
    '----------------------------------------------------------------------------------
    '`PROC :<name> [SECTION ::<name>] [PARAMS <list>] [RETURN <list>] { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_PROC Then
        'Move forward past the keyword into the first parameter
        'ERROR: The file ended before the first parameter
        If Not Token.Forward() Then GoTo ErrEndOfFile
        
        'ERROR: No label name for procedure
        If Token.Kind <> TOKEN_LABEL Then Let ScopeRoot = RaiseError( _
            OZ80_ERROR_EXPECTED_PROC_NAME, Token.Line, Token.Col _
        ): Exit Function
        
        'TODO: Normalise the label name and add to the label database _
         (we shouldn't be basing things off of the string database IDs)
        
        'Remember the String ID of the label name
        Let StringKey = Token.Value
        
        'Create the new block to assemble into
        Call AddBlock(StringKey)
        
        'Check for the optional parameters for a PROC, _
         namely "SECTION", "PARAMS", "RETURN"
        
        'The file cannot end until the block has been opened and closed
        If Not Token.Forward() Then GoTo ErrEndOfFile
        
ProcParams:
        If Token.Kind = TOKEN_KEYWORD_SECTION Then '...................................
            'Move forward to get the section name
            'ERROR: file ended before the section name?
            If Not Token.Forward() Then Let ScopeRoot = RaiseError( _
                OZ80_ERROR_ENDOFFILE, Token.Line, Token.Col _
            ): Exit Function
            
            'ERROR: Not a section name following the section keyword
            If Token.Kind <> TOKEN_SECTION Then Let ScopeRoot = RaiseError( _
                OZ80_ERROR_EXPECTED_SECTION_NAME, Token.Line, Token.Col _
            ): Exit Function
            
            'TODO: Has the section been defined?
            
            'TODO: Normalise the section name?
            
            'Provide the block with the section ID
            Let Block.Section = Token.Value
            
        ElseIf Token.Kind = TOKEN_KEYWORD_PARAMS Then '................................
            'TODO: Skip through the params list
        ElseIf Token.Kind = TOKEN_KEYWORD_RETURN Then '................................
            'TODO: Skip through the returns list
        End If
        
        'TODO: Look for the opening brace
        'TODO: Create the new block
        
    '----------------------------------------------------------------------------------
    '`SECTION ::<sectionName> [BANK <expr>] [SLOT <expr>]` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_SECTION Then
        Let ScopeRoot = ScopeSection()
    
    '----------------------------------------------------------------------------------
    '`STRUCT :<structureName> AS #<objectName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_STRUCT Then
        'TODO
        
    '----------------------------------------------------------------------------------
    '`TABLE :<tableName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_TABLE Then
        'TODO
    
    '----------------------------------------------------------------------------------
    '`VAR #<variable> [<expr>]` _
     ----------------------------------------------------------------------------------
    ElseIf Token.Kind = TOKEN_KEYWORD_VAR Then
        'Move forward past the keyword into the first parameter
        'ERROR: The file ended before the first parameter
        If Not Token.Forward() Then GoTo ErrEndOfFile
        
        'ERROR: Not a variable name for `VAR`
        If Token.Kind <> TOKEN_VARIABLE Then Let ScopeRoot = RaiseError( _
            OZ80_ERROR_EXPECTED_VAR_NAME, Token.Line, Token.Col _
        ): Exit Function
        
        'Fetch the key for the label name in the string database, _
         (i.e. remember the label name before moving past the token)
        Let StringKey = Token.Value
        
        'Move forward to the value parameter
        'NOTE: the value parameter is optional, and the file could end right here -- _
               if so, it'll skip over retreiving the value
        If Token.Forward() And Token.IsExpression Then
            'Calculate the value of the expression
            Let ScopeRoot = ScopeExpression(ReturnValue, ValueIsIndefinite)
            'If the expression is faulty, an error will already have been raised, _
             we only need to fold upward to terminate the assembling
            If ScopeRoot <> OZ80_ERROR_NONE Then Exit Function
            'ERROR: A value must be known in order to define a variable, i.e. _
             yet-unknown values, such as label addresses, cannot be used
            If ValueIsIndefinite Then Let ScopeRoot = RaiseError( _
                OZ80_ERROR_INDEFINITE, Token.Line, Token.Col _
            ): Exit Function
        Else
            'There is no given value for the label, re-use the existing value, _
             otherwise use zero
            If Variables.Exists(StringKey) Then
                Let ReturnValue = Variables.Item(StringKey)
            Else
                Let ReturnValue = 0
            End If
            Let ValueIsIndefinite = False
        End If
        
        'Is the variable already defined?
        If Variables.Exists(StringKey) Then
            'Update the value
            Let Variables(StringKey) = ReturnValue
        Else
            'Create the named value
            Call Variables.Add(StringKey, ReturnValue)
        End If
        
        'Log the result of this
        RaiseEvent Message(1, OZ80_LOG_DEBUG, _
            "VAR " & Token.Strings(StringKey) & " " & _
                Right$(Space(40) & ReturnValue, 40 - Len(Token.Strings(StringKey))) & _
                " $" & Right$("000" & Hex(ReturnValue), 4) _
        )
        
        'TODO: Are we on the next token? If no value parameter, then NO!
        GoTo Continue
    
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
ErrUnexpected:
ErrEndOfFile:
        Let ScopeRoot = RaiseError( _
            OZ80_ERROR_ENDOFFILE, Token.Line, Token.Col _
        ): Exit Function
    End If
    
    'If an error had occured, the return value of this function will have been set _
     to the error number and the error message raised (by event)
    If ScopeRoot <> OZ80_ERROR_NONE Then Exit Function
    
    'As long as there is content left to process in the file, keep going
    If Token.Forward() Then GoTo Continue
End Function

'ScopeExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ScopeExpression( _
    ByRef ReturnValue As Long, ByRef ValueIsIndefinite As Boolean _
) As OZ80_ERROR
    Dim Operand1 As Long                'First value in an expression
    Dim Operand2 As Long                'Next value in an expression
    Dim Operator As OZ80_TOKEN          'Operator to use with the expressions
    
    Let ValueIsIndefinite = False
    
    'Begin with fetching the first operand
    Let ScopeExpression = ScopeOperand(Operand1, ValueIsIndefinite)
    If ScopeExpression <> OZ80_ERROR_NONE Then Stop
    Let ReturnValue = Operand1
    
'    'The file could legitimately end with an expression, so don't error out if the _
'     token stream comes to an end since the expression is valid. The parent scopes _
'     will handle an unexpected end error
'    If NextToken() = False Then Exit Function
    
    'If next token is not an operator assume the expression has finished, _
      the parent will handle whatever the next token is used for
    If Not Token.IsOperator Then Exit Function
    
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = Token.Kind
    
    'Get the second operand
    'TODO: Error - unexpected end
    If Token.Forward() = False Then Stop
    Let ScopeExpression = ScopeOperand(Operand2, ValueIsIndefinite)
    'Cannot have an operator hanging there without an operand following
    If ScopeExpression <> OZ80_ERROR_NONE Then
        'TODO: Error - invalid use of operator
        Stop
    End If
    
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If ValueIsIndefinite = True Then
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
    If Token.IsOperator Then
        'Set the current result as the first operand and go around again
        Let Operand1 = ReturnValue
        GoTo Continue
    End If
End Function

'ScopeMemory : Work out a memory reference, i.e. `($????)` / `(ix+$??)` _
 ======================================================================================
Private Function ScopeMemory( _
    ByRef Kind As OZ80_PARAM, _
    ByRef Value As Long, ByRef ValueIsIndefinite As Boolean _
) As OZ80_ERROR
    'If the current token is an open parentheses move forward into the expression
    If Token.Kind = OZ80_TOKEN.TOKEN_PARENOPEN Then
        If Not Token.Forward() Then Stop
    End If
    
    'The first item in the expression can be a register, typically HL, IX or IY
    If Token.IsRegister Then
        'The `IN` & `OUT` instructions accept only the C register
        If Token.Kind = TOKEN_Z80_C Then
            Let Kind = PARAM_C
        
        '`LD` accepts registers BC & DE
        ElseIf Token.Kind = TOKEN_Z80_BC Then
            Let Kind = PARAM_BC
        ElseIf Token.Kind = TOKEN_Z80_DE Then
            Let Kind = PARAM_DE
        ElseIf Token.Kind = TOKEN_Z80_HL Then
            Let Kind = PARAM_HL
            
        ElseIf Token.Kind = TOKEN_Z80_IX Then
            Let Kind = PARAM_IX
            'IX can have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf Token.Kind = TOKEN_Z80_IY Then
            Let Kind = PARAM_IY
            'IY can also have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf Token.Kind = TOKEN_Z80_SP Then
            Let Kind = PARAM_SP
        Else
            'ERROR: register not allowed here
            Stop
        End If
        
        'Move to the next token, then check that it's the end of the memory reference
        'TODO: out of tokens
        If Not Token.Forward() Then Stop
        GoTo EndScope
        
AddExpr:
        'Look for the `+` following the register name
        If Not Token.Forward() Then Stop
        'If present, skip over it and check for an expression
        If Token.IsOperator Then
            If Token.Kind = TOKEN_OPERATOR_ADD Then
                If Not Token.Forward() Then Stop
            End If
            
        'If it's just a register in parens, then finish
        ElseIf Token.Kind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
            GoTo Finish
        End If
    End If
    
    'Try and calculate the expression. It might come back unknown and this will be _
     passed on to the parent to decide if that's a problem or not
    'TODO: handle overflow
    Let ScopeMemory = ScopeExpression(Value, ValueIsIndefinite)
    'Mark this as a value in the memory reference
    Let Kind = Kind Or PARAM_VAL

EndScope:
    If Token.Kind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
        Call Token.Forward
        GoTo Finish
    End If
    'ERROR: Unexpected token in memory expression
    Stop

Finish:
    'Finally, mark this as a memory reference
    Let Kind = Kind Or PARAM_MEM
End Function

'ScopeOperand : Get the value of one half of a calculation _
 ======================================================================================
Private Function ScopeOperand( _
    ByRef ReturnValue As Long, ByRef ValueIsIndefinite As Boolean _
) As OZ80_ERROR
    'Is it a named value?
    'TODO: Handle scope expansion for names
    If Token.Kind = TOKEN_LABEL _
    Or Token.Kind = TOKEN_PROPERTY_USE Then
        'TODO: get the full label from the property
        
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Labels.Exists(Token.Value) = False Then
            Let ValueIsIndefinite = True
        Else
            Let ReturnValue = Variables(Token.Value)
        End If
        
    ElseIf Token.Kind = TOKEN_VARIABLE Then
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Variables.Exists(Token.Value) = False Then
            Let ValueIsIndefinite = True
        Else
            Let ReturnValue = Variables(Token.Value)
        End If
        
    'This is an easy one
    ElseIf Token.Kind = TOKEN_NUMBER Then
        Let ReturnValue = Token.Value
        
    Else
        'ERROR: Not a valid operand. We won't produce an error message here as the _
         error might be exepected (i.e. an optional parameter). We fold upwards and _
         whomever knows best will deal with the error
        Debug.Print Token.Line & ", " & Token.Col
        Stop
        Let ScopeOperand = OZ80_ERROR_EXPRESSION
    End If
    
    'Move beyond the current token now it's been processed
    Call Token.Forward
End Function

'ScopeSection : Process a section directive _
 ======================================================================================
Private Function ScopeSection() As OZ80_ERROR
    'Format: _
     `SECTION ::<name> [BANK <expr>] [SLOT <expr>]`
    
    'ERROR: The file ended before the first parameter
    If Not Token.Forward() Then Let ScopeSection = RaiseError( _
        OZ80_ERROR_ENDOFFILE, Token.Line, Token.Col _
    )
    
    'ERROR: Not a section name following the section keyword
    If Token.Kind <> TOKEN_SECTION Then Let ScopeSection = RaiseError( _
        OZ80_ERROR_EXPECTED_SECTION_NAME, Token.Line, Token.Col _
    ): Exit Function
    
    'Keep ahold of the Section identifier
    Dim SectionKey As Long
    Let SectionKey = Token.Value
    
    'Has this section already been defined?
    If SectionNames.Exists(SectionKey) Then Let ScopeSection = RaiseError( _
        OZ80_ERROR_DUPLICATE_SECTION, Token.Line, Token.Col _
    ): Exit Function
    
    'Define the section name - fetch the string from the TokenStream's _
     string database (identified by the CRC value on the token)
    Call SectionNames.Add(SectionKey, Token.Strings(SectionKey))
    'Default bank as undefined, the assembler will choose
    Call SectionBanks.Add(SectionKey, -1)
    Call SectionSlots.Add(SectionKey, 0)
    
    '----------------------------------------------------------------------------------
    
    'If the file ends here it's valid, as the params are optional
    If Not Token.Forward() Then Exit Function
    
NextParam:
    'We need to remember the current token as the call to `ScopeExpression` will _
     move the token pointer forward, invalidating `Token.Kind` at this scope
    Dim ParamToken As OZ80_TOKEN
    Let ParamToken = Token.Kind
    
    If ParamToken = TOKEN_KEYWORD_BANK _
    Or ParamToken = TOKEN_KEYWORD_SLOT _
    Then
        'Don't allow a duplicate param
        Dim IsBankDefined As Boolean, IsSlotDefined As Boolean
        If IsBankDefined And (ParamToken = TOKEN_KEYWORD_BANK) Then
            'TODO: specific error message
            Let ScopeSection = RaiseError( _
                OZ80_ERROR_DUPLICATE_BANK, Token.Line, Token.Col _
            ): Exit Function
        ElseIf IsSlotDefined And (ParamToken = TOKEN_KEYWORD_SLOT) Then
            'TODO: specific error message
            Let ScopeSection = RaiseError( _
                OZ80_ERROR_DUPLICATE_SLOT, Token.Line, Token.Col _
            ): Exit Function
        End If
        
        'ERROR: File ending before we have the value
        If Not Token.Forward() Then Let ScopeSection = RaiseError( _
            OZ80_ERROR_ENDOFFILE, Token.Line, Token.Col _
        ): Exit Function
        
        'ERROR: If not an expression, it can't be accepted as a value
        If Not Token.IsExpression Then Let ScopeSection = RaiseError( _
            OZ80_ERROR_EXPECTED, Token.Line, Token.Col _
        ): Exit Function
        
        'Evaluate the expression
        Dim ReturnValue As Long, ValueIsIndefinite As Boolean
        Let ScopeSection = ScopeExpression(ReturnValue, ValueIsIndefinite)
        
        'You can't use an indefinite value for a bank/slot number
        If ValueIsIndefinite Then Let ScopeSection = RaiseError( _
            OZ80_ERROR_INDEFINITE, Token.Line, Token.Col _
        ): Exit Function
        
        'TODO: Is the value within range?
        
        'Assign the value to the right place
        If ParamToken = TOKEN_KEYWORD_BANK Then
            Let SectionBanks(SectionKey) = ReturnValue
            Let IsBankDefined = True
        ElseIf ParamToken = TOKEN_KEYWORD_SLOT Then
            Let SectionSlots(SectionKey) = ReturnValue
            Let IsSlotDefined = True
        End If
        
        'Loop around to check for a second param
        'NOTE: `ScopeExpression` will have moved the token pointer ahead, _
               so we should already be on the next keyword
        GoTo NextParam
    End If
End Function

'ScopeZ80 : Process a Z80 instruction _
 ======================================================================================
Private Function ScopeZ80() As OZ80_ERROR
    'Remember which Z80 instruction is being processed
    Dim Instruction As OZ80_TOKEN: Let Instruction = Token.Kind
    'Remember the line number of this since a label or property on a different line _
     will be a definition, not a parameter
    Dim Line As Long: Let Line = Token.Line
    'Skip over the Z80 instruction token now that we know it
    If Not Token.Forward() Then Stop
    
    'If there appears to be a parameter, skip over doing parameterless instructions
    If Token.IsParameter = True Then GoTo With1Parameter
    
    'TODO: Error handling return
    Call Block.AddInstruction(Instruction)
    Exit Function
    
    'TODO: Handle POP / PUSH
    
    '----------------------------------------------------------------------------------
With1Parameter:
    'Parameter expressions might return an unknown value _
     (such as when a label, which has not yet been addressed, is used)
    Dim Param1 As oz80Param, Param2 As oz80Param
    'Get the first parameter
    Let ScopeZ80 = GetParameter(Param1)
    If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
    
    'Check if a second parameter exists
    If Token.IsParameter = True Then GoTo With2Parameters
    
    'TODO: Error handling return
    Call Block.AddInstruction(Instruction, Param1.Kind, Param1.Value)
    
    Exit Function

    '----------------------------------------------------------------------------------
With2Parameters:
    'Fetch the second parameter
    Let ScopeZ80 = GetParameter(Param2)
    If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
            
    'TODO: Error handling return
    Call Block.AddInstruction( _
        Instruction, Param1.Kind, Param1.Value, Param2.Kind, Param2.Value _
    )
End Function

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'AddBlock : Begin a new block to hold assembled Z80 code / data _
 ======================================================================================
Private Sub AddBlock(ByRef StringID As Long)
    'Create and attach the new block
    Set Block = New oz80_Block
    'Add it to the collection
    Call Blocks.Add(StringID, Block)
End Sub

'GetParameter : Fetch the type / value of a Z80 instruction parameter _
 ======================================================================================
Private Function GetParameter(ByRef Result As oz80Param) As OZ80_ERROR
    'Clear any previous result since this is ByRef
    With Result
        Let .Kind = 0
        Let .Value = 0
        Let .ValueIndefinite = False
    End With
    
    'If the current token is a register, map this to the parameter kind
    If Token.IsRegister Then '...................................................
        If Token.Kind = TOKEN_Z80_A Then Let Result.Kind = PARAM_A
        If Token.Kind = TOKEN_Z80_B Then Let Result.Kind = PARAM_B
        If Token.Kind = TOKEN_Z80_C Then Let Result.Kind = PARAM_C
        If Token.Kind = TOKEN_Z80_D Then Let Result.Kind = PARAM_D
        If Token.Kind = TOKEN_Z80_E Then Let Result.Kind = PARAM_E
        If Token.Kind = TOKEN_Z80_H Then Let Result.Kind = PARAM_H
        If Token.Kind = TOKEN_Z80_L Then Let Result.Kind = PARAM_L
        If Token.Kind = TOKEN_Z80_I Then Let Result.Kind = PARAM_I
        If Token.Kind = TOKEN_Z80_R Then Let Result.Kind = PARAM_R
        If Token.Kind = TOKEN_Z80_AF Then Let Result.Kind = PARAM_AF
        If Token.Kind = TOKEN_Z80_BC Then Let Result.Kind = PARAM_BC
        If Token.Kind = TOKEN_Z80_DE Then Let Result.Kind = PARAM_DE
        If Token.Kind = TOKEN_Z80_HL Then Let Result.Kind = PARAM_HL
        If Token.Kind = TOKEN_Z80_IX Then Let Result.Kind = PARAM_IX
        If Token.Kind = TOKEN_Z80_IY Then Let Result.Kind = PARAM_IY
        If Token.Kind = TOKEN_Z80_SP Then Let Result.Kind = PARAM_SP
        'Now move beyond the register token
        Call Token.Forward
    
    'If the token is a memory reference, such as `(ix+$8)`, _
     then get the associated register and value
    ElseIf Token.Kind = TOKEN_PARENOPEN Then '.........................................
        Let GetParameter = ScopeMemory( _
            Result.Kind, Result.Value, Result.ValueIndefinite _
        )
        'TODO: Error - the memory reference is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
        
    'Anything else classes as an expression (number / labels / variables / operators)
    ElseIf Token.IsExpression Then '...................................................
        Let Result.Kind = PARAM_VAL
        Let GetParameter = ScopeExpression(Result.Value, Result.ValueIndefinite)
        'TODO: Error - the expression is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
    
    Else
        'ERROR: Not an expression
        Let GetParameter = RaiseError( _
            OZ80_ERROR_EXPRESSION_Z80, Token.Line, Token.Col _
        )
    End If
End Function

'ParseFile : Parse a source code file into a TokenStream object _
 ======================================================================================
Private Function ParseFile(ByRef FilePath As String) As OZ80_ERROR
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
     Note that file/folder case will be corrected and could throw a false negative
    If LCase$(FSO.GetAbsolutePathName(FilePath)) = LCase$(FilePath) Then
        'We'll skip ahead to checking the file exists
        Let TestPath = FilePath
    Else
        'Path is relative! If no other file has been parsed yet we will need to use _
         the current directory as a base
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
    If FSO.FileExists(TestPath) = False Then Let ParseFile = RaiseError( _
        OZ80_ERROR_FILENOTFOUND, Token.Line, Token.Col _
    ): GoTo Finish
    
    '[3] Parse the file into a TokenStream: _
     ----------------------------------------------------------------------------------
    'We'll use the CRC of the filepath string as its key/ID
    Dim CRC As Long
    Let CRC = oz80.CRC.Calculate(TestPath)
    
    'If the file has already been parsed, re-use it
    If Files.Exists(CRC) = True Then
        'Reference the existing TokenStream object
        Set Token = Files.Item(CRC)
    Else
        'Create a tokeniser object to hold the machine representation of the text file; _
         the assembler doesn't work with the original text directly
        Set Token = New oz80_TokenStream
        'Explode the source code file into tokens
        Let ParseFile = Token.Tokenise(TestPath)
        'If the file failed to parse, an error code / message will already have been _
         raised via an event, we only need to exit
        If ParseFile <> OZ80_ERROR_NONE Then GoTo Finish
        'Add the TokenStream to the files collection
        Call Files.Add(CRC, Token)
    End If
    
Finish:
    Set FSO = Nothing
    
    'If there were any error, an error event and message would already have been _
     raised, we only need to fold upwards and the parent will handle terminating the _
     assembling process early
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Function RaiseError( _
    ByRef Number As OZ80_ERROR, ByRef Line As Long, ByRef Col As Long _
) As OZ80_ERROR
    'Fetch the detailed error text
    Dim Title As String, Description As String
    Call oz80.GetOZ80Error(Number, Title, Description)
    
    'Pass the error up to the controller
    RaiseEvent Error(Number, Title, Description, Line, Col)
    
    'We return the same error number we were given only to make the tail recursion _
     simpler, e.g. `let ScopeRoot = RaiseError(OZ80_ERROR_...`
    Let RaiseError = Number
End Function
