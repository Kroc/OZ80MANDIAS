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

'/// API //////////////////////////////////////////////////////////////////////////////

'Cast String to Array: _
 --------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB string directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for VB arrays _
 <msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx>
Private Type SAFEARRAY1D
    cDims      As Integer   'The count of dimensions
    fFeatures  As Integer   'Flags used by the SafeArray
    cbElements As Long      'The size of an array element
    cLocks     As Long      'Lock count
    pvData     As Long      'Pointer to the data
    cElements  As Long      'The number of elements in the dimension
    lLbound    As Long      'The lower bound of the dimension
End Type

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

'Undocumented VB6 API to get at the array pointer
Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Byte            'Empty array that will be hacked

'/// CONSTANTS ////////////////////////////////////////////////////////////////////////

'This structure is used when parsing parameters of Z80 instructions
Private Type oz80Param
    Kind As OZ80_PARAM                  'The kind of parameter, e.g. register, value
    Value As Long                       'The numerical value for the parameter
    ValueIndefinite As Boolean          'If the parameter has unknown values (labels)
End Type

'This is the official polynomial used by CRC32 in PKZip. _
 Often the polynomial is shown reversed (04C11DB7)
Private Const CRC_POLYNOMIAL As Long = &HEDB88320

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

'Store a look up table for the CRC calculations
Private CRCTable(0 To 255) As Long

'For timing this process
Private Stopwatch As PerformanceCounter

'Each source code file is broken down into a machine-readable representation which _
 we store here. When parsing one file and we come across an `INCLUDE` statement, _
 another token stream gets added here
Private Files As Scripting.Dictionary

'The current source file being processed, attached such that we can get the _
 message and error events from it
Private WithEvents TokenStream As oz80_TokenStream
Attribute TokenStream.VB_VarHelpID = -1

'Caches of the current token information since dereferencing the object is slow
Private TokenKind As OZ80_TOKEN
Private TokenValue As Long
Private TokenLine As Long
Private TokenCol As Long

'The line number of the previous token, so that we can tell when a label/property _
 definition occurs on a new line and isn't therefore part of a list
Private PrevLine As Long

'The token stream doesn't store the text of named items, merely a reference to the _
 string stored here. We use this for error messages and concatenating properties to _
 their parent variable / label
Private Strings As Scripting.Dictionary

'The numerical values of variables / labels / properties &c. will be stored here. _
 The key for a name is the CRC32 of the string (from `Strings` above)

'Variables and Labels are kept separately as the first assembling pass determines _
 the final location of the labels in the ROM and the second pass then fills in the _
 expressions that could not be evaluated, but we need to clear out the variable _
 values produced during the first pass so that they don't change the behaviour of _
 the second pass

Private Variables As Scripting.Dictionary
Private Labels As Scripting.Dictionary

'Each block of code / data is assembled separately from the others and with no final _
 position in the ROM in mind until all source has been parsed. Once all source is _
 parsed, the banks are chosen for the blocks and the assembly copied into the ROM
Private Blocks As Scripting.Dictionary
'The current block being assembled into
Private Block As oz80_Block

'--------------------------------------------------------------------------------------

Private ROM() As Byte                   'The output!
Private ROMSize As Long                 'Size of the ROM (32/64/128/256/512/1024 KB)

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'Progress messages are passed through an event, the handler can choose to show them _
 or not. I didn't want to spam `Debug.Print` output
Event Message(ByRef Text As String)

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
    
    'Create the CRC look-up table _
     (I will not pretend to understand how this works other than it goes through _
      shifting and flipping bits, make of that what you will)
    Dim i As Long, Value As Long
    For i = 0 To 255
        Value = i
        Dim ii As Long
        For ii = 8 To 1 Step -1
            If (Value And 1) Then
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                Let Value = Value Xor CRC_POLYNOMIAL
            Else
                Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next
        Let CRCTable(i) = Value
    Next
    
    With CastStrSA1D
        .cDims = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .cbElements = 2&
        .cLocks = 1&
        .lLbound = 0&
    End With
    
    Let CastStrPtr = msvbvm60_VarPtr(CastStrArr())
    Call msvbvm60_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Set Variables = Nothing
    Set Labels = Nothing
    Set Strings = Nothing
    Set Blocks = Nothing
    Set Stopwatch = Nothing
    Set TokenStream = Nothing
    Set Files = Nothing
    Erase ROM, CRCTable
    
    'Clean up our array hacking (for casting a string to an array) _
     otherwise VB will crash when it frees up the memory
    Call msvbvm60_PutMem4(CastStrPtr, 0&)
End Sub

'EVENT <TokenStream> Error : Detailed error message from a TokenStream object _
 ======================================================================================
Private Sub TokenStream_Error( _
    ByVal Number As OZ80_ERROR, ByRef Title As String, ByRef Description As String, _
    ByVal Line As Long, ByVal Col As Long _
)   'Pass this up to the controller, this class will shortly end processing
    RaiseEvent Error(Number, Title, Description, Line, Col)
End Sub

'EVENT <TokenStream> Message : Basic progress message from a TokenStream object _
 ======================================================================================
Private Sub TokenStream_Message(Text As String)
    'It's up to the controller to decide to display (or `Debug.Print` any messages), _
     we're not going to unnecessarily spam their console
    RaiseEvent Message(Text)
End Sub

'EVENT <TokenStream> StringAdded : While parsing a source file, a string name was read _
 ======================================================================================
Private Sub TokenStream_StringAdded(ByRef Key As Long, ByRef StringData As String)
    'We need to provide a unique key, we use the CRC32 of the string. _
     When this event finishes, the TokenStream will add a token using this key
    Let Key = CRCString(StringData)
    'Add the label/variable/property name to the string storage
    If Strings.Exists(Key) = False Then Call Strings.Add(Key, StringData)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Assemble : Compile a Z80 binary from a given starting source code file _
 ======================================================================================
Public Function Assemble(ByVal FilePath As String) As OZ80_ERROR
    'Announce ourselves
    RaiseEvent Message("OZ80MANDIAS v" & App.Major & "." & App.Minor & "," & App.Revision)
    RaiseEvent Message("* Assembling token stream")
    
    'We'll measure the time this takes
    Call Stopwatch.Reset
    
    'Initialise the storage for label/variable/property names
    Set Strings = New Scripting.Dictionary
    'Initialise the storage for the TokenStream representations of the source files
    Set Files = New Scripting.Dictionary
    
    'Stage 1: Parse Source _
     ----------------------------------------------------------------------------------
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
    Call RewindTokenStream
    Let Assemble = ScopeRoot()
    
    'Phase 2: _
     ..................................................................................
    'TODO
    
    RaiseEvent Message("- Final ROM size is " & (ROMSize \ 1024) & " KB")
    RaiseEvent Message("- Time: " & Format$(Stopwatch.Elapsed / 1000, "0.000") & "s")
    
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
'<root>     PROC        :label      [PARAMS register
'                                           number]
'                                   [RETURN register
'                                           number]
'                                   {       <code>
'                                   }
'
'           INCLUDE     <text>
'
'           OBJECT      #object     <object>
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
'

'ScopeRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Function ScopeRoot() As OZ80_ERROR
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
    'These are used for evaluating expressions in parameters
    Dim ReturnValue As Long, ValueIsIndefinite As Boolean
    
Continue:
    'ERROR: Not a keyword on root scope
    If Not TokenStream.IsKeyword Then Let ScopeRoot = RaiseError( _
        OZ80_ERROR_UNEXPECTED, _
        "Unexpected content at this position", _
        "Only the keywords INCLUDE, OBJECT, PROC, STRUCT, TABLE or VAR " & _
        "are allowed here.", TokenLine, TokenCol _
    ): Exit Function
    
    'Remember the keyword before we move past it
    Dim Keyword As OZ80_TOKEN
    Let Keyword = TokenKind
    'Move forward past the keyword into the first parameter
    'ERROR: The file ended before the first parameter
    If NextToken() = False Then GoTo ErrEndOfFile
    
    '----------------------------------------------------------------------------------
    '`INCLUDE <text>` _
     ----------------------------------------------------------------------------------
    If Keyword = TOKEN_KEYWORD_INCLUDE Then
        'TODO: Capture the filename
        'TODO: Parse the file
        'TODO: Return to the previous file
        
    '----------------------------------------------------------------------------------
    '`VAR #<variable> [<list>]` _
     ----------------------------------------------------------------------------------
    ElseIf Keyword = TOKEN_KEYWORD_VAR Then
        
        'ERROR: Not a variable name for `VAR`
        If TokenKind <> TOKEN_VARIABLE Then Let ScopeRoot = RaiseError( _
            OZ80_ERROR_UNEXPECTED, _
            "Unexpected content at this point", _
            "Expected a variable name.", TokenLine, TokenCol _
        ): Exit Function
        
        'Fetch the key for the label name in the string database, _
         (i.e. remember the label name before moving past the token)
        Dim NameKey As Long
        Let NameKey = TokenValue
        
        'Move forward to the value parameter
        'NOTE: the value parameter is optional, and the file could end right here -- _
               if so, it'll skip over retreiving the value
        If NextToken() = True And TokenStream.IsExpression Then
            'Calculate the value of the expression
            Let ScopeRoot = ScopeExpression(ReturnValue, ValueIsIndefinite)
            'If the expression is faulty, an error will already have been raised, _
             we only need to fold upward to terminate the assembling
            If ScopeRoot <> OZ80_ERROR_NONE Then Exit Function
            'ERROR: A value must be known in order to define a variable, i.e. _
             yet-unknown values, such as label addresses, cannot be used
            If ValueIsIndefinite = True Then Let ScopeRoot = RaiseError( _
                OZ80_ERROR_INDEFINITEVALUE, _
                "Indefinite value cannot be used here", _
                "A variable cannot be defined with an indefinite value, that is, " & _
                    "an expression containing a yet-unknown value, such as a label. " & _
                    "label addresses are not set until after assembly.", _
                    TokenLine, TokenCol _
            ): Exit Function
        Else
            'There is no given value for the label, re-use the existing value, _
             otherwise use zero
            If Variables.Exists(NameKey) = True Then
                Let ReturnValue = Variables.Item(NameKey)
            Else
                Let ReturnValue = 0
            End If
            Let ValueIsIndefinite = False
        End If
        
        'Is the variable already defined?
        If Variables.Exists(NameKey) Then
            'Update the value
            Let Variables(NameKey) = ReturnValue
        Else
            'Create the named value
            Call Variables.Add(NameKey, ReturnValue)
        End If
        
        RaiseEvent Message("= VAR " & Strings(NameKey) & " " & ReturnValue)
        
        'TODO: Are we on the next token? If no value parameter, then NO!
        GoTo Continue
        
    '----------------------------------------------------------------------------------
    '`OBJECT #<objectName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Keyword = TOKEN_KEYWORD_OBJECT Then
        'TODO
        
    '----------------------------------------------------------------------------------
    '`PROC :<name> [SECTION ::<name>] [PARAMS <list>] [RETURN <list>] { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Keyword = TOKEN_KEYWORD_PROC Then
        'ERROR: No label name for procedure
        If TokenKind <> TOKEN_LABEL Then Let ScopeRoot = RaiseError( _
            OZ80_ERROR_UNEXPECTED, _
            "Unexpected content at this point", _
            "A label name must always follow the `PROC` keyword", _
            TokenLine, TokenCol _
        ): Exit Function
        
        'TODO: Capture the label name
        
        'TODO: Section? (normalise the label name)
        
        'TODO: Skip through the params list
        
        'TODO: Skip through the returns list
        
        'TODO: Create the new block
        
        
    '----------------------------------------------------------------------------------
    '`STRUCT :<structureName> AS #<objectName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Keyword = TOKEN_KEYWORD_STRUCT Then
        'TODO
        
    '----------------------------------------------------------------------------------
    '`TABLE :<tableName> { [...] }` _
     ----------------------------------------------------------------------------------
    ElseIf Keyword = TOKEN_KEYWORD_TABLE Then
        'TODO
    
    'ERROR: Not a valid keyword at this scope! _
     ----------------------------------------------------------------------------------
    Else
ErrUnexpected:
ErrEndOfFile:
        Let ScopeRoot = RaiseError( _
            OZ80_ERROR_ENDOFFILE, _
            "Unexpected end of file", _
            "Parameter expected!", TokenLine, TokenCol _
        ): Exit Function
    End If
    
    If NextToken() = True Then GoTo Continue
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
    If Not TokenStream.IsOperator Then Exit Function
    
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = TokenKind
    
    'Get the second operand
    'TODO: Error - unexpected end
    If NextToken() = False Then Stop
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
    If TokenStream.IsOperator Then
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
    If TokenKind = OZ80_TOKEN.TOKEN_PARENOPEN Then
        If NextToken() = False Then Stop
    End If
    
    'The first item in the expression can be a register, typically HL, IX or IY
    If TokenStream.IsRegister Then
        'The `IN` & `OUT` instructions accept only the C register
        If TokenKind = TOKEN_Z80_C Then
            Let Kind = PARAM_C
        
        '`LD` accepts registers BC & DE
        ElseIf TokenKind = TOKEN_Z80_BC Then
            Let Kind = PARAM_BC
        ElseIf TokenKind = TOKEN_Z80_DE Then
            Let Kind = PARAM_DE
        ElseIf TokenKind = TOKEN_Z80_HL Then
            Let Kind = PARAM_HL
            
        ElseIf TokenKind = TOKEN_Z80_IX Then
            Let Kind = PARAM_IX
            'IX can have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf TokenKind = TOKEN_Z80_IY Then
            Let Kind = PARAM_IY
            'IY can also have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf TokenKind = TOKEN_Z80_SP Then
            Let Kind = PARAM_SP
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
        If TokenStream.IsOperator Then
            If TokenKind = TOKEN_OPERATOR_ADD Then
                If NextToken() = False Then Stop
            End If
            
        'If it's just a register in parens, then finish
        ElseIf TokenKind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
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
    If TokenKind = OZ80_TOKEN.TOKEN_PARENCLOSE Then
        Call NextToken
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
    If TokenKind = TOKEN_LABEL _
    Or TokenKind = TOKEN_PROPERTY Then
        'TODO: get the full label from the property
        
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Labels.Exists(TokenValue) = False Then
            Let ValueIsIndefinite = True
        Else
            Let ReturnValue = Variables(TokenValue)
        End If
        
    ElseIf TokenKind = TOKEN_VARIABLE Then
        'If the named value is not yet known, the expression cannot be evaluated _
         until it is defined later on
        If Variables.Exists(TokenValue) = False Then
            Let ValueIsIndefinite = True
        Else
            Let ReturnValue = Variables(TokenValue)
        End If
        
    'This is an easy one
    ElseIf TokenKind = TOKEN_NUMBER Then
        Let ReturnValue = TokenValue
        
    Else
        'ERROR: Not a valid operand. We won't produce an error message here as the _
         error might be exepected (i.e. an optional parameter). We fold upwards and _
         whomever knows best will deal with the error
        Debug.Print TokenLine & ", " & TokenCol
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
    Dim Instruction As OZ80_TOKEN: Let Instruction = TokenKind
    'Remember the line number of this since a label or property on a different line _
     will be a definition, not a parameter
    Dim Line As Long: Let Line = TokenLine
    'Skip over the Z80 instruction token now that we know it
    If NextToken() = False Then Stop
    
    'If there appears to be a parameter, skip over doing parameterless instructions
    If IsParameter() = True Then GoTo With1Parameter
    
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
    If IsParameter() = True Then GoTo With2Parameters
    
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

'CRCString : Calculate a unique hash for a string _
 ======================================================================================
Private Function CRCString(ByRef Data As String) As Long
    Dim Length As Long
    
    'Cast the string to an array by attaching its pointer to our dummy VB array header
    With CastStrSA1D
       Let .pvData = StrPtr(Data)
       Let .cElements = LenB(Data)
       Let Length = .cElements
    End With

    'Calculate the CRC by twiddling the bits
    
    Dim i As Long
    Dim iLookup As Byte

    Let CRCString = &HFFFFFFFF
    Let Length = Length - 1
    For i = 0 To Length
        Let iLookup = (CRCString And &HFF) Xor CastStrArr(i)
        Let CRCString = ((CRCString And &HFFFFFF00) \ &H100) And 16777215
        Let CRCString = CRCString Xor CRCTable(iLookup)
    Next

    Let CRCString = Not (CRCString)
End Function

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
    If TokenStream.IsRegister Then '...................................................
        If TokenKind = TOKEN_Z80_A Then Let Result.Kind = PARAM_A
        If TokenKind = TOKEN_Z80_B Then Let Result.Kind = PARAM_B
        If TokenKind = TOKEN_Z80_C Then Let Result.Kind = PARAM_C
        If TokenKind = TOKEN_Z80_D Then Let Result.Kind = PARAM_D
        If TokenKind = TOKEN_Z80_E Then Let Result.Kind = PARAM_E
        If TokenKind = TOKEN_Z80_H Then Let Result.Kind = PARAM_H
        If TokenKind = TOKEN_Z80_L Then Let Result.Kind = PARAM_L
        If TokenKind = TOKEN_Z80_I Then Let Result.Kind = PARAM_I
        If TokenKind = TOKEN_Z80_R Then Let Result.Kind = PARAM_R
        If TokenKind = TOKEN_Z80_AF Then Let Result.Kind = PARAM_AF
        If TokenKind = TOKEN_Z80_BC Then Let Result.Kind = PARAM_BC
        If TokenKind = TOKEN_Z80_DE Then Let Result.Kind = PARAM_DE
        If TokenKind = TOKEN_Z80_HL Then Let Result.Kind = PARAM_HL
        If TokenKind = TOKEN_Z80_IX Then Let Result.Kind = PARAM_IX
        If TokenKind = TOKEN_Z80_IY Then Let Result.Kind = PARAM_IY
        If TokenKind = TOKEN_Z80_SP Then Let Result.Kind = PARAM_SP
        'Now move beyond the register token
        Call NextToken
    
    'If the token is a memory reference, such as `(ix+$8)`, _
     then get the associated register and value
    ElseIf TokenKind = TOKEN_PARENOPEN Then '..........................................
        Let GetParameter = ScopeMemory( _
            Result.Kind, Result.Value, Result.ValueIndefinite _
        )
        'TODO: Error - the memory reference is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
        
    'Anything else classes as an expression (number / labels / variables / operators)
    ElseIf TokenStream.IsExpression Then '.............................................
        Let Result.Kind = PARAM_VAL
        Let GetParameter = ScopeExpression(Result.Value, Result.ValueIndefinite)
        'TODO: Error - the expression is invalid
        If GetParameter <> OZ80_ERROR_NONE Then Stop
    
    Else
        'ERROR: Not an expression
        Let GetParameter = RaiseError(OZ80_ERROR_EXPRESSION, _
            "Unexpected content at this point", _
            "Z80 instruction parameter expected. Not a valid parameter expression.", _
            TokenLine, TokenCol _
        )
    End If
End Function

'IsParameter : Check if the current token would be considered a Z80 parameter _
 ======================================================================================
Private Function IsParameter() As Boolean
    'If the stream ended, always return false
    If TokenStream.EOF = True Then Exit Function
    'A parameter can be a number, label/variable/property, register _
     or memory reference
    If TokenKind = TOKEN_LABEL Then
        'A label on a new line is a definition, not a parameter
        If TokenLine = PrevLine Then Let IsParameter = True
    ElseIf TokenKind = TOKEN_NUMBER Then
        Let IsParameter = True
    ElseIf TokenKind = TOKEN_PARENOPEN Then
        Let IsParameter = True
    ElseIf TokenKind = TOKEN_PROPERTY Then
        'A property on a new line is a definition, not a parameter
        If TokenLine = PrevLine Then Let IsParameter = True
    ElseIf TokenStream.IsRegister Then
        Let IsParameter = True
    ElseIf TokenKind = TOKEN_VARIABLE Then
        Let IsParameter = True
    End If
End Function

'NextToken : Move to the next token in the token stream _
 ======================================================================================
Private Function NextToken() As Boolean
    'This function is a mere wrapper simply because we need to track when a new line _
     occurs between tokens to diferentiate label definitions and parameters
    With TokenStream
        Let PrevLine = .Line
        'We return false if the stream comes to an end
        Let NextToken = .Forward()
        'Cache the token information as dereferencing the object every time is slow
        Let TokenKind = .Kind
        Let TokenValue = .Value
        Let TokenLine = .Line
        Let TokenCol = .Col
    End With
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
        If TokenStream Is Nothing Then
            Let TestPath = FSO.BuildPath(CurDir(), FilePath)
        Else
            'When another file has been parsed before, _
             we work relatively from that file's directory
            Let TestPath = FSO.GetFile(TokenStream.SourceFile).Path
            'TODO: Error?
            Let TestPath = FSO.BuildPath(TestPath, FilePath)
        End If
    End If
    
    '[2] Check file exists: _
     ----------------------------------------------------------------------------------
    'ERROR: File Not Found
    If FSO.FileExists(TestPath) = False Then Let ParseFile = RaiseError( _
        OZ80_ERROR_FILENOTFOUND, _
        "File not found", _
        "There is no file '" & TestPath & "'. " & _
            "Note that files are included relative to the parent file.", _
        TokenLine, TokenCol _
    ): GoTo Finish
    
    '[3] Parse the file into a TokenStream: _
     ----------------------------------------------------------------------------------
    'We'll use the CRC of the filepath string as its key/ID
    Dim CRC As Long
    Let CRC = CRCString(TestPath)
    
    'If the file has already been parsed, re-use it
    If Files.Exists(CRC) = True Then
        'Reference the existing TokenStream object
        Set TokenStream = Files.Item(CRC)
    Else
        'Create a tokeniser object to hold the machine representation of the text file; _
         the assembler doesn't work with the original text directly
        Set TokenStream = New oz80_TokenStream
        'Explode the source code file into tokens
        Let ParseFile = TokenStream.Tokenise(TestPath)
        'If the file failed to parse, an error code / message will already have been _
         raised via an event, we only need to exit
        If ParseFile <> OZ80_ERROR_NONE Then GoTo Finish
        'Add the TokenStream to the files collection
        Call Files.Add(CRC, TokenStream)
    End If
    
Finish:
    Set FSO = Nothing
    
    'If there were any error, an error event and message would already have been _
     raised, we only need to fold upwards and the parent will handle terminating the _
     assembling process early
End Function

'RewindTokenStream : Due to the cache we have to wrap the rewind function _
 ======================================================================================
Private Sub RewindTokenStream()
    With TokenStream
        Call .Rewind
        Let PrevLine = 0
        Let TokenKind = .Kind
        Let TokenValue = .Value
        Let TokenLine = .Line
        Let TokenCol = .Col
    End With
End Sub

'RiaseError : provide error details through an event _
 ======================================================================================
Private Function RaiseError( _
    ByRef Number As OZ80_ERROR, _
    ByRef Title As String, _
    ByRef Description As String, _
    Optional ByRef Line As Long = 0, Optional ByRef Col As Long = 0 _
) As OZ80_ERROR
    'Pass the error up to the controller
    RaiseEvent Error(Number, Description, Title, Line, Col)
    
    'We return the same error number we were given only to make the tail recursion _
     simpler, e.g. `let ScopeRoot = RaiseError(OZ80_ERROR_...`
    Let RaiseError = Number
End Function
