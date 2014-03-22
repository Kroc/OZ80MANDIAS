VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80Assembler"
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
'CLASS :: oz80Assembler

'This class will take a token stream and do the actual assembling into a Z80 binary

'/// API //////////////////////////////////////////////////////////////////////////////

'<msdn.microsoft.com/en-us/library/windows/desktop/ms724408%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetTickCount Lib "kernel32" Alias "GetTickCount" ( _
) As Long

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

Private TokenIndex As Long
Private Tokens() As oz80Token
Private TokenCount As Long
Private Token As oz80Token

Private Bank As Long
Private Slot As Long
Private At As Long

Private ROM() As Byte
Private ROMSize As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Erase ROM
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Process : Assemble a Z80 binary from a given token stream _
 ======================================================================================
Public Function Process(ByRef TokenStream As oz80Tokeniser) As OZ80_ERROR
    Debug.Print
    Debug.Print "* Assembling token stream"
    
    'Measure the time this takes
    Dim Start As Long
    Let Start = kernel32_GetTickCount()
    
    Let TokenIndex = 0
    Call TokenStream.GetTokenArray(Tokens)
    Let TokenCount = UBound(Tokens)
    
    'Set the assembly point to the beginning
    Let Bank = 0
    Let Slot = 0
    Let At = 0
    
    'Default to 32 KB ROM (the ROM size will double each time it overflows)
    Let ROMSize = 32768
    ReDim ROM(0 To ROMSize - 1) As Byte
    
    'Phase 1 : Determine the size of the assembly and fill in all label positions _
     ----------------------------------------------------------------------------------
    Do
        Call ScopeRoot
    Loop While NextToken()
    
    Debug.Print "- Final ROM size is " & (ROMSize \ 1024) & " KB"
    Debug.Print "- Time: " & Format$((kernel32_GetTickCount() - Start) / 1000, "0.000") & "s"
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
'           : Label     [{operator} <expr>]
'           {           <expr>      }
'
'
'<list>     <text>      [,          <list>]
'           <expr>      [,          <list>]
'
'<root>     :label
'
'           FILL        BINARY      <text>
'                       <list>
'
'           INCLUDE     <list>
'                       BINARY      <text>  [START <expr>]  [LENGTH <expr>]
'                                                           [STOP   <expr>]
'
'           TABLE       :label      <table>
'
'           OBJECT      #object     <object>
'
'           STRUCT      :label      AS      #object         <struct>
'
'           PROC        :label      [PARAMS register
'                                           number]
'                                   [RETURN register
'                                           number]
'                                   <block>
'
'           <code>
'
'<code>     DATA        <list>
'           BYTE        <list>
'           WORD        <list>
'
'           DEF         #variable   [<list>]
'
'           <asm>
'
'<block>    {           .label      [AS     #object]
'                       -
'                       +
'                       =
'                       <code>
'
'           }
'

'ScopeRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Function ScopeRoot() As OZ80_ERROR
    'When a file begins, it is at the top-most scope. Certain keywords are only _
     allowed here. Once all nested scopes are exhausted, we return here
    
    If Token.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: Define the label at the current assembly point
        Exit Function
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'Check allowed keywords at this Scope
        If Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_FILL Then
            'TODO
            Exit Function
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_INCLUDE Then
            'TODO
            Exit Function
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_TABLE Then
            'TODO
            Exit Function
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_STRUCT Then
            'TODO
            Exit Function
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_PROC Then
            'TODO
            Exit Function
            
        End If
    End If
    
    'The root Scope also allows for source code, which is a Scope shared by blocks
    Call ScopeCode
    
End Function

'ScopeCode : Process areas where source code is allowed _
 ======================================================================================
Private Function ScopeCode() As OZ80_ERROR
    'This scope allows data statements and Z80 source code, e.g. root and procedures
    
    If Token.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'TODO: Which keywords are allowed here
        If Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_BYTE Then
            'TODO
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DATA Then
            'TODO
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DEF Then
            'TODO
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_WORD Then
            'TODO
        Else
            'TODO: Error - keyword not allowed at this Scope
'            Stop
        End If
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_Z80 Then
        'Process a Z80 instruction ...
        Let ScopeCode = ScopeZ80()
    
    Else
        'TODO: Error - unexpected content at this point
'        Stop
    End If
End Function

'ScopeExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ScopeExpression( _
    ByRef Result As Long, ByRef Unknown As Boolean _
) As OZ80_ERROR
    Dim Operand1 As Long                'First value in an expression
    Dim Operand2 As Long                'Next value in an expression
    Dim Operator As OZ80_TOKEN_DATA     'Operator to use with the expressions
    
    Let Unknown = False
    
    'Begin with fetching the first operand
    Let ScopeExpression = ScopeOperand(Operand1, Unknown)
    If ScopeExpression <> OZ80_ERROR_NONE Then GoTo ErrorOut
    Let Result = Operand1
    
    'The file could legitimately end with an expression, so don't error out if the _
     token stream comes to an end since the expression is valid. The parent scopes _
     will handle an unexpected end error
    If NextToken() = False Then Exit Function
    'If next token is not an operator assume the expression has finished, _
      the parent will handle whatever the next token is used for
    If Token.Kind <> OZ80_TOKEN.TOKEN_OPERATOR Then Exit Function
    
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = Token.Data
    
    'Get the second operand
    If NextToken() = False Then
        'TODO: Error - unexpected end
        Stop
    End If
    Let ScopeExpression = ScopeOperand(Operand2, Unknown)
    'Cannot have an operator hanging there without an operand following
    If ScopeExpression <> OZ80_ERROR_NONE Then
        'TODO: Error - invalid use of operator
        Stop
    End If
    
    'If either operand were a yet-unknowable value (such as a label or property), _
     then don't bother applying the operator, but do continue the expression so that _
     we may reach the end and continue processing the tokens that follow
    If Unknown = True Then
        'Do nothing here, just fall out of the if block, skipping operator processing
    
    ElseIf Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let Result = Operand1 + Operand2
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let Result = Operand1 - Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let Result = Operand1 * Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        'TODO: Should we use Doubles so as to be accurate?
        Let Result = Operand1 / Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let Result = Operand1 ^ Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        'TODO
        Let Result = Operand1 Mod Operand2
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        'TODO
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        'TODO
    
    ElseIf Operator = TOKEN_OPERATOR_REP Then
        'TODO: Need to evaluate if this is the right place for this
    
    End If
    
    'The file could legitimately end with an expression, so don't error out if the _
     token stream comes to an end since the expression is valid. The parent scopes _
     will handle an unexpected end error
    If NextToken() = False Then Exit Function
    
    'Does the expression continue?
    If Token.Kind = OZ80_TOKEN.TOKEN_OPERATOR Then
        'Set the current result as the first operand and go around again
        Let Operand1 = Result
        GoTo Continue
    End If

ErrorOut:
End Function

'ScopeMemory : Work out a memory reference, i.e. `($????)` / `(ix+$??)` _
 ======================================================================================
Private Function ScopeMemory( _
    ByRef Register As OZ80_TOKEN_DATA, _
    ByRef Value As Long, ByRef Unknown As Boolean _
) As OZ80_ERROR
    'It's assumed you've already checked the current token is an open parentheses _
     before calling this function, so move forward into the expression
    If NextToken() = False Then GoTo OutOfTokens
    
    'The first item in the expression can be a register, typically HL, IX or IY
    If Token.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
        'The `IN` & `OUT` instructions accept only the C register
        If Token.Data = TOKEN_REGISTER_C Then
            Let Register = TOKEN_REGISTER_C
        
        '`LD` accepts registers BC & DE
        ElseIf Token.Data = TOKEN_REGISTER_BC Then
            Let Register = TOKEN_REGISTER_BC
        ElseIf Token.Data = TOKEN_REGISTER_DE Then
            Let Register = TOKEN_REGISTER_DE
            
        ElseIf Token.Data = TOKEN_REGISTER_HL Then
            'TODO
            Let Register = TOKEN_REGISTER_HL
            
        ElseIf Token.Data = TOKEN_REGISTER_IX Then
            Let Register = TOKEN_REGISTER_IX
            'IX can have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf Token.Data = TOKEN_REGISTER_IY Then
            Let Register = TOKEN_REGISTER_IY
            'IY can also have an additional expression, e.g. `ld a, (ix+$8)`
            GoTo AddExpr
            
        ElseIf Token.Data = TOKEN_REGISTER_SP Then
            Let Register = TOKEN_REGISTER_SP
        Else
            'ERROR: register not allowed here
            Stop
        End If
        'Move to the next token, then check that it's the end of the memory reference
        'TODO: out of tokens
        If NextToken() = False Then Stop
        GoTo EndScope
    End If

AddExpr:
    'The expression must begin with `+`
    If Token.Kind = OZ80_TOKEN.TOKEN_OPERATOR & Token.Data = TOKEN_OPERATOR_ADD Then
        'Try and calculate the expression. It might come back unknown and this will be _
         passed on to the parent to decide if that's a problem or not
        Let ScopeMemory = ScopeExpression(Value, Unknown)
    End If
    
EndScope:
    If Token.Kind = OZ80_TOKEN.TOKEN_PARENCLOSE Then Exit Function
    'ERROR: Unexpected token in memory expression
    'TODO
    Exit Function
    
OutOfTokens:
    'TODO
End Function

'ScopeOperand : Get the value of one half of a calculation _
 ======================================================================================
Private Function ScopeOperand( _
    ByRef Operand As Long, ByRef Unknown As Boolean _
) As OZ80_ERROR
    If Token.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: If the label is not yet known, the expression cannot be evaluated _
         until the label is reached
        Let Unknown = True
         
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_NUMBER Then
        'This is an easy one
        Let Operand = Token.Data
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_PROPERTY Then
        'If a property, check the scope to see which variable / label it belongs to
        'TODO
        Let Unknown = True
                
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_VARIABLE Then
        'TODO: If the variable is not yet defined, error
        Let Unknown = True
        
    Else
        'ERROR: Not a valid operand
        Debug.Print Token.Line & ", " & Token.Col
        Stop
    End If
End Function

''ScopeParentheses : Calculate the contents of parentheses _
' ======================================================================================
'Private Function ScopeParentheses(ByRef Result As OZ80_EXPR) As OZ80_ERROR
'    'It's assumed you've already checked the current token is an open parentheses _
'     before calling this function, so move forward into the expression
'    If NextToken() = False Then GoTo OutOfTokens
'
'    'Now process the expression; when it ends it will fold back here for us to check _
'     for the closing parentheses
'    If ScopeExpression(Result) <> OZ80_ERROR_NONE Then
'        GoTo ErrorOut
'    End If
'
'    'After the expression, there should be a closing parentheses
'    If NextToken() = False Then GoTo OutOfTokens
'    If Token.Kind <> OZ80_TOKEN.TOKEN_PARENCLOSE Then
'        'ERROR: Expression did not terminate correctly
'        'TODO: How do we handle lists within parentheses?
'        GoTo ErrorOut
'    End If
'
'    Exit Function
'
'ErrorOut:
'OutOfTokens:
'End Function

'ScopeZ80 : Process a Z80 instruction _
 ======================================================================================
Private Function ScopeZ80() As OZ80_ERROR
    'Parameter expressions might return an unknown value (such as when a label, which _
     has not been reached yet, is used), which we may have to handle differently
    Dim Value As Long, Unknown As Boolean
    
    'ADC : Add with Carry _
     ----------------------------------------------------------------------------------
    If Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_ADC Then
        If Token.Line = 7835 Then Stop
        
        'The instruction has been read, move on to the parameters
        If NextToken() = False Then GoTo OutOfTokens
        'In all instances, the first parameter *MUST* be register A or HL
        'TODO: Support assumed A, i.e. `adc $8`?
        If Token.Kind <> OZ80_TOKEN.TOKEN_REGISTER Then
            'ERROR: First parameter to ADC must be a register
            Let ScopeZ80 = OZ80_ERROR_Z80_PARAMETER
            Call RaiseError(ScopeZ80, _
                "Z80 instruction `ADC` requires the first parameter to be either " & _
                "the `A` register or the `HL` register.", _
                Token.Line, Token.Col _
            ): Exit Function
        End If
        
        'The 16-bit parameters are the simplest to handle: _
         `ADC hl, bc|de|hl|sp` ........................................................
        If Token.Data = TOKEN_REGISTER_HL Then
            'Move to the second parameter
            If NextToken() = False Then GoTo OutOfTokens
            
            If Token.Kind <> OZ80_TOKEN.TOKEN_REGISTER Then
                'ERROR: With HL as first, second parameter must be a register
                Let ScopeZ80 = OZ80_ERROR_Z80_PARAMETER
                Call RaiseError(ScopeZ80, _
                    "Z80 instruction `ADC` expects the second parameter to be a " & _
                    "register when the first parameter is the `HL` register.", _
                    Token.Line, Token.Col _
                ): Exit Function
            End If
            
            If Token.Data = TOKEN_REGISTER_BC Then
                'Add opcodes for `ADC hl, bc`
                Call AddByte(&HED): Call AddByte(&H4A)
            ElseIf Token.Data = TOKEN_REGISTER_DE Then
                'Add opcodes for `ADC hl, de`
                Call AddByte(&HED): Call AddByte(&H5A)
            ElseIf Token.Data = TOKEN_REGISTER_HL Then
                'Add opcodes for `ADC hl, hl`
                Call AddByte(&HED): Call AddByte(&H6A)
            ElseIf Token.Data = TOKEN_REGISTER_SP Then
                'Add opcodes for `ADC hl, sp`
                Call AddByte(&HED): Call AddByte(&H7A)
            Else
                'ERROR: Invalid register for `ADC hl, ??`
                Let ScopeZ80 = OZ80_ERROR_Z80_PARAMETER
                Call RaiseError(ScopeZ80, _
                    "When the first parameter of the `ADC` Z80 instruction is the " & _
                    "`HL` register, the second parameter may only be one of the " & _
                    "following registers: `BC`, `DE`, `HL` or `SP`.", _
                    Token.Line, Token.Col _
                )
            End If
            Exit Function
        End If
        
        'If not HL then the first parameter *must* be A
        If Token.Data <> TOKEN_REGISTER_A Then
            'ERROR: First parameter to ADC must be register A or HL
            Stop
        End If
        
        'Move to the second parameter
        If NextToken() = False Then GoTo OutOfTokens
        
        'If the second parameter is an 8-bit register: _
         `ADC a, a|b|c|d|e|h|l` .......................................................
        If Token.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
            If Token.Data = TOKEN_REGISTER_A Then
                Call AddByte(&H88 + 7)  'Add opcodes for `ADC a, a`
            ElseIf Token.Data = TOKEN_REGISTER_B Then
                Call AddByte(&H88 + 0)  'Add opcodes for `ADC a, b`
            ElseIf Token.Data = TOKEN_REGISTER_C Then
                Call AddByte(&H88 + 1)  'Add opcodes for `ADC a, c`
            ElseIf Token.Data = TOKEN_REGISTER_D Then
                Call AddByte(&H88 + 2)  'Add opcodes for `ADC a, d`
            ElseIf Token.Data = TOKEN_REGISTER_E Then
                Call AddByte(&H88 + 3)  'Add opcodes for `ADC a, e`
            ElseIf Token.Data = TOKEN_REGISTER_H Then
                Call AddByte(&H88 + 4)  'Add opcodes for `ADC a, h`
            ElseIf Token.Data = TOKEN_REGISTER_L Then
                Call AddByte(&H88 + 5)  'Add opcodes for `ADC a, l`
            Else
                'ERROR: If second parameter is a register it must be A/B/C/D/E/H or L
                Stop
            End If
            Exit Function
        End If
        
        'Check if there's a memory reference
        If Token.Kind = OZ80_TOKEN.TOKEN_PARENOPEN Then
            Dim Register As OZ80_TOKEN_DATA
            Let ScopeZ80 = ScopeMemory(Register, Value, Unknown)
            If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
            
            '`ADC a, (hl)` ............................................................
            If Register = TOKEN_REGISTER_HL Then
                'Add opcodes for `ADC a, (hl)`
                Call AddByte(&H8E)
                
            'ADC a, (ix+$8|iy+$8) .....................................................
            ElseIf Register = TOKEN_REGISTER_IX Or Register = TOKEN_REGISTER_IY Then
                'TODO: Check the value is not out of 8-bit range
                'Add opcodes for `ADC a, (ix+$8|iy+$8)`
                If Register = TOKEN_REGISTER_IX Then
                    Call AddByte(&HDD)
                Else
                    Call AddByte(&HFD)
                End If
                Call AddByte(&H8E)
                'TODO: If value was invalid, make sure to come back to it later
                Call AddByte(CByte(Value))
                
            Else
                'ERROR: Invalid memory reference for this instruction
                Stop
            End If
            Exit Function
        End If
        
        'Finally check for a numerical expression _
         `ADC a, $8` ..................................................................
        Let ScopeZ80 = ScopeExpression(Value, Unknown)
        If ScopeZ80 <> OZ80_ERROR_NONE Then Stop
        'TODO: Check the value is not out of 8-bit range
        'Add the opcodes for `ADC a, $8`
        Call AddByte(&HCE)
        'TODO: If value was invalid, make sure to come back to it later
        Call AddByte(CByte(Value))
    
    'ADD _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_ADD Then
        'ADD a, a|b|c|d|e|h|l
        'ADD a, $8
        'ADD a, (hl|ix+$8|iy+$8)
        'ADD hl, bc|de|hl|sp
        'ADD ix, bc|de|ix|sp
        'ADD iy, bc|de|iy|sp
    
    'AND : Bitwise and _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_AND Then
        'AND a|b|c|d|e|h|l
        'AND $8
        'AND (hl|ix+$8|iy+$8)
    
    'BIT : Bit comparison _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_BIT Then
        'BIT 0-7, a|b|c|d|e|h|l
        'BIT 0-7, (hl|ix+$8|iy+$8)
    
    'CALL : Subroutine call _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CALL Then
        'CALL $16
        'CALL c|nc|m|p|z|nz|pe|po, $16
    
    'CCF : Clear the Carry Flag _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CCF Then
        'No parameters, add opcode $3F
        Call AddByte(&H3F)
    
    'CP : Compare _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CP Then
        'CP a|b|c|d|e|h|l
        'CP $8
        'CP (hl|ix+$8|iy+$8)
    
    'CPD : Compare & Decrement _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPD Then
        'No parameters, add opcodes $ED $A9
        Call AddByte(&HED): Call AddByte(&HA9)
    
    'CPDR : Compare, Decrement & Repeat _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPDR Then
        'No parameters, add opcodes $ED $B9
        Call AddByte(&HED): Call AddByte(&HB9)
    
    'CPI : Compare & Increment _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPI Then
        'No parameters, add opcodes $ED $A1
        Call AddByte(&HED): Call AddByte(&HA1)
            
    'CPIR : Compare, Increment & Repeat _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPIR Then
        'No parameters, add opcodes $ED $B1
        Call AddByte(&HED): Call AddByte(&HB1)
    
    'CPL : "Complement" -- 8-bit invert _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPL Then
        'No parameters, add opcode $2F
        Call AddByte(&H2F)
    
    'DAA : Decimal Adjust Accumulator (decimal add) _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DAA Then
        'No parameters, add opcode $27
        Call AddByte(&H27)
    
    'DEC : Decrement _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DEC Then
        'DEC a|b|c|d|e|h|l
        'DEC (hl|ix+$8|iy+$8)
        'DEC bc|de|hl|sp|ix|iy
    
    'DI : Disable Interrupts _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DI Then
        'No parameters, add opcode $F3
        Call AddByte(&HF3)
        
    'DJNZ : Decrement, Jump if Not Zero _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DJNZ Then
        'DJNZ $8
    
    'EI : Enable Interrupts _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EI Then
        'No parameters, add opcode $FB
        Call AddByte(&HFB)
    
    'EX : Exchange registers _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EX Then
        'EX (sp), hl|ix|iy
        'EX af, af'
        'EX de, hl
    
    'EXX : Swap the registers with their shadow values _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EXX Then
        'No parameters, add opcode $D9
        Call AddByte(&HD9)
    
    'HALT : Stop the processor until the next interrupt _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_HALT Then
        'No parameters, add opcode $76
        Call AddByte(&H76)
        
    'IM : Interrupt Mode _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IM Then
        'IM 0-2
    
    'IN : Input from port _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IN Then
        'IN a, ($8)
        'IN (c)
        'IN a|b|c|d|e|h|l, (c)
    
    'INC : Increment _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INC Then
        'INC a|b|c|d|e|h|l
        'INC bc|de|hl|sp|ix|iy
        'INC (hl|ix+$8|iy+$8)
        
    'IND : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IND Then
        'No parameters, add opcodes $ED $AA
        Call AddByte(&HED): Call AddByte(&HAA)
    
    'INDR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INDR Then
        'No parameters, add opcodes $ED $BA
        Call AddByte(&HED): Call AddByte(&HBA)
    
    'INI : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INI Then
        'No parameters, add opcodes $ED $A2
        Call AddByte(&HED): Call AddByte(&HA2)
    
    'INIR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INIR Then
        'No parameters, add opcodes $ED $B2
        Call AddByte(&HED): Call AddByte(&HB2)
        
    'JP : Jump to absolute address (within memory space) _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_JP Then
        'JP $16
        'JP (hl|ix|iy)
        'JP c|nc|m|p|z|nz|pe|po, $16
    
    'JR : Jump relatively, forwards or backwards _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_JR Then
        'JR $8
        'JR c|nc|z|nz, $8
    
    'LD : Load registers & values _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LD Then
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
    
    'LDD : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDD Then
        'No parameters, add opcodes $ED $A8
        Call AddByte(&HED): Call AddByte(&HA8)
    
    'LDDR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDDR Then
        'No parameters, add opcodes $ED $B8
        Call AddByte(&HED): Call AddByte(&HB8)
    
    'LDI : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDI Then
        'No parameters, add opcodes $ED $A0
        Call AddByte(&HED): Call AddByte(&HA0)
        
    'LDIR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDIR Then
        'No parameters, add opcodes $ED $B0
        Call AddByte(&HED): Call AddByte(&HB0)
    
    'NEG : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NEG Then
        'No parameters, add opcodes $ED $44
        Call AddByte(&HED): Call AddByte(&H44)
    
    'NOP : No Operation _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NOP Then
        'No parameters, add opcode $00
        Call AddByte(&H0)
    
    'OR : Bitwise Or _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OR Then
        'OR a|b|c|d|e|h|l|$8
        'OR (hl|ix+$8|iy+$8)
    
    'OUT : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUT Then
        'OUT ($8), a
        'OUT (c), 0|a|b|c|d|e|h|l
    
    'OUTD : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTD Then
        'No parameters, add opcodes $ED $AB
        Call AddByte(&HED): Call AddByte(&HAB)
    
    'OTDR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTDR Then
        'No parameters, add opcodes $ED $BB
        Call AddByte(&HED): Call AddByte(&HBB)
    
    'OUTI : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTI Then
        'No parameters, add opcodes $ED $A3
        Call AddByte(&HED): Call AddByte(&HA3)
    
    'OTIR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTIR Then
        'No parameters, add opcodes $ED $B3
        Call AddByte(&HED): Call AddByte(&HB3)
        
    'POP : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_POP Then
        'POP af|bc|de|hl|ix|iy
    
    'PUSH : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_PUSH Then
        'PUSH af|bc|de|hl|ix|iy
    
    'RES : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RES Then
        'RES 0-7, a|b|c|d|e|h|l
        'RES 0-7, (hl|ix+$8|iy+$8)
    
    'RET : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RET Then
        'RET
        'RET c|nc|m|p|z|nz|pe|po
    
    'RETI : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RETI Then
        'No parameters, add opcodes $ED $4D
        Call AddByte(&HED): Call AddByte(&H4D)
    
    'RETN : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RETN Then
        'No parameters, add opcodes $ED $45
        Call AddByte(&HED): Call AddByte(&H45)
    
    'RLA : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLA Then
        'No parameters, add opcode $17
        Call AddByte(&H17)
    
    'RL : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RL Then
        'RL a|b|c|d|e|h|l
        'RL (hl|ix+$8|iy+$8)
    
    'RLCA : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLCA Then
        'No parameters, add opcode $07
        Call AddByte(&H7)
        
    'RLC : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLC Then
        'RLC a|b|c|d|e|h|l
        'RLC (hl|ix+$8|iy+$8)
    
    'RLD : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLD Then
        'No parameters, add opcodes $ED $6F
        Call AddByte(&HED): Call AddByte(&H6F)
    
    'RRA : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRA Then
        'No parameters, add opcode $1F
        Call AddByte(&H1F)
    
    'RR : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RR Then
        'RR a|b|c|d|e|h|l
        'RR (hl|ix+$8|iy+$8)
    
    'RRCA : Rotate Accumulator Right (Circular) _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRCA Then
        'No parameters, add opcode $0F
        Call AddByte(&HF)
    
    'RRC : Rotate Right (Circular) _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRC Then
        'RRC a|b|c|d|e|h|l
        'RRC (hl|ix+$8|iy+$8)
    
    'RRD : 4-bit Rotate Right _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRD Then
        'No parameters, add opcodes $ED $67
        Call AddByte(&HED): Call AddByte(&H67)
    
    'RST : "Reset" - Jump into page 0 _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RST Then
        'RST $00|$08|$10|$18|$20|$28|$30|$38
    
    'SBC : Subtract with Carry _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SBC Then
        'SBC a|b|c|d|e|h|l
        'SBC a, $8
        'SBC (hl)
        'SBC a, (ix+$8|iy+$8)
        'SBC hl, bc|de|hl|sp
    
    'SCF : Set the Carry Flag _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SCF Then
        'No parameters, add opcode $37
        Call AddByte(&H37)
    
    'SET : Set individual bit _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SET Then
        'SET 0-7, a|b|c|d|e|h|l
        'SET 0-7, (hl|ix+$8|iy+$8)
    
    'SLA : _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SLA Then
        'SLA a|b|c|d|e|h|l
        'SLA (hl|ix+$8|iy+$8)
    
    'SRA : Arithmetic Shift Right _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SRA Then
        'SRA a|b|c|d|e|h|l
        'SRA (hl|ix+$8|iy+$8)
    
    'SLL : Logical Shift Left (undocumented opcode) _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SLL Then
        'SLL a|b|c|d|e|h|l
        'SLL (hl|ix+$8|iy+$8)
    
    'SRL : Logical Shift Right _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SRL Then
        'SRL a|b|c|d|e|h|l
        'SRL (hl|ix+$8|iy+$8)
    
    'SUB : Subtract _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SUB Then
        'SUB a|b|c|d|e|h|l|$8
        'SUB (hl|ix+$8|iy+$8)
    
    'XOR : Exclusive bitwise Or _
     ----------------------------------------------------------------------------------
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_XOR Then
        'XOR a|b|c|d|e|h|l|$8
        'XOR (hl|ix+$8|iy+$8)
    
    Else
        'ERROR: Unrecognised Z80 instruction
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
    Let ROM(At) = Value
    Let At = At + 1
    'Handle ROM overflowing
    If At > ROMSize Then
        Let ROMSize = ROMSize * 2
        ReDim Preserve ROM(0 To ROMSize) As Byte
        Debug.Print "! ROM size was increased to " & (ROMSize \ 1024) & " KB"
    End If
End Sub

'NextToken : Move to the next token in the token stream _
 ======================================================================================
Private Function NextToken() As Boolean
    Let TokenIndex = TokenIndex + 1
    If TokenIndex > TokenCount Then Exit Function
    Let Token = Tokens(TokenIndex)
    Let NextToken = True
End Function

'RiaseError : provide error details through an event _
 ======================================================================================
Private Sub RaiseError( _
    ByRef Number As OZ80_ERROR, _
    ByRef Description As String, _
    Optional ByRef Line As Long = 0, Optional ByRef Col As Long = 0 _
)
    Debug.Print
    Debug.Print "! ERROR: #" & Number
    If Line > 0 And Col > 0 Then
        Debug.Print "- Line: " & Format$(Line, "#,#") & " Col: " & Col
    End If
    Debug.Print "- " & Description
    Debug.Print
    
    Call Err.Clear
    Let Err.Number = vbObjectError + Number
    Let Err.Description = Description
    Let Err.Source = "Line #" & Line & " Col #" & Col
    
'    RaiseEvent Error(Err)
End Sub
