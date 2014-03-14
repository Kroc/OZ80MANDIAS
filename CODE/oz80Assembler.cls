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

'/// API //////////////////////////////////////////////////////////////////////////////

'<msdn.microsoft.com/en-us/library/windows/desktop/ms724408%28v=vs.85%29.aspx>
Private Declare Function kernel32_GetTickCount Lib "kernel32" Alias "GetTickCount" ( _
) As Long

'/// CLASS VARIABLES //////////////////////////////////////////////////////////////////

Private TokenIndex As Long
Private Tokens() As oz80Token
Private TokenCount As Long
Private Token As oz80Token

Private Type OZ80_NAMEITEM
    Name As String
    Defined As Boolean
End Type

Private Bank As Long
Private Slot As Long
Private At As Long

Private ROM() As Byte
Private ROMSize As Long

'This is used to provide the result to an expression evaluation; either it can be _
 calculated, in which case the result value is provided, otherwise it cannot yet be _
 calculated and the reason why is returned (undefined variable / label ...)
Private Type OZ80_EXPR
    Kind As OZ80_TOKEN
    Data As OZ80_TOKEN_DATA
    Value As Long
End Type


'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Erase ROM
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

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
    
    Do
        Call ContextRoot
    Loop While NextToken()
    
    Debug.Print "- Final ROM size is " & (ROMSize \ 1024) & " KB"
    Debug.Print "- Time: " & Format$((kernel32_GetTickCount() - Start) / 1000, "0.000") & "s"
End Function

'/// CONTEXT HANDLERS /////////////////////////////////////////////////////////////////

'<text>     "           {ASCII}     "
'
'<expr>     <str>
'           <number>    [<operator> <expr>]
'           #var        [<operator> <expr>]
'           : Label     [<operator> <expr>]
'           (           <expr>      )
'
'
'<list>     <expr>      [,  <list>]
'
'<root>     :label
'
'           FILL        BINARY      <text>
'                       <list>
'
'           INCLUDE     [BINARY]    <text>  [START <expr>]  [LENGTH <expr>]
'                                                           [STOP <expr>]
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

'ContextRoot : When we're not within the scope of anything else _
 ======================================================================================
Private Function ContextRoot() As OZ80_ERROR
    If Token.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: Define the label at the current assembly point
        Exit Function
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'Check allowed keywords at this context
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
    
    'The root context also allows for source code, which is a context shared by blocks
    Call ContextCode
    
End Function

'ContextCode : Process areas where source code is allowed _
 ======================================================================================
Private Function ContextCode() As OZ80_ERROR
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
            'TODO: Error - keyword not allowed at this context
'            Stop
        End If
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_Z80 Then
        'Process a Z80 instruction ...
        Let ContextCode = ContextZ80()
    
    Else
        'TODO: Error - unexpected content at this point
'        Stop
    End If
End Function

'ContextExpression: Evaluate a calculation into a value _
 ======================================================================================
Private Function ContextExpression(ByRef Result As OZ80_EXPR) As OZ80_ERROR
    Dim Operand1 As OZ80_EXPR
    Dim Operand2 As OZ80_EXPR
    Dim Operator As OZ80_TOKEN_DATA
    
    'Begin with an operand
    'TODO: Check for quote -- quotes are a lone expression and cannot be an operand
    Let ContextExpression = ContextOperand(Operand1)
    If ContextExpression <> OZ80_ERROR_NONE Then GoTo ErrorOut
    
    Let Result.Kind = Operand1.Kind
    Let Result.Data = Operand1.Data
    Let Result.Value = Operand1.Value
    
    'TODO: Somehow record tokens that need values filling in later (such as labels)
    
    'Get the operator:
    If NextToken() = False Then GoTo OutOfTokens
    '(if not an operator assume the expression has finished)
    If Token.Kind <> OZ80_TOKEN.TOKEN_OPERATOR Then Exit Function
Continue:
    'Remember the operator for after we have operand 2 to work with
    Let Operator = Token.Data
    
    'Get the second operand
    If NextToken() = False Then GoTo OutOfTokens
    Let ContextExpression = ContextOperand(Operand2)
    'Cannot have an operator hanging there without an operand following
    If ContextExpression <> OZ80_ERROR_NONE Then GoTo ErrorOut
    
    'TODO: What if Operand1 is a register or an unknown label?
    
    'Now apply the operator
    If Operator = TOKEN_OPERATOR_ADD Then
        'TODO: Check for overflow beforehand
        Let Result.Value = Operand1.Value + Operand2.Value
        
    ElseIf Operator = TOKEN_OPERATOR_SUB Then
        'TODO: Check for overflow beforehand
        Let Result.Value = Operand1.Value - Operand2.Value
    
    ElseIf Operator = TOKEN_OPERATOR_MUL Then
        'TODO: Check for overflow beforehand
        Let Result.Value = Operand1.Value * Operand2.Value
    
    ElseIf Operator = TOKEN_OPERATOR_DIV Then
        'TODO: Check for divide by 0
        'TODO: Should we use Doubles so as to be accurate?
        Let Result.Value = Operand1.Value / Operand2.Value
    
    ElseIf Operator = TOKEN_OPERATOR_POW Then
        'TODO: Check for overflow
        Let Result.Value = Operand1.Value ^ Operand2.Value
    
    ElseIf Operator = TOKEN_OPERATOR_MOD Then
        'TODO
        Let Result.Value = Operand1.Value Mod Operand2.Value
    
    ElseIf Operator = TOKEN_OPERATOR_AND Then
        'TODO
    
    ElseIf Operator = TOKEN_OPERATOR_OR Then
        'TODO
    
    ElseIf Operator = TOKEN_OPERATOR_REP Then
        'TODO: Need to evaluate if this is the right place for this
    
    End If
    
    'Check if the expression continues...
    If NextToken() = False Then GoTo OutOfTokens
    If Token.Kind = OZ80_TOKEN.TOKEN_OPERATOR Then
        'Set the current result as the first operand and go around again
        Let Operand1.Value = Result.Value
        GoTo Continue
    End If
    
    'At this point the expression is complete, but we need to inform the parent if _
     it's incomplete due to yet-unknown values (such as labels). The parent can _
     reserve the space in the ROM and fill in the correct value later on
    
    Exit Function
    
ErrorOut:
OutOfTokens:
    Stop
End Function

'ContextOperand : Get the value of one half of a calculation _
 ======================================================================================
Private Function ContextOperand(ByRef Operand As OZ80_EXPR) As OZ80_ERROR
    If Token.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: If the label is not yet known, the expression cannot be evaluated _
         until the label is reached
        Let Operand.Kind = TOKEN_LABEL
         
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_NUMBER Then
        'This is an easy one
        Let Operand.Kind = TOKEN_NUMBER
        Let Operand.Value = Token.Data
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_PROPERTY Then
        'If a property, check the scope to see which variable / label it belongs to
        'TODO
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
        'An expression could contain a register, such as with `ld a, (ix+$8)`. _
         When this happens, set the type of the operand to "register" and leave it _
         to `ContextExpression` to manage it
        Let Operand.Kind = TOKEN_REGISTER
        Let Operand.Data = Token.Data
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_VARIABLE Then
        'TODO: If the variable is not yet defined, error
        
    Else
        'ERROR: Not a valid operand
    End If
    Exit Function

ErrorOut:
End Function

'ContextParentheses : Calculate the contents of parentheses _
 ======================================================================================
Private Function ContextParentheses(ByRef Result As OZ80_EXPR) As OZ80_ERROR
    'It's assumed you've already checked the current token is an open parentheses _
     before calling this function, so move forward into the expression
    If NextToken() = False Then GoTo OutOfTokens
    
    'Now process the expression; when it ends it will fold back here for us to check _
     for the closing parentheses
    If ContextExpression(Result) <> OZ80_ERROR_NONE Then
        GoTo ErrorOut
    End If
    
    'After the expression, there should be a closing parentheses
    If NextToken() = False Then GoTo OutOfTokens
    If Token.Kind <> OZ80_TOKEN.TOKEN_PARENCLOSE Then
        'ERROR: Expression did not terminate correctly
        'TODO: How do we handle lists within parentheses?
        GoTo ErrorOut
    End If
    
    Exit Function

ErrorOut:
OutOfTokens:
End Function

'ContextZ80 : Process a Z80 instruction _
 ======================================================================================
Private Function ContextZ80() As OZ80_ERROR
    'ADC : Add with Carry _
     ----------------------------------------------------------------------------------
    If Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_ADC Then
        'The instruction has been read, move on to the parameters
        If NextToken() = False Then GoTo OutOfTokens
        'In all instances, the first parameter *MUST* be register A or HL
        'TODO: Support assumed A, i.e. `adc $8`?
        If Token.Kind <> OZ80_TOKEN.TOKEN_REGISTER Then
            'ERROR: First parameter to ADC must be a register
            Let ContextZ80 = OZ80_ERROR_Z80_PARAMETER
            Call RaiseError(ContextZ80, _
                "Z80 instruction `ADC` requires the first parameter to be either " & _
                "the `A` register or the `HL` register.", _
                Token.Line, Token.Col _
            ): Exit Function
        End If
        
        'The 16-bit parameters are the simplest to handle: _
         ADC hl, bc|de|hl|sp ..........................................................
        If Token.Data = TOKEN_REGISTER_HL Then
            'Move to the second parameter
            If NextToken() = False Then GoTo OutOfTokens
            
            If Token.Kind <> OZ80_TOKEN.TOKEN_REGISTER Then
                'ERROR: With HL as first, second parameter must be a register
                Let ContextZ80 = OZ80_ERROR_Z80_PARAMETER
                Call RaiseError(ContextZ80, _
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
                Let ContextZ80 = OZ80_ERROR_Z80_PARAMETER
                Call RaiseError(ContextZ80, _
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
         ADC a, a|b|c|d|e|h|l .........................................................
        If Token.Kind = OZ80_TOKEN.TOKEN_REGISTER Then
            If Token.Data = TOKEN_REGISTER_A Then
                Call AddByte(&H88 + 7)  'Add opcodes for `ADC a, a`
            ElseIf Token.Data = TOKEN_REGISTER_B Then
                Call AddByte(&H88)      'Add opcodes for `ADC a, b`
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
        End If
        
        'Try to evaluate the next parameter
        Dim Expr As OZ80_EXPR
        Call ContextExpression(Expr)
        
        'ADC a, (hl) ..................................................................
        'ADC a, (ix+$8|iy+$8) .........................................................
        'ADC a, $8 ....................................................................
        
    
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
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPD Then '................. `cpd` ...
        'No parameters, add opcodes $ED $A9
        Call AddByte(&HED): Call AddByte(&HA9)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPDR Then '............... `cpdr` ...
        'No parameters, add opcodes $ED $B9
        Call AddByte(&HED): Call AddByte(&HB9)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPI Then '................. `cpi` ...
        'No parameters, add opcodes $ED $A1
        Call AddByte(&HED): Call AddByte(&HA1)
            
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPIR Then '............... `cpir` ...
        'No parameters, add opcodes $ED $B1
        Call AddByte(&HED): Call AddByte(&HB1)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPL Then '................. `cpl` ...
        'No parameters, add opcode $2F
        Call AddByte(&H2F)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DAA Then '................. `daa` ...
        'No parameters, add opcode $27
        Call AddByte(&H27)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DEC Then '................. `dec` ...
        'DEC a|b|c|d|e|h|l
        'DEC (hl|ix+$8|iy+$8)
        'DEC bc|de|hl|sp|ix|iy
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DI Then '................... `di` ...
        'No parameters, add opcode $F3
        Call AddByte(&HF3)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DJNZ Then '............... `djnz` ...
        'DJNZ $8
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EI Then '................... `ei` ...
        'No parameters, add opcode $FB
        Call AddByte(&HFB)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EX Then '................... `ex` ...
        'EX (sp), hl|ix|iy
        'EX af, af'
        'EX de, hl
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EXX Then '................. `exx` ...
        'No parameters, add opcode $D9
        Call AddByte(&HD9)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_HALT Then '............... `halt` ...
        'No parameters, add opcode $76
        Call AddByte(&H76)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IM Then '................... `im` ...
        'IM 0-2
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IN Then '................... `in` ...
        'IN a, ($8)
        'IN (c)
        'IN a|b|c|d|e|h|l, (c)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INC Then '................. `inc` ...
        'INC a|b|c|d|e|h|l
        'INC bc|de|hl|sp|ix|iy
        'INC (hl|ix+$8|iy+$8)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_IND Then '................. `ind` ...
        'No parameters, add opcodes $ED $AA
        Call AddByte(&HED): Call AddByte(&HAA)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INDR Then '............... `indr` ...
        'No parameters, add opcodes $ED $BA
        Call AddByte(&HED): Call AddByte(&HBA)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INI Then '................. `ini` ...
        'No parameters, add opcodes $ED $A2
        Call AddByte(&HED): Call AddByte(&HA2)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INIR Then '............... `inir` ...
        'No parameters, add opcodes $ED $B2
        Call AddByte(&HED): Call AddByte(&HB2)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_JP Then '................... `jp` ...
        'JP $16
        'JP (hl|ix|iy)
        'JP c|nc|m|p|z|nz|pe|po, $16
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_JR Then '................... `jr` ...
        'JR $8
        'JR c|nc|z|nz, $8
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LD Then '................... `ld` ...
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
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDD Then '................. `ldd` ...
        'No parameters, add opcodes $ED $A8
        Call AddByte(&HED): Call AddByte(&HA8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDDR Then '............... `lddr` ...
        'No parameters, add opcodes $ED $B8
        Call AddByte(&HED): Call AddByte(&HB8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDI Then '................. `ldi` ...
        'No parameters, add opcodes $ED $A0
        Call AddByte(&HED): Call AddByte(&HA0)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDIR Then '............... `ldir` ...
        'No parameters, add opcodes $ED $B0
        Call AddByte(&HED): Call AddByte(&HB0)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NEG Then '................. `neg` ...
        'No parameters, add opcodes $ED $44
        Call AddByte(&HED): Call AddByte(&H44)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NOP Then '................. `nop` ...
        'No parameters, add opcode $00
        Call AddByte(&H0)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OR Then '................... `or` ...
        'OR a|b|c|d|e|h|l|$8
        'OR (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUT Then '................. `out` ...
        'OUT ($8), a
        'OUT (c), 0|a|b|c|d|e|h|l
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTD Then '............... `outd` ...
        'No parameters, add opcodes $ED $AB
        Call AddByte(&HED): Call AddByte(&HAB)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTDR Then '............... `otdr` ...
        'No parameters, add opcodes $ED $BB
        Call AddByte(&HED): Call AddByte(&HBB)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTI Then '............... `outi` ...
        'No parameters, add opcodes $ED $A3
        Call AddByte(&HED): Call AddByte(&HA3)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTIR Then '............... `otir` ...
        'No parameters, add opcodes $ED $B3
        Call AddByte(&HED): Call AddByte(&HB3)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_POP Then '................. `pop` ...
        'POP af|bc|de|hl|ix|iy
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_PUSH Then '............... `push` ...
        'PUSH af|bc|de|hl|ix|iy
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RES Then '................. `res` ...
        'RES 0-7, a|b|c|d|e|h|l
        'RES 0-7, (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RET Then '................. `ret` ...
        'RET
        'RET c|nc|m|p|z|nz|pe|po
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RETI Then '............... `reti` ...
        'No parameters, add opcodes $ED $4D
        Call AddByte(&HED): Call AddByte(&H4D)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RETN Then '............... `retn` ...
        'No parameters, add opcodes $ED $45
        Call AddByte(&HED): Call AddByte(&H45)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLA Then '................. `rla` ...
        'No parameters, add opcode $17
        Call AddByte(&H17)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RL Then '................... `rl` ...
        'RL a|b|c|d|e|h|l
        'RL (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLCA Then '............... `rlca` ...
        'No parameters, add opcode $07
        Call AddByte(&H7)
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLC Then '................. `rlc` ...
        'RLC a|b|c|d|e|h|l
        'RLC (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLD Then '................. `rld` ...
        'No parameters, add opcodes $ED $6F
        Call AddByte(&HED): Call AddByte(&H6F)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRA Then '................. `rra` ...
        'No parameters, add opcode $1F
        Call AddByte(&H1F)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RR Then '................... `rr` ...
        'RR a|b|c|d|e|h|l
        'RR (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRCA Then '............... `rrca` ...
        'No parameters, add opcode $0F
        Call AddByte(&HF)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRC Then '................. `rrc` ...
        'RRC a|b|c|d|e|h|l
        'RRC (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRD Then '................. `rrd` ...
        'No parameters, add opcodes $ED $67
        Call AddByte(&HED): Call AddByte(&H67)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RST Then '................. `rst` ...
        'RST $00|$08|$10|$18|$20|$28|$30|$38
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SBC Then '................. `sbc` ...
        'SBC a|b|c|d|e|h|l
        'SBC a, $8
        'SBC (hl)
        'SBC a, (ix+$8|iy+$8)
        'SBC hl, bc|de|hl|sp
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SCF Then '................. `scf` ...
        'No parameters, add opcode $37
        Call AddByte(&H37)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SET Then '................. `set` ...
        'SET 0-7, a|b|c|d|e|h|l
        'SET 0-7, (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SLA Then '................. `sla` ...
        'SLA a|b|c|d|e|h|l
        'SLA (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SRA Then '................. `sra` ...
        'SRA a|b|c|d|e|h|l
        'SRA (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SLL Then '................. `sll` ...
        'SLL a|b|c|d|e|h|l
        'SLL (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SRL Then '................. `srl` ...
        'SRL a|b|c|d|e|h|l
        'SRL (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_SUB Then '................. `sub` ...
        'SUB a|b|c|d|e|h|l|$8
        'SUB (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_XOR Then '................. `xor` ...
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
