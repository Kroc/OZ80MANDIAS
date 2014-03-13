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

Private TokenIndex As Long
Private Tokens() As oz80Token
Private TokenCount As Long
Private Token As oz80Token

Private Bank As Long
Private Slot As Long
Private At As Long

Private ROM() As Byte

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Erase ROM
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Public Function Process(ByRef TokenStream As oz80Tokeniser) As OZ80_ERROR
    Let TokenIndex = 0
    Call TokenStream.GetTokenArray(Tokens)
    Let TokenCount = UBound(Tokens)
    
    Debug.Print
    Debug.Print "* Assembling token stream"
    
    Call ContextRoot
    
    Debug.Print "- Complete"
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
Private Sub ContextRoot()
    If Token.Kind = OZ80_TOKEN.TOKEN_LABEL Then
        'TODO: Define the label at the current assembly point
        Exit Sub
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'Check allowed keywords at this context
        If Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_FILL Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_INCLUDE Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_TABLE Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_STRUCT Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_PROC Then
            'TODO
            Exit Sub
            
        End If
    End If
    
    'The root context also allows for source code, which is a context shared by blocks
    Call ContextCode
End Sub

'ContextCode : Process areas where source code is allowed _
 ======================================================================================
Private Sub ContextCode()
    If Token.Kind = OZ80_TOKEN.TOKEN_KEYWORD Then
        'TODO: Which keywords are allowed here
        If Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_BYTE Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DATA Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_DEF Then
            'TODO
            Exit Sub
            
        ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_KEYWORD_WORD Then
            'TODO
            Exit Sub
        Else
            'TODO: Error - keyword not allowed at this context
            Stop
        End If
        
    ElseIf Token.Kind = OZ80_TOKEN.TOKEN_Z80 Then
        'Process a Z80 instruction ...
        Call ContextZ80
        Exit Sub
        
    End If
    
    'TODO: Error - unexpected content at this point
'    Stop
End Sub

'ContextZ80 : Process a Z80 instruction _
 ======================================================================================
Private Sub ContextZ80()
    If Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_ADC Then '..................... `adc` ...
        'The instruction has been read, move on to the parameters
        If NextToken() = False Then GoTo OutOfTokens
        'In all instances, the first parameter *MUST* be register A or HL
        'TODO: Support assumed A, i.e. `adc $8`
        If Token.Kind = OZ80_TOKEN.TOKEN_REGISTER _
        And (Token.Data = TOKEN_REGISTER_A Or Token.Data = TOKEN_REGISTER_HL) Then
            'Move to the second parameter and fall through to further processing
            If NextToken() = False Then GoTo OutOfTokens
        Else
            'ERROR: First parameter to ADC must be register A
            Stop
        End If
        
        'If the second parameter is an 8-bit register: _
         ADC a, a|b|c|d|e|h|l
        
        
        'ADC a, (hl)                    'memory
        'ADC a, (ix+$8|iy+$8)           'indexed 8-bit
        'ADC a, $8                      '8-bit value
        'ADC hl, bc|de|hl|sp
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_ADD Then '................. `add` ...
        'ADD a, a|b|c|d|e|h|l
        'ADD a, $8
        'ADD a, (hl|ix+$8|iy+$8)
        'ADD hl, bc|de|hl|sp
        'ADD ix, bc|de|ix|sp
        'ADD iy, bc|de|iy|sp
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_AND Then '................. `and` ...
        'AND a|b|c|d|e|h|l
        'AND $8
        'AND (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_BIT Then '................. `bit` ...
        'BIT 0-7, a|b|c|d|e|h|l
        'BIT 0-7, (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CALL Then '............... `call` ...
        'CALL $16
        'CALL c|nc|m|p|z|nz|pe|po, $16
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CCF Then '................. `ccf` ...
        'No parameters, add opcode $3F
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CP Then '................... `cp` ...
        'CP a|b|c|d|e|h|l
        'CP $8
        'CP (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPD Then '................. `cpd` ...
        'No parameters, add opcodes $ED $A9
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPDR Then '............... `cpdr` ...
        'No parameters, add opcodes $ED $B9
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPI Then '................. `cpi` ...
        'No parameters, add opcodes $ED $A1
            
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPIR Then '............... `cpir` ...
        'No parameters, add opcodes $ED $B1
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_CPL Then '................. `cpl` ...
        'No parameters, add opcode $2F
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DAA Then '................. `daa` ...
        'No parameters, add opcode $27
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DEC Then '................. `dec` ...
        'DEC a|b|c|d|e|h|l
        'DEC (hl|ix+$8|iy+$8)
        'DEC bc|de|hl|sp|ix|iy
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DI Then '................... `di` ...
        'No parameters, add opcode $F3
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_DJNZ Then '............... `djnz` ...
        'DJNZ $8
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EI Then '................... `ei` ...
        'No parameters, add opcode $FB
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EX Then '................... `ex` ...
        'EX (sp), hl|ix|iy
        'EX af, af'
        'EX de, hl
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_EXX Then '................. `exx` ...
        'No parameters, add opcode $D9
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_HALT Then '............... `halt` ...
        'No parameters, add opcode $76
    
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
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INDR Then '............... `indr` ...
        'No parameters, add opcodes $ED $BA
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INI Then '................. `ini` ...
        'No parameters, add opcodes $ED $A2
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_INIR Then '............... `inir` ...
        'No parameters, add opcodes $ED $B2
        
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
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDDR Then '............... `lddr` ...
        'No parameters, add opcodes $ED $B8
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDI Then '................. `ldi` ...
        'No parameters, add opcodes $ED $A0
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_LDIR Then '............... `ldir` ...
        'No parameters, add opcodes $ED $B0
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NEG Then '................. `neg` ...
        'No parameters, add opcodes $ED $44
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_NOP Then '................. `nop` ...
        'No parameters, add opcode $00
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OR Then '................... `or` ...
        'OR a|b|c|d|e|h|l|$8
        'OR (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUT Then '................. `out` ...
        'OUT ($8), a
        'OUT (c), 0|a|b|c|d|e|h|l
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTD Then '............... `outd` ...
        'No parameters, add opcodes $ED $AB
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTDR Then '............... `otdr` ...
        'No parameters, add opcodes $ED $BB
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OUTI Then '............... `outi` ...
        'No parameters, add opcodes $ED $A3
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_OTIR Then '............... `otir` ...
        'No parameters, add opcodes $ED $B3
        
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
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RETN Then '............... `retn` ...
        'No parameters, add opcodes $ED $45
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLA Then '................. `rla` ...
        'No parameters, add opcode $17
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RL Then '................... `rl` ...
        'RL a|b|c|d|e|h|l
        'RL (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLCA Then '............... `rlca` ...
        'No parameters, add opcode $07
        
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLC Then '................. `rlc` ...
        'RLC a|b|c|d|e|h|l
        'RLC (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RLD Then '................. `rld` ...
        'No parameters, add opcodes $ED $6F
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRA Then '................. `rra` ...
        'No parameters, add opcode $1F
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RR Then '................... `rr` ...
        'RR a|b|c|d|e|h|l
        'RR (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRCA Then '............... `rrca` ...
        'No parameters, add opcode $0F
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRC Then '................. `rrc` ...
        'RRC a|b|c|d|e|h|l
        'RRC (hl|ix+$8|iy+$8)
    
    ElseIf Token.Data = OZ80_TOKEN_DATA.TOKEN_Z80_RRD Then '................. `rrd` ...
        'No parameters, add opcodes $ED $67
    
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
    Exit Sub

OutOfTokens:
    'TODO: handle the token stream ending unexpectedly
    Stop
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

Private Function NextToken() As Boolean
    Let TokenIndex = TokenIndex + 1
    If TokenIndex > TokenCount Then Exit Function
    Let Token = Tokens(TokenIndex)
    Let NextToken = True
End Function
