VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OZ80VM"
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
'CLASS :: OZ80VM

'Being both an assembler and a programming langauge, it would be too slow to keep _
 walking over the source code to do the more complex programmy things like functions. _
 Once all the source code is in memory we can instead write out an optimised list of _
 instructions that simplifies all the actions of the user's program down to a stream _
 of read/calculate/write instructions

'To achieve this we need a Virtual Machine -- not a Virtual Machine for running one _
 Operating System within another, but an emulator for a very simple hypothetical _
 processor that reads instructions, does calulcations and outputs results

'OZ80's VM is based on a number of VM designs, including Forth (stack-based), _
 a Turing Machine ('tape'-based) and a hybrid of other features including simple _
 registers (Accumulator). This is because OZ80's VM is designed for file reading & _
 writing at a high-level and not for emulating a low-level system across multiple _
 environments

'The following points are very important to understand:

' * The VM is self-contained, and the terminology herein is unique to OZ80_VM

' * OZ80_VM works internally with "Tapes". A Tape is an array in memory with a reading _
    Head (Index). The Tape can be wound forward or backward and data read/written. _
    Files are read into memory as a Tape and handled as any other Tape in the program

' * Ports are communication channels with the world outside the VM. _
    These are managed by class events you can sink to interact with the running VM

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'Each Instruction consists of a 4-byte Opcode (purely for speed reasons, _
 1-byte would suffice) and an 8-Byte 64-Bit floating-point Parameter value
Private Type OZ80VM_Instruction
    Opcode As Long
    Param As Double
    
    'TODO: We should include metadata from the original source for error handling, _
           e.g. File, Line & Col
End Type

Public Enum OZ80VM_OPCODES
    VM_HALT                             'Stops the VM
    
    'Tape Operations
    '..................................................................................
    VM_CTAP                             'Creates a new Tape
    VM_TAPE                             'Selects a Tape
    VM_LOAD                             'Loads a Tape from Disk
    VM_SAVE                             'Saves a Tape to Disk
    
    VM_HEAD                             'Gets the position of the current Tape
    VM_LENT                             '"Length of Tape"
    
    VM_FWRD                             'Forward the Tape a given number of places
    VM_RWND                             'Rewind the Tape a given number of places
    VM_WIND                             'Wind to a specific point in the Tape
    
    VM_RBYT                             'Read a Byte from Tape
    VM_RWRD                             'Read a Word (two bytes) from Tape
    VM_RTRP                             'Read a Triple (three bytes) from Tape
    VM_RLNG                             'Read a Long (four bytes) from Tape
    VM_RFLT                             'Read a Float (8-bytes) from Tape
    
    VM_WBYT                             'Write a Byte to Tape
    VM_WWRD                             'Write a Word (two bytes) to Tape
    VM_WTRP                             'Write a Triple (three bytes) to Tape
    VM_WLNG                             'Write a Long (four bytes) to Tape
    VM_WFLT                             'Write a Float (8-bytes) to Tape
    
    VM_DTAP                             '"Drop Tape", throws away a Tape
    
    'Port operations
    '..................................................................................
    VM_OPEN                             'Opens a Port
    VM_PORT                             'Selects a Port
    VM_SEND                             'Sends A down current Port
    VM_READ                             'Fetches A from a Port
    VM_WAIT                             'Suspends VM until response from Port
    VM_SHUT                             'Shuts a Port
    
    'Stack Operations
    '..................................................................................
    VM_PUSH                             'Push Accumulator to Stack
    VM_PULL                             'Pull top of Stack into Accumulator
    VM_DROP                             'Throw away the top Value on the Stack
    
    'Calculations
    '..................................................................................
    VM_LETA                             'Sets A to the Value in Param
    VM_SWAP                             'Swaps A & B
    
    VM_ADDA                             'A + B
    VM_SUBA                             'A - B
    VM_MULA                             'A * B
    VM_DIVA                             'A / B
    VM_POWA                             'A ^ B
    VM_MODA                             'A MOD B
    VM_ANDA                             'A AND B
    VM_AORB                             'A OR B
    
    'Error Handling
    '..................................................................................
    VM_EROR                             'Gets the Error state
    
    'Flow Control
    '..................................................................................
    VM_JUMP                             'Jumps to a specific Instruction
    VM_CALL                             'Jumps to a function, remembering return
    VM_RTRN                             'Return from a function
End Enum

Private Program() As OZ80VM_Instruction
Private Length As Long                  'Current length of the code in use
Private Buffer As Long

'Processor:
'--------------------------------------------------------------------------------------
Private A As Double                     'Accumulator
Private B As Double                     'B-Accumulator

Private PC As Double                    'Program-Counter (current Instruction index)

'Tapes: (binary arrays in memory)
'--------------------------------------------------------------------------------------
Private Type OZ80VM_Tape
    Data() As Byte
    Head As Long
End Type

Private Tapes() As OZ80VM_Tape
Private Tape As Long                    'Current Tape in use

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'When the VM opens a communications channel
Event PortOpen( _
    ByRef Port As Long _
)

'When the VM wants to read from a Port. _
 Your program can respond to this event and feed the VM the data value
Event PortRead( _
    ByRef Port As Long, _
    ByRef Data As Double _
)

'When the VM is sending data to a Port
Event PortSend( _
    ByRef Port As Long, _
    ByRef Data As Long _
)

'When the VM closes a communications channel
Event PortShut( _
    ByRef Port As Long _
)

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Call Clear
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddInstruction : Add an Instruction to the current program
'======================================================================================
Public Sub AddInstruction( _
    ByRef Operation As OZ80VM_OPCODES, _
    ByRef Parameter As Double _
)
    Let Length = Length + 1
    ReDim Preserve Program(1 To Length) As OZ80VM_Instruction
    
    Let Program(Length).Opcode = Operation
    Let Program(Length).Param = Parameter
End Sub

'Clear : Erase the program
'======================================================================================
Public Sub Clear()
    Erase Program, Tapes
    Let A = 0: Let B = 0: Let PC = 0
End Sub

'Load : Load a program from file
'======================================================================================
Public Sub Load( _
    ByRef FilePath As String _
)
    'Clear any existing program
    Call Clear
    
    'Open the file
    Dim FileNumber As Long: Let FileNumber = FreeFile
    Open FilePath For Input Lock Write As #FileNumber
    
    Do While Not EOF(FileNumber)
        'Read a line in
        Dim Line As String
        Line Input #FileNumber, Line
        
        'Trim the whitespace
        Let Line = Trim(Line)
        
        If (Left$(Line, 1) <> ";") _
        Or (Line <> vbNullString) _
        Then
            Dim Text As String
            Let Text = Left$(Line, 4)
            
            Dim Operation As OZ80VM_OPCODES
            Select Case Text
            Case "HALT": Let Operation = VM_HALT
            Case "CTAP": Let Operation = VM_CTAP
            Case "TAPE": Let Operation = VM_TAPE
            Case "LOAD": Let Operation = VM_LOAD
            Case "SAVE": Let Operation = VM_SAVE
            Case "HEAD": Let Operation = VM_HEAD
            Case "LENT": Let Operation = VM_LENT
            Case "FWRD": Let Operation = VM_FWRD
            Case "RWND": Let Operation = VM_RWND
            Case "WIND": Let Operation = VM_WIND
            Case "RBYT": Let Operation = VM_RBYT
            Case "RWRD": Let Operation = VM_RWRD
            Case "RTRP": Let Operation = VM_RTRP
            Case "RLNG": Let Operation = VM_RLNG
            Case "RFLT": Let Operation = VM_RFLT
            Case "WBYT": Let Operation = VM_WBYT
            Case "WWRD": Let Operation = VM_WWRD
            Case "WTRP": Let Operation = VM_WTRP
            Case "WLNG": Let Operation = VM_WLNG
            Case "WFLT": Let Operation = VM_WFLT
            Case "DTAP": Let Operation = VM_DTAP
            Case "OPEN": Let Operation = VM_OPEN
            Case "PORT": Let Operation = VM_PORT
            Case "SEND": Let Operation = VM_SEND
            Case "READ": Let Operation = VM_READ
            Case "WAIT": Let Operation = VM_WAIT
            Case "SHUT": Let Operation = VM_SHUT
            Case "PUSH": Let Operation = VM_PUSH
            Case "PULL": Let Operation = VM_PULL
            Case "DROP": Let Operation = VM_DROP
            Case "LETA": Let Operation = VM_LETA
            Case "SWAP": Let Operation = VM_SWAP
            Case "ADDA": Let Operation = VM_ADDA
            Case "SUBA": Let Operation = VM_SUBA
            Case "MULA": Let Operation = VM_MULA
            Case "DIVA": Let Operation = VM_DIVA
            Case "POWA": Let Operation = VM_POWA
            Case "MODA": Let Operation = VM_MODA
            Case "ANDA": Let Operation = VM_ANDA
            Case "AORB": Let Operation = VM_AORB
            Case "EROR": Let Operation = VM_EROR
            Case "JUMP": Let Operation = VM_JUMP
            Case "CALL": Let Operation = VM_CALL
            Case "RTRN": Let Operation = VM_RTRN
            
            Case Else
                Stop
            End Select
            
            Let Text = Trim(Mid$(Line, 5))
            
            If Text = vbNullString Then
                Call AddInstruction(Operation, 0)
            Else
                Dim Param As Double
                Let Param = CDbl(Text)
                Call AddInstruction(Operation, Param)
            End If
        End If
    Loop
    
    'Close the file
    Close #FileNumber
End Sub

'Run : Run the current program
'======================================================================================
Public Sub Run()
    Let A = 0: Let B = 0: Let PC = 0
    
End Sub

'Save : Save the program to file
'======================================================================================
Public Sub Save()
    '
End Sub
