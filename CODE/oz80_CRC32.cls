VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oz80_CRC32"
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
'CLASS :: oz80_CRC32

'Checking strings against existing values is done often so a quick hashing routine _
 is used to get a unique number for each unique string.

'WARNING: This code is derived from "Calculate Checksum (CRC32 and CRC16)" _
 by Fredrik Qvarfort, and modified by Kroc Camen (reformatted and removed CRC16)
'Please contact him personally for permission to redistribute this source code _
 <www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=12638&lngWId=1>

'/// API //////////////////////////////////////////////////////////////////////////////

'<msdn.microsoft.com/en-gb/library/windows/desktop/ms633571%28v=vs.85%29.aspx>
Private Declare Function CallWindowProc Lib "user32" Alias "user32_CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

'/// CLASS DATA ///////////////////////////////////////////////////////////////////////

Private CRC32 As Long
Private CRCASM() As Byte                'Assembly code will be stored here
Private CRCTable(0 To 255) As Long      'A lookup table for the CRC calculation

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Initialise a precomputed CRC32 lookup table
    Let CRCTable(0) = &H0:          Let CRCTable(128) = &HEDB88320
    Let CRCTable(1) = &H77073096:   Let CRCTable(129) = &H9ABFB3B6
    Let CRCTable(2) = &HEE0E612C:   Let CRCTable(130) = &H3B6E20C
    Let CRCTable(3) = &H990951BA:   Let CRCTable(131) = &H74B1D29A
    Let CRCTable(4) = &H76DC419:    Let CRCTable(132) = &HEAD54739
    Let CRCTable(5) = &H706AF48F:   Let CRCTable(133) = &H9DD277AF
    Let CRCTable(6) = &HE963A535:   Let CRCTable(134) = &H4DB2615
    Let CRCTable(7) = &H9E6495A3:   Let CRCTable(135) = &H73DC1683
    Let CRCTable(8) = &HEDB8832:    Let CRCTable(136) = &HE3630B12
    Let CRCTable(9) = &H79DCB8A4:   Let CRCTable(137) = &H94643B84
    Let CRCTable(10) = &HE0D5E91E:  Let CRCTable(138) = &HD6D6A3E
    Let CRCTable(11) = &H97D2D988:  Let CRCTable(139) = &H7A6A5AA8
    Let CRCTable(12) = &H9B64C2B:   Let CRCTable(140) = &HE40ECF0B
    Let CRCTable(13) = &H7EB17CBD:  Let CRCTable(141) = &H9309FF9D
    Let CRCTable(14) = &HE7B82D07:  Let CRCTable(142) = &HA00AE27
    Let CRCTable(15) = &H90BF1D91:  Let CRCTable(143) = &H7D079EB1
    Let CRCTable(16) = &H1DB71064:  Let CRCTable(144) = &HF00F9344
    Let CRCTable(17) = &H6AB020F2:  Let CRCTable(145) = &H8708A3D2
    Let CRCTable(18) = &HF3B97148:  Let CRCTable(146) = &H1E01F268
    Let CRCTable(19) = &H84BE41DE:  Let CRCTable(147) = &H6906C2FE
    Let CRCTable(20) = &H1ADAD47D:  Let CRCTable(148) = &HF762575D
    Let CRCTable(21) = &H6DDDE4EB:  Let CRCTable(149) = &H806567CB
    Let CRCTable(22) = &HF4D4B551:  Let CRCTable(150) = &H196C3671
    Let CRCTable(23) = &H83D385C7:  Let CRCTable(151) = &H6E6B06E7
    Let CRCTable(24) = &H136C9856:  Let CRCTable(152) = &HFED41B76
    Let CRCTable(25) = &H646BA8C0:  Let CRCTable(153) = &H89D32BE0
    Let CRCTable(26) = &HFD62F97A:  Let CRCTable(154) = &H10DA7A5A
    Let CRCTable(27) = &H8A65C9EC:  Let CRCTable(155) = &H67DD4ACC
    Let CRCTable(28) = &H14015C4F:  Let CRCTable(156) = &HF9B9DF6F
    Let CRCTable(29) = &H63066CD9:  Let CRCTable(157) = &H8EBEEFF9
    Let CRCTable(30) = &HFA0F3D63:  Let CRCTable(158) = &H17B7BE43
    Let CRCTable(31) = &H8D080DF5:  Let CRCTable(159) = &H60B08ED5
    Let CRCTable(32) = &H3B6E20C8:  Let CRCTable(160) = &HD6D6A3E8
    Let CRCTable(33) = &H4C69105E:  Let CRCTable(161) = &HA1D1937E
    Let CRCTable(34) = &HD56041E4:  Let CRCTable(162) = &H38D8C2C4
    Let CRCTable(35) = &HA2677172:  Let CRCTable(163) = &H4FDFF252
    Let CRCTable(36) = &H3C03E4D1:  Let CRCTable(164) = &HD1BB67F1
    Let CRCTable(37) = &H4B04D447:  Let CRCTable(165) = &HA6BC5767
    Let CRCTable(38) = &HD20D85FD:  Let CRCTable(166) = &H3FB506DD
    Let CRCTable(39) = &HA50AB56B:  Let CRCTable(167) = &H48B2364B
    Let CRCTable(40) = &H35B5A8FA:  Let CRCTable(168) = &HD80D2BDA
    Let CRCTable(41) = &H42B2986C:  Let CRCTable(169) = &HAF0A1B4C
    Let CRCTable(42) = &HDBBBC9D6:  Let CRCTable(170) = &H36034AF6
    Let CRCTable(43) = &HACBCF940:  Let CRCTable(171) = &H41047A60
    Let CRCTable(44) = &H32D86CE3:  Let CRCTable(172) = &HDF60EFC3
    Let CRCTable(45) = &H45DF5C75:  Let CRCTable(173) = &HA867DF55
    Let CRCTable(46) = &HDCD60DCF:  Let CRCTable(174) = &H316E8EEF
    Let CRCTable(47) = &HABD13D59:  Let CRCTable(175) = &H4669BE79
    Let CRCTable(48) = &H26D930AC:  Let CRCTable(176) = &HCB61B38C
    Let CRCTable(49) = &H51DE003A:  Let CRCTable(177) = &HBC66831A
    Let CRCTable(50) = &HC8D75180:  Let CRCTable(178) = &H256FD2A0
    Let CRCTable(51) = &HBFD06116:  Let CRCTable(179) = &H5268E236
    Let CRCTable(52) = &H21B4F4B5:  Let CRCTable(180) = &HCC0C7795
    Let CRCTable(53) = &H56B3C423:  Let CRCTable(181) = &HBB0B4703
    Let CRCTable(54) = &HCFBA9599:  Let CRCTable(182) = &H220216B9
    Let CRCTable(55) = &HB8BDA50F:  Let CRCTable(183) = &H5505262F
    Let CRCTable(56) = &H2802B89E:  Let CRCTable(184) = &HC5BA3BBE
    Let CRCTable(57) = &H5F058808:  Let CRCTable(185) = &HB2BD0B28
    Let CRCTable(58) = &HC60CD9B2:  Let CRCTable(186) = &H2BB45A92
    Let CRCTable(59) = &HB10BE924:  Let CRCTable(187) = &H5CB36A04
    Let CRCTable(60) = &H2F6F7C87:  Let CRCTable(188) = &HC2D7FFA7
    Let CRCTable(61) = &H58684C11:  Let CRCTable(189) = &HB5D0CF31
    Let CRCTable(62) = &HC1611DAB:  Let CRCTable(190) = &H2CD99E8B
    Let CRCTable(63) = &HB6662D3D:  Let CRCTable(191) = &H5BDEAE1D
    Let CRCTable(64) = &H76DC4190:  Let CRCTable(192) = &H9B64C2B0
    Let CRCTable(65) = &H1DB7106:   Let CRCTable(193) = &HEC63F226
    Let CRCTable(66) = &H98D220BC:  Let CRCTable(194) = &H756AA39C
    Let CRCTable(67) = &HEFD5102A:  Let CRCTable(195) = &H26D930A
    Let CRCTable(68) = &H71B18589:  Let CRCTable(196) = &H9C0906A9
    Let CRCTable(69) = &H6B6B51F:   Let CRCTable(197) = &HEB0E363F
    Let CRCTable(70) = &H9FBFE4A5:  Let CRCTable(198) = &H72076785
    Let CRCTable(71) = &HE8B8D433:  Let CRCTable(199) = &H5005713
    Let CRCTable(72) = &H7807C9A2:  Let CRCTable(200) = &H95BF4A82
    Let CRCTable(73) = &HF00F934:   Let CRCTable(201) = &HE2B87A14
    Let CRCTable(74) = &H9609A88E:  Let CRCTable(202) = &H7BB12BAE
    Let CRCTable(75) = &HE10E9818:  Let CRCTable(203) = &HCB61B38
    Let CRCTable(76) = &H7F6A0DBB:  Let CRCTable(204) = &H92D28E9B
    Let CRCTable(77) = &H86D3D2D:   Let CRCTable(205) = &HE5D5BE0D
    Let CRCTable(78) = &H91646C97:  Let CRCTable(206) = &H7CDCEFB7
    Let CRCTable(79) = &HE6635C01:  Let CRCTable(207) = &HBDBDF21
    Let CRCTable(80) = &H6B6B51F4:  Let CRCTable(208) = &H86D3D2D4
    Let CRCTable(81) = &H1C6C6162:  Let CRCTable(209) = &HF1D4E242
    Let CRCTable(82) = &H856530D8:  Let CRCTable(210) = &H68DDB3F8
    Let CRCTable(83) = &HF262004E:  Let CRCTable(211) = &H1FDA836E
    Let CRCTable(84) = &H6C0695ED:  Let CRCTable(212) = &H81BE16CD
    Let CRCTable(85) = &H1B01A57B:  Let CRCTable(213) = &HF6B9265B
    Let CRCTable(86) = &H8208F4C1:  Let CRCTable(214) = &H6FB077E1
    Let CRCTable(87) = &HF50FC457:  Let CRCTable(215) = &H18B74777
    Let CRCTable(88) = &H65B0D9C6:  Let CRCTable(216) = &H88085AE6
    Let CRCTable(89) = &H12B7E950:  Let CRCTable(217) = &HFF0F6A70
    Let CRCTable(90) = &H8BBEB8EA:  Let CRCTable(218) = &H66063BCA
    Let CRCTable(91) = &HFCB9887C:  Let CRCTable(219) = &H11010B5C
    Let CRCTable(92) = &H62DD1DDF:  Let CRCTable(220) = &H8F659EFF
    Let CRCTable(93) = &H15DA2D49:  Let CRCTable(221) = &HF862AE69
    Let CRCTable(94) = &H8CD37CF3:  Let CRCTable(222) = &H616BFFD3
    Let CRCTable(95) = &HFBD44C65:  Let CRCTable(223) = &H166CCF45
    Let CRCTable(96) = &H4DB26158:  Let CRCTable(224) = &HA00AE278
    Let CRCTable(97) = &H3AB551CE:  Let CRCTable(225) = &HD70DD2EE
    Let CRCTable(98) = &HA3BC0074:  Let CRCTable(226) = &H4E048354
    Let CRCTable(99) = &HD4BB30E2:  Let CRCTable(227) = &H3903B3C2
    Let CRCTable(100) = &H4ADFA541: Let CRCTable(228) = &HA7672661
    Let CRCTable(101) = &H3DD895D7: Let CRCTable(229) = &HD06016F7
    Let CRCTable(102) = &HA4D1C46D: Let CRCTable(230) = &H4969474D
    Let CRCTable(103) = &HD3D6F4FB: Let CRCTable(231) = &H3E6E77DB
    Let CRCTable(104) = &H4369E96A: Let CRCTable(232) = &HAED16A4A
    Let CRCTable(105) = &H346ED9FC: Let CRCTable(233) = &HD9D65ADC
    Let CRCTable(106) = &HAD678846: Let CRCTable(234) = &H40DF0B66
    Let CRCTable(107) = &HDA60B8D0: Let CRCTable(235) = &H37D83BF0
    Let CRCTable(108) = &H44042D73: Let CRCTable(236) = &HA9BCAE53
    Let CRCTable(109) = &H33031DE5: Let CRCTable(237) = &HDEBB9EC5
    Let CRCTable(110) = &HAA0A4C5F: Let CRCTable(238) = &H47B2CF7F
    Let CRCTable(111) = &HDD0D7CC9: Let CRCTable(239) = &H30B5FFE9
    Let CRCTable(112) = &H5005713C: Let CRCTable(240) = &HBDBDF21C
    Let CRCTable(113) = &H270241AA: Let CRCTable(241) = &HCABAC28A
    Let CRCTable(114) = &HBE0B1010: Let CRCTable(242) = &H53B39330
    Let CRCTable(115) = &HC90C2086: Let CRCTable(243) = &H24B4A3A6
    Let CRCTable(116) = &H5768B525: Let CRCTable(244) = &HBAD03605
    Let CRCTable(117) = &H206F85B3: Let CRCTable(245) = &HCDD70693
    Let CRCTable(118) = &HB966D409: Let CRCTable(246) = &H54DE5729
    Let CRCTable(119) = &HCE61E49F: Let CRCTable(247) = &H23D967BF
    Let CRCTable(120) = &H5EDEF90E: Let CRCTable(248) = &HB3667A2E
    Let CRCTable(121) = &H29D9C998: Let CRCTable(249) = &HC4614AB8
    Let CRCTable(122) = &HB0D09822: Let CRCTable(250) = &H5D681B02
    Let CRCTable(123) = &HC7D7A8B4: Let CRCTable(251) = &H2A6F2B94
    Let CRCTable(124) = &H59B33D17: Let CRCTable(252) = &HB40BBE37
    Let CRCTable(125) = &H2EB40D81: Let CRCTable(253) = &HC30C8EA1
    Let CRCTable(126) = &HB7BD5C3B: Let CRCTable(254) = &H5A05DF1B
    Let CRCTable(127) = &HC0BA6CAD: Let CRCTable(255) = &H2D02EF8D
    
    'Get the ASM code into memory: _
     ----------------------------------------------------------------------------------
    Dim ASM As String
    Let ASM = _
        "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F46497" & _
        "5F28B4D088901595B585E5F89EC5DC21000"
    
    ReDim CRCASM(0 To Len(ASM) \ 2 - 1)
    
    Dim i As Long
    For i = 1 To Len(ASM) Step 2
        CRCASM(i \ 2) = Val("&H" & Mid$(ASM, i, 2))
    Next
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    Erase CRCTable
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'AddBytes : _
 ======================================================================================
Public Function AppendBytes(ByRef ByteArray() As Byte) As Long
    'We need to add a simple error trapping
    'here because if the bytearray is not
    'dimensioned we want it to just skip
    'the assembler code call below
    On Local Error GoTo NoData
    
    'Precalculate the size of the byte array
    Dim ByteSize As Long
    Let ByteSize = UBound(ByteArray) - LBound(ByteArray) + 1
  
    'No error trapping needed, if something
    'goes bad below something is definitely
    'fishy with your computer
    On Local Error GoTo 0
      
    'Run the pre-compiled assembler code
    Call CallWindowProc( _
        VarPtr(CRCASM(0)), VarPtr(CRC32), _
        VarPtr(ByteArray(LBound(ByteArray))), VarPtr(CRCTable(0)), ByteSize _
    )
  
NoData:
    'Return the current CRC value
    Let AppendBytes = (Not CRC32)
End Function

'AddString : _
 ======================================================================================
Public Function AppendString(ByRef Text As String) As Long

    'Convert the string into a byte array
    'and send it to the function that can
    'handle bytearrays
    Let AppendString = AppendBytes(StrConv(Text, vbFromUnicode))
  
End Function

'CalculateBytes : _
 ======================================================================================
Public Function CalculateBytes(ByRef ByteArray() As Byte) As Long

    'Reset the current CRC calculation
    Let CRC32 = &HFFFFFFFF
  
    'Calculate the CRC from the bytearray
    'and return the current CRC value
    Let CalculateBytes = AppendBytes(ByteArray)
  
End Function

'CalculateFile : _
 ======================================================================================
Public Function CalculateFile(ByRef Filename As String) As Long

    Dim Filenr As Integer
    Dim ByteArray() As Byte
  
    'Make sure the file contains data
    'to avoid errors later below
    If (FileLen(Filename) = 0) Then Exit Function
  
    'Open the file in binary mode, read
    'the data into a bytearray and then
    'close the file
    Filenr = FreeFile
    Open Filename For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr
  
    'Now send the bytearray to the function
    'that can calculate a CRC from it
    CalculateFile = CalculateBytes(ByteArray)

End Function

'CalculateString : _
 ======================================================================================
Public Function CalculateString(ByRef Text As String) As Long

    'Convert the string into a bytearray
    'and send it to the function that
    'calculates the CRC from a bytearray
    Let CalculateString = CalculateBytes(StrConv(Text, vbFromUnicode))
  
End Function
