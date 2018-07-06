Attribute VB_Name = "rtPseudoVars"
' ---------------------------------
' --- rtPseudoVars - 02/09/2010 ---
' ---------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 02/09/2010 - SBakker - URD 11076
'            - Make sure .CursorLocation is always the first property set.
' 01/22/2007 - SBAKKER - URD 9739
'            - Added error checking on the cnSQL object before using it, to
'              prevent hanging runtime processes.
' 06/26/2006 - Added KBCX function, which returns the number of keyboard chars
'              just like KBC does, but always returns zero if a keyboard script
'              is currently running. Must change CADOL programs to use this:
'                 IF KBCX#0 THEN ... ! ABORTING
'              Can only be used around Aborting code, as other parts of scripts
'              (like Policy Main Menu) need to still check KBC accurately.
'              Also added corresponding LET_KBCX routine.
' 01/20/2006 - Added N64-N99, A3-A9 to E3-E9, F10-F99.
' 01/20/2006 - Added error checking in LET_G1 to LET_G9, making sure that G has
'              been set to the current user number before setting the value.
'            - Added G_OFS and LET_G_OFS routines, to allow "G[x]" syntax to be
'              used in Cadol programs. IDRISMakeLib and IDRIS_IDE have always
'              allowed this syntax, but it wouldn't have compiled before.
' ------------------------------------------------------------------------------

' --- N(Offset) ---

Public Sub LET_N_OFS(ByVal Offset As Long, ByVal Value As Currency)
   Select Case Offset
      Case 0: N = Value
      Case 1: N1 = Value
      Case 2: N2 = Value
      Case 3: N3 = Value
      Case 4: N4 = Value
      Case 5: N5 = Value
      Case 6: N6 = Value
      Case 7: N7 = Value
      Case 8: N8 = Value
      Case 9: N9 = Value
      Case 10: N10 = Value
      Case 11: N11 = Value
      Case 12: N12 = Value
      Case 13: N13 = Value
      Case 14: N14 = Value
      Case 15: N15 = Value
      Case 16: N16 = Value
      Case 17: N17 = Value
      Case 18: N18 = Value
      Case 19: N19 = Value
      Case 20: N20 = Value
      Case 21: N21 = Value
      Case 22: N22 = Value
      Case 23: N23 = Value
      Case 24: N24 = Value
      Case 25: N25 = Value
      Case 26: N26 = Value
      Case 27: N27 = Value
      Case 28: N28 = Value
      Case 29: N29 = Value
      Case 30: N30 = Value
      Case 31: N31 = Value
      Case 32: N32 = Value
      Case 33: N33 = Value
      Case 34: N34 = Value
      Case 35: N35 = Value
      Case 36: N36 = Value
      Case 37: N37 = Value
      Case 38: N38 = Value
      Case 39: N39 = Value
      Case 40: N40 = Value
      Case 41: N41 = Value
      Case 42: N42 = Value
      Case 43: N43 = Value
      Case 44: N44 = Value
      Case 45: N45 = Value
      Case 46: N46 = Value
      Case 47: N47 = Value
      Case 48: N48 = Value
      Case 49: N49 = Value
      Case 50: N50 = Value
      Case 51: N51 = Value
      Case 52: N52 = Value
      Case 53: N53 = Value
      Case 54: N54 = Value
      Case 55: N55 = Value
      Case 56: N56 = Value
      Case 57: N57 = Value
      Case 58: N58 = Value
      Case 59: N59 = Value
      Case 60: N60 = Value
      Case 61: N61 = Value
      Case 62: N62 = Value
      Case 63: N63 = Value
      ' --- high numeric registers ---
      Case 64: N64 = Value
      Case 65: N65 = Value
      Case 66: N66 = Value
      Case 67: N67 = Value
      Case 68: N68 = Value
      Case 69: N69 = Value
      Case 70: N70 = Value
      Case 71: N71 = Value
      Case 72: N72 = Value
      Case 73: N73 = Value
      Case 74: N74 = Value
      Case 75: N75 = Value
      Case 76: N76 = Value
      Case 77: N77 = Value
      Case 78: N78 = Value
      Case 79: N79 = Value
      Case 80: N80 = Value
      Case 81: N81 = Value
      Case 82: N82 = Value
      Case 83: N83 = Value
      Case 84: N84 = Value
      Case 85: N85 = Value
      Case 86: N86 = Value
      Case 87: N87 = Value
      Case 88: N88 = Value
      Case 89: N89 = Value
      Case 90: N90 = Value
      Case 91: N91 = Value
      Case 92: N92 = Value
      Case 93: N93 = Value
      Case 94: N94 = Value
      Case 95: N95 = Value
      Case 96: N96 = Value
      Case 97: N97 = Value
      Case 98: N98 = Value
      Case 99: N99 = Value
      Case Else
         ThrowError "LET_N_OFS", "Invalid Offset: " & Trim$(Str$(Offset))
   End Select
End Sub
Public Function N_OFS(ByVal Offset As Long) As Currency
   Select Case Offset
      Case 0: N_OFS = N
      Case 1: N_OFS = N1
      Case 2: N_OFS = N2
      Case 3: N_OFS = N3
      Case 4: N_OFS = N4
      Case 5: N_OFS = N5
      Case 6: N_OFS = N6
      Case 7: N_OFS = N7
      Case 8: N_OFS = N8
      Case 9: N_OFS = N9
      Case 10: N_OFS = N10
      Case 11: N_OFS = N11
      Case 12: N_OFS = N12
      Case 13: N_OFS = N13
      Case 14: N_OFS = N14
      Case 15: N_OFS = N15
      Case 16: N_OFS = N16
      Case 17: N_OFS = N17
      Case 18: N_OFS = N18
      Case 19: N_OFS = N19
      Case 20: N_OFS = N20
      Case 21: N_OFS = N21
      Case 22: N_OFS = N22
      Case 23: N_OFS = N23
      Case 24: N_OFS = N24
      Case 25: N_OFS = N25
      Case 26: N_OFS = N26
      Case 27: N_OFS = N27
      Case 28: N_OFS = N28
      Case 29: N_OFS = N29
      Case 30: N_OFS = N30
      Case 31: N_OFS = N31
      Case 32: N_OFS = N32
      Case 33: N_OFS = N33
      Case 34: N_OFS = N34
      Case 35: N_OFS = N35
      Case 36: N_OFS = N36
      Case 37: N_OFS = N37
      Case 38: N_OFS = N38
      Case 39: N_OFS = N39
      Case 40: N_OFS = N40
      Case 41: N_OFS = N41
      Case 42: N_OFS = N42
      Case 43: N_OFS = N43
      Case 44: N_OFS = N44
      Case 45: N_OFS = N45
      Case 46: N_OFS = N46
      Case 47: N_OFS = N47
      Case 48: N_OFS = N48
      Case 49: N_OFS = N49
      Case 50: N_OFS = N50
      Case 51: N_OFS = N51
      Case 52: N_OFS = N52
      Case 53: N_OFS = N53
      Case 54: N_OFS = N54
      Case 55: N_OFS = N55
      Case 56: N_OFS = N56
      Case 57: N_OFS = N57
      Case 58: N_OFS = N58
      Case 59: N_OFS = N59
      Case 60: N_OFS = N60
      Case 61: N_OFS = N61
      Case 62: N_OFS = N62
      Case 63: N_OFS = N63
      ' --- high numeric registers ---
      Case 64: N_OFS = N64
      Case 65: N_OFS = N65
      Case 66: N_OFS = N66
      Case 67: N_OFS = N67
      Case 68: N_OFS = N68
      Case 69: N_OFS = N69
      Case 70: N_OFS = N70
      Case 71: N_OFS = N71
      Case 72: N_OFS = N72
      Case 73: N_OFS = N73
      Case 74: N_OFS = N74
      Case 75: N_OFS = N75
      Case 76: N_OFS = N76
      Case 77: N_OFS = N77
      Case 78: N_OFS = N78
      Case 79: N_OFS = N79
      Case 80: N_OFS = N80
      Case 81: N_OFS = N81
      Case 82: N_OFS = N82
      Case 83: N_OFS = N83
      Case 84: N_OFS = N84
      Case 85: N_OFS = N85
      Case 86: N_OFS = N86
      Case 87: N_OFS = N87
      Case 88: N_OFS = N88
      Case 89: N_OFS = N89
      Case 90: N_OFS = N90
      Case 91: N_OFS = N91
      Case 92: N_OFS = N92
      Case 93: N_OFS = N93
      Case 94: N_OFS = N94
      Case 95: N_OFS = N95
      Case 96: N_OFS = N96
      Case 97: N_OFS = N97
      Case 98: N_OFS = N98
      Case 99: N_OFS = N99
      Case Else
         ThrowError "N_OFS", "Invalid Offset: " & Trim$(Str$(Offset))
   End Select
End Function

' --- KEY ---

Public Function KEY() As String
   KEY = GetAlpha(MemPos_Key)
End Function
Public Function KEY_OFS(ByVal Offset As Long) As String
   KEY_OFS = GetAlpha(MemPos_Key + Offset)
End Function
Public Sub LET_KEY(ByVal Value As String)
   LetAlpha MemPos_Key, Value
End Sub
Public Sub LET_KEY_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_Key + Offset, Value
End Sub
Public Sub SPOOL_KEY(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_Key, Size, Value
End Sub
Public Sub SPOOL_KEY_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_Key + Offset, Size, Value
End Sub

' --- DATEVAL ---

Public Function DATEVAL() As String
   DATEVAL = GetAlpha(MemPos_DateVal)
End Function
Public Function DATEVAL_OFS(ByVal Offset As Long) As String
   DATEVAL_OFS = GetAlpha(MemPos_DateVal + Offset)
End Function
Public Sub LET_DATEVAL(ByVal Value As String)
   LetAlpha MemPos_DateVal, Value
End Sub
Public Sub LET_DATEVAL_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_DateVal + Offset, Value
End Sub
Public Sub SPOOL_DATEVAL(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_DateVal, Size, Value
End Sub
Public Sub SPOOL_DATEVAL_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_DateVal + Offset, Size, Value
End Sub

' --- A ---

Public Function A() As String
   A = GetAlpha(MemPos_A)
End Function
Public Function A_OFS(ByVal Offset As Long) As String
   A_OFS = GetAlpha(MemPos_A + Offset)
End Function
Public Sub LET_A(ByVal Value As String)
   LetAlpha MemPos_A, Value
End Sub
Public Sub LET_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A + Offset, Value
End Sub
Public Sub SPOOL_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A, Size, Value
End Sub
Public Sub SPOOL_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A + Offset, Size, Value
End Sub

Public Function A1() As String
   A1 = GetAlpha(MemPos_A1)
End Function
Public Function A1_OFS(ByVal Offset As Long) As String
   A1_OFS = GetAlpha(MemPos_A1 + Offset)
End Function
Public Sub LET_A1(ByVal Value As String)
   LetAlpha MemPos_A1, Value
End Sub
Public Sub LET_A1_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A1 + Offset, Value
End Sub
Public Sub SPOOL_A1(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A1, Size, Value
End Sub
Public Sub SPOOL_A1_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A1 + Offset, Size, Value
End Sub

Public Function A2() As String
   A2 = GetAlpha(MemPos_A2)
End Function
Public Function A2_OFS(ByVal Offset As Long) As String
   A2_OFS = GetAlpha(MemPos_A2 + Offset)
End Function
Public Sub LET_A2(ByVal Value As String)
   LetAlpha MemPos_A2, Value
End Sub
Public Sub LET_A2_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A2 + Offset, Value
End Sub
Public Sub SPOOL_A2(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A2, Size, Value
End Sub
Public Sub SPOOL_A2_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A2 + Offset, Size, Value
End Sub

Public Function A3() As String
   A3 = GetAlpha(MemPos_A3)
End Function
Public Function A3_OFS(ByVal Offset As Long) As String
   A3_OFS = GetAlpha(MemPos_A3 + Offset)
End Function
Public Sub LET_A3(ByVal Value As String)
   LetAlpha MemPos_A3, Value
End Sub
Public Sub LET_A3_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A3 + Offset, Value
End Sub
Public Sub SPOOL_A3(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A3, Size, Value
End Sub
Public Sub SPOOL_A3_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A3 + Offset, Size, Value
End Sub

Public Function A4() As String
   A4 = GetAlpha(MemPos_A4)
End Function
Public Function A4_OFS(ByVal Offset As Long) As String
   A4_OFS = GetAlpha(MemPos_A4 + Offset)
End Function
Public Sub LET_A4(ByVal Value As String)
   LetAlpha MemPos_A4, Value
End Sub
Public Sub LET_A4_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A4 + Offset, Value
End Sub
Public Sub SPOOL_A4(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A4, Size, Value
End Sub
Public Sub SPOOL_A4_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A4 + Offset, Size, Value
End Sub

Public Function A5() As String
   A5 = GetAlpha(MemPos_A5)
End Function
Public Function A5_OFS(ByVal Offset As Long) As String
   A5_OFS = GetAlpha(MemPos_A5 + Offset)
End Function
Public Sub LET_A5(ByVal Value As String)
   LetAlpha MemPos_A5, Value
End Sub
Public Sub LET_A5_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A5 + Offset, Value
End Sub
Public Sub SPOOL_A5(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A5, Size, Value
End Sub
Public Sub SPOOL_A5_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A5 + Offset, Size, Value
End Sub

Public Function A6() As String
   A6 = GetAlpha(MemPos_A6)
End Function
Public Function A6_OFS(ByVal Offset As Long) As String
   A6_OFS = GetAlpha(MemPos_A6 + Offset)
End Function
Public Sub LET_A6(ByVal Value As String)
   LetAlpha MemPos_A6, Value
End Sub
Public Sub LET_A6_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A6 + Offset, Value
End Sub
Public Sub SPOOL_A6(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A6, Size, Value
End Sub
Public Sub SPOOL_A6_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A6 + Offset, Size, Value
End Sub

Public Function A7() As String
   A7 = GetAlpha(MemPos_A7)
End Function
Public Function A7_OFS(ByVal Offset As Long) As String
   A7_OFS = GetAlpha(MemPos_A7 + Offset)
End Function
Public Sub LET_A7(ByVal Value As String)
   LetAlpha MemPos_A7, Value
End Sub
Public Sub LET_A7_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A7 + Offset, Value
End Sub
Public Sub SPOOL_A7(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A7, Size, Value
End Sub
Public Sub SPOOL_A7_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A7 + Offset, Size, Value
End Sub

Public Function A8() As String
   A8 = GetAlpha(MemPos_A8)
End Function
Public Function A8_OFS(ByVal Offset As Long) As String
   A8_OFS = GetAlpha(MemPos_A8 + Offset)
End Function
Public Sub LET_A8(ByVal Value As String)
   LetAlpha MemPos_A8, Value
End Sub
Public Sub LET_A8_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A8 + Offset, Value
End Sub
Public Sub SPOOL_A8(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A8, Size, Value
End Sub
Public Sub SPOOL_A8_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A8 + Offset, Size, Value
End Sub

Public Function A9() As String
   A9 = GetAlpha(MemPos_A9)
End Function
Public Function A9_OFS(ByVal Offset As Long) As String
   A9_OFS = GetAlpha(MemPos_A9 + Offset)
End Function
Public Sub LET_A9(ByVal Value As String)
   LetAlpha MemPos_A9, Value
End Sub
Public Sub LET_A9_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_A9 + Offset, Value
End Sub
Public Sub SPOOL_A9(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A9, Size, Value
End Sub
Public Sub SPOOL_A9_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_A9 + Offset, Size, Value
End Sub

' --- B ---

Public Function B() As String
   B = GetAlpha(MemPos_B)
End Function
Public Function B_OFS(ByVal Offset As Long) As String
   B_OFS = GetAlpha(MemPos_B + Offset)
End Function
Public Sub LET_B(ByVal Value As String)
   LetAlpha MemPos_B, Value
End Sub
Public Sub LET_B_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B + Offset, Value
End Sub
Public Sub SPOOL_B(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B, Size, Value
End Sub
Public Sub SPOOL_B_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B + Offset, Size, Value
End Sub

Public Function B1() As String
   B1 = GetAlpha(MemPos_B1)
End Function
Public Function B1_OFS(ByVal Offset As Long) As String
   B1_OFS = GetAlpha(MemPos_B1 + Offset)
End Function
Public Sub LET_B1(ByVal Value As String)
   LetAlpha MemPos_B1, Value
End Sub
Public Sub LET_B1_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B1 + Offset, Value
End Sub
Public Sub SPOOL_B1(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B1, Size, Value
End Sub
Public Sub SPOOL_B1_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B1 + Offset, Size, Value
End Sub

Public Function B2() As String
   B2 = GetAlpha(MemPos_B2)
End Function
Public Function B2_OFS(ByVal Offset As Long) As String
   B2_OFS = GetAlpha(MemPos_B2 + Offset)
End Function
Public Sub LET_B2(ByVal Value As String)
   LetAlpha MemPos_B2, Value
End Sub
Public Sub LET_B2_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B2 + Offset, Value
End Sub
Public Sub SPOOL_B2(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B2, Size, Value
End Sub
Public Sub SPOOL_B2_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B2 + Offset, Size, Value
End Sub

Public Function B3() As String
   B3 = GetAlpha(MemPos_B3)
End Function
Public Function B3_OFS(ByVal Offset As Long) As String
   B3_OFS = GetAlpha(MemPos_B3 + Offset)
End Function
Public Sub LET_B3(ByVal Value As String)
   LetAlpha MemPos_B3, Value
End Sub
Public Sub LET_B3_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B3 + Offset, Value
End Sub
Public Sub SPOOL_B3(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B3, Size, Value
End Sub
Public Sub SPOOL_B3_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B3 + Offset, Size, Value
End Sub

Public Function B4() As String
   B4 = GetAlpha(MemPos_B4)
End Function
Public Function B4_OFS(ByVal Offset As Long) As String
   B4_OFS = GetAlpha(MemPos_B4 + Offset)
End Function
Public Sub LET_B4(ByVal Value As String)
   LetAlpha MemPos_B4, Value
End Sub
Public Sub LET_B4_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B4 + Offset, Value
End Sub
Public Sub SPOOL_B4(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B4, Size, Value
End Sub
Public Sub SPOOL_B4_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B4 + Offset, Size, Value
End Sub

Public Function B5() As String
   B5 = GetAlpha(MemPos_B5)
End Function
Public Function B5_OFS(ByVal Offset As Long) As String
   B5_OFS = GetAlpha(MemPos_B5 + Offset)
End Function
Public Sub LET_B5(ByVal Value As String)
   LetAlpha MemPos_B5, Value
End Sub
Public Sub LET_B5_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B5 + Offset, Value
End Sub
Public Sub SPOOL_B5(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B5, Size, Value
End Sub
Public Sub SPOOL_B5_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B5 + Offset, Size, Value
End Sub

Public Function B6() As String
   B6 = GetAlpha(MemPos_B6)
End Function
Public Function B6_OFS(ByVal Offset As Long) As String
   B6_OFS = GetAlpha(MemPos_B6 + Offset)
End Function
Public Sub LET_B6(ByVal Value As String)
   LetAlpha MemPos_B6, Value
End Sub
Public Sub LET_B6_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B6 + Offset, Value
End Sub
Public Sub SPOOL_B6(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B6, Size, Value
End Sub
Public Sub SPOOL_B6_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B6 + Offset, Size, Value
End Sub

Public Function B7() As String
   B7 = GetAlpha(MemPos_B7)
End Function
Public Function B7_OFS(ByVal Offset As Long) As String
   B7_OFS = GetAlpha(MemPos_B7 + Offset)
End Function
Public Sub LET_B7(ByVal Value As String)
   LetAlpha MemPos_B7, Value
End Sub
Public Sub LET_B7_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B7 + Offset, Value
End Sub
Public Sub SPOOL_B7(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B7, Size, Value
End Sub
Public Sub SPOOL_B7_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B7 + Offset, Size, Value
End Sub

Public Function B8() As String
   B8 = GetAlpha(MemPos_B8)
End Function
Public Function B8_OFS(ByVal Offset As Long) As String
   B8_OFS = GetAlpha(MemPos_B8 + Offset)
End Function
Public Sub LET_B8(ByVal Value As String)
   LetAlpha MemPos_B8, Value
End Sub
Public Sub LET_B8_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B8 + Offset, Value
End Sub
Public Sub SPOOL_B8(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B8, Size, Value
End Sub
Public Sub SPOOL_B8_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B8 + Offset, Size, Value
End Sub

Public Function B9() As String
   B9 = GetAlpha(MemPos_B9)
End Function
Public Function B9_OFS(ByVal Offset As Long) As String
   B9_OFS = GetAlpha(MemPos_B9 + Offset)
End Function
Public Sub LET_B9(ByVal Value As String)
   LetAlpha MemPos_B9, Value
End Sub
Public Sub LET_B9_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_B9 + Offset, Value
End Sub
Public Sub SPOOL_B9(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B9, Size, Value
End Sub
Public Sub SPOOL_B9_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_B9 + Offset, Size, Value
End Sub

' --- C ---

Public Function C() As String
   C = GetAlpha(MemPos_C)
End Function
Public Function C_OFS(ByVal Offset As Long) As String
   C_OFS = GetAlpha(MemPos_C + Offset)
End Function
Public Sub LET_C(ByVal Value As String)
   LetAlpha MemPos_C, Value
End Sub
Public Sub LET_C_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C + Offset, Value
End Sub
Public Sub SPOOL_C(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C, Size, Value
End Sub
Public Sub SPOOL_C_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C + Offset, Size, Value
End Sub

Public Function C1() As String
   C1 = GetAlpha(MemPos_C1)
End Function
Public Function C1_OFS(ByVal Offset As Long) As String
   C1_OFS = GetAlpha(MemPos_C1 + Offset)
End Function
Public Sub LET_C1(ByVal Value As String)
   LetAlpha MemPos_C1, Value
End Sub
Public Sub LET_C1_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C1 + Offset, Value
End Sub
Public Sub SPOOL_C1(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C1, Size, Value
End Sub
Public Sub SPOOL_C1_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C1 + Offset, Size, Value
End Sub

Public Function C2() As String
   C2 = GetAlpha(MemPos_C2)
End Function
Public Function C2_OFS(ByVal Offset As Long) As String
   C2_OFS = GetAlpha(MemPos_C2 + Offset)
End Function
Public Sub LET_C2(ByVal Value As String)
   LetAlpha MemPos_C2, Value
End Sub
Public Sub LET_C2_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C2 + Offset, Value
End Sub
Public Sub SPOOL_C2(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C2, Size, Value
End Sub
Public Sub SPOOL_C2_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C2 + Offset, Size, Value
End Sub

Public Function C3() As String
   C3 = GetAlpha(MemPos_C3)
End Function
Public Function C3_OFS(ByVal Offset As Long) As String
   C3_OFS = GetAlpha(MemPos_C3 + Offset)
End Function
Public Sub LET_C3(ByVal Value As String)
   LetAlpha MemPos_C3, Value
End Sub
Public Sub LET_C3_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C3 + Offset, Value
End Sub
Public Sub SPOOL_C3(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C3, Size, Value
End Sub
Public Sub SPOOL_C3_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C3 + Offset, Size, Value
End Sub

Public Function C4() As String
   C4 = GetAlpha(MemPos_C4)
End Function
Public Function C4_OFS(ByVal Offset As Long) As String
   C4_OFS = GetAlpha(MemPos_C4 + Offset)
End Function
Public Sub LET_C4(ByVal Value As String)
   LetAlpha MemPos_C4, Value
End Sub
Public Sub LET_C4_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C4 + Offset, Value
End Sub
Public Sub SPOOL_C4(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C4, Size, Value
End Sub
Public Sub SPOOL_C4_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C4 + Offset, Size, Value
End Sub

Public Function C5() As String
   C5 = GetAlpha(MemPos_C5)
End Function
Public Function C5_OFS(ByVal Offset As Long) As String
   C5_OFS = GetAlpha(MemPos_C5 + Offset)
End Function
Public Sub LET_C5(ByVal Value As String)
   LetAlpha MemPos_C5, Value
End Sub
Public Sub LET_C5_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C5 + Offset, Value
End Sub
Public Sub SPOOL_C5(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C5, Size, Value
End Sub
Public Sub SPOOL_C5_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C5 + Offset, Size, Value
End Sub

Public Function C6() As String
   C6 = GetAlpha(MemPos_C6)
End Function
Public Function C6_OFS(ByVal Offset As Long) As String
   C6_OFS = GetAlpha(MemPos_C6 + Offset)
End Function
Public Sub LET_C6(ByVal Value As String)
   LetAlpha MemPos_C6, Value
End Sub
Public Sub LET_C6_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C6 + Offset, Value
End Sub
Public Sub SPOOL_C6(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C6, Size, Value
End Sub
Public Sub SPOOL_C6_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C6 + Offset, Size, Value
End Sub

Public Function C7() As String
   C7 = GetAlpha(MemPos_C7)
End Function
Public Function C7_OFS(ByVal Offset As Long) As String
   C7_OFS = GetAlpha(MemPos_C7 + Offset)
End Function
Public Sub LET_C7(ByVal Value As String)
   LetAlpha MemPos_C7, Value
End Sub
Public Sub LET_C7_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C7 + Offset, Value
End Sub
Public Sub SPOOL_C7(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C7, Size, Value
End Sub
Public Sub SPOOL_C7_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C7 + Offset, Size, Value
End Sub

Public Function C8() As String
   C8 = GetAlpha(MemPos_C8)
End Function
Public Function C8_OFS(ByVal Offset As Long) As String
   C8_OFS = GetAlpha(MemPos_C8 + Offset)
End Function
Public Sub LET_C8(ByVal Value As String)
   LetAlpha MemPos_C8, Value
End Sub
Public Sub LET_C8_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C8 + Offset, Value
End Sub
Public Sub SPOOL_C8(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C8, Size, Value
End Sub
Public Sub SPOOL_C8_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C8 + Offset, Size, Value
End Sub

Public Function C9() As String
   C9 = GetAlpha(MemPos_C9)
End Function
Public Function C9_OFS(ByVal Offset As Long) As String
   C9_OFS = GetAlpha(MemPos_C9 + Offset)
End Function
Public Sub LET_C9(ByVal Value As String)
   LetAlpha MemPos_C9, Value
End Sub
Public Sub LET_C9_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_C9 + Offset, Value
End Sub
Public Sub SPOOL_C9(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C9, Size, Value
End Sub
Public Sub SPOOL_C9_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_C9 + Offset, Size, Value
End Sub

' --- D ---

Public Function D() As String
   D = GetAlpha(MemPos_D)
End Function
Public Function D_OFS(ByVal Offset As Long) As String
   D_OFS = GetAlpha(MemPos_D + Offset)
End Function
Public Sub LET_D(ByVal Value As String)
   LetAlpha MemPos_D, Value
End Sub
Public Sub LET_D_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D + Offset, Value
End Sub
Public Sub SPOOL_D(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D, Size, Value
End Sub
Public Sub SPOOL_D_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D + Offset, Size, Value
End Sub

Public Function D1() As String
   D1 = GetAlpha(MemPos_D1)
End Function
Public Function D1_OFS(ByVal Offset As Long) As String
   D1_OFS = GetAlpha(MemPos_D1 + Offset)
End Function
Public Sub LET_D1(ByVal Value As String)
   LetAlpha MemPos_D1, Value
End Sub
Public Sub LET_D1_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D1 + Offset, Value
End Sub
Public Sub SPOOL_D1(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D1, Size, Value
End Sub
Public Sub SPOOL_D1_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D1 + Offset, Size, Value
End Sub

Public Function D2() As String
   D2 = GetAlpha(MemPos_D2)
End Function
Public Function D2_OFS(ByVal Offset As Long) As String
   D2_OFS = GetAlpha(MemPos_D2 + Offset)
End Function
Public Sub LET_D2(ByVal Value As String)
   LetAlpha MemPos_D2, Value
End Sub
Public Sub LET_D2_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D2 + Offset, Value
End Sub
Public Sub SPOOL_D2(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D2, Size, Value
End Sub
Public Sub SPOOL_D2_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D2 + Offset, Size, Value
End Sub

Public Function D3() As String
   D3 = GetAlpha(MemPos_D3)
End Function
Public Function D3_OFS(ByVal Offset As Long) As String
   D3_OFS = GetAlpha(MemPos_D3 + Offset)
End Function
Public Sub LET_D3(ByVal Value As String)
   LetAlpha MemPos_D3, Value
End Sub
Public Sub LET_D3_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D3 + Offset, Value
End Sub
Public Sub SPOOL_D3(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D3, Size, Value
End Sub
Public Sub SPOOL_D3_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D3 + Offset, Size, Value
End Sub

Public Function D4() As String
   D4 = GetAlpha(MemPos_D4)
End Function
Public Function D4_OFS(ByVal Offset As Long) As String
   D4_OFS = GetAlpha(MemPos_D4 + Offset)
End Function
Public Sub LET_D4(ByVal Value As String)
   LetAlpha MemPos_D4, Value
End Sub
Public Sub LET_D4_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D4 + Offset, Value
End Sub
Public Sub SPOOL_D4(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D4, Size, Value
End Sub
Public Sub SPOOL_D4_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D4 + Offset, Size, Value
End Sub

Public Function D5() As String
   D5 = GetAlpha(MemPos_D5)
End Function
Public Function D5_OFS(ByVal Offset As Long) As String
   D5_OFS = GetAlpha(MemPos_D5 + Offset)
End Function
Public Sub LET_D5(ByVal Value As String)
   LetAlpha MemPos_D5, Value
End Sub
Public Sub LET_D5_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D5 + Offset, Value
End Sub
Public Sub SPOOL_D5(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D5, Size, Value
End Sub
Public Sub SPOOL_D5_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D5 + Offset, Size, Value
End Sub

Public Function D6() As String
   D6 = GetAlpha(MemPos_D6)
End Function
Public Function D6_OFS(ByVal Offset As Long) As String
   D6_OFS = GetAlpha(MemPos_D6 + Offset)
End Function
Public Sub LET_D6(ByVal Value As String)
   LetAlpha MemPos_D6, Value
End Sub
Public Sub LET_D6_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D6 + Offset, Value
End Sub
Public Sub SPOOL_D6(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D6, Size, Value
End Sub
Public Sub SPOOL_D6_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D6 + Offset, Size, Value
End Sub

Public Function D7() As String
   D7 = GetAlpha(MemPos_D7)
End Function
Public Function D7_OFS(ByVal Offset As Long) As String
   D7_OFS = GetAlpha(MemPos_D7 + Offset)
End Function
Public Sub LET_D7(ByVal Value As String)
   LetAlpha MemPos_D7, Value
End Sub
Public Sub LET_D7_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D7 + Offset, Value
End Sub
Public Sub SPOOL_D7(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D7, Size, Value
End Sub
Public Sub SPOOL_D7_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D7 + Offset, Size, Value
End Sub

Public Function D8() As String
   D8 = GetAlpha(MemPos_D8)
End Function
Public Function D8_OFS(ByVal Offset As Long) As String
   D8_OFS = GetAlpha(MemPos_D8 + Offset)
End Function
Public Sub LET_D8(ByVal Value As String)
   LetAlpha MemPos_D8, Value
End Sub
Public Sub LET_D8_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D8 + Offset, Value
End Sub
Public Sub SPOOL_D8(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D8, Size, Value
End Sub
Public Sub SPOOL_D8_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D8 + Offset, Size, Value
End Sub

Public Function D9() As String
   D9 = GetAlpha(MemPos_D9)
End Function
Public Function D9_OFS(ByVal Offset As Long) As String
   D9_OFS = GetAlpha(MemPos_D9 + Offset)
End Function
Public Sub LET_D9(ByVal Value As String)
   LetAlpha MemPos_D9, Value
End Sub
Public Sub LET_D9_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_D9 + Offset, Value
End Sub
Public Sub SPOOL_D9(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D9, Size, Value
End Sub
Public Sub SPOOL_D9_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_D9 + Offset, Size, Value
End Sub

' --- E ---

Public Function E() As String
   E = GetAlpha(MemPos_E)
End Function
Public Function E_OFS(ByVal Offset As Long) As String
   E_OFS = GetAlpha(MemPos_E + Offset)
End Function
Public Sub LET_E(ByVal Value As String)
   LetAlpha MemPos_E, Value
End Sub
Public Sub LET_E_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E + Offset, Value
End Sub
Public Sub SPOOL_E(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E, Size, Value
End Sub
Public Sub SPOOL_E_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E + Offset, Size, Value
End Sub

Public Function E1() As String
   E1 = GetAlpha(MemPos_E1)
End Function
Public Function E1_OFS(ByVal Offset As Long) As String
   E1_OFS = GetAlpha(MemPos_E1 + Offset)
End Function
Public Sub LET_E1(ByVal Value As String)
   LetAlpha MemPos_E1, Value
End Sub
Public Sub LET_E1_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E1 + Offset, Value
End Sub
Public Sub SPOOL_E1(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E1, Size, Value
End Sub
Public Sub SPOOL_E1_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E1 + Offset, Size, Value
End Sub

Public Function E2() As String
   E2 = GetAlpha(MemPos_E2)
End Function
Public Function E2_OFS(ByVal Offset As Long) As String
   E2_OFS = GetAlpha(MemPos_E2 + Offset)
End Function
Public Sub LET_E2(ByVal Value As String)
   LetAlpha MemPos_E2, Value
End Sub
Public Sub LET_E2_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E2 + Offset, Value
End Sub
Public Sub SPOOL_E2(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E2, Size, Value
End Sub
Public Sub SPOOL_E2_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E2 + Offset, Size, Value
End Sub

Public Function E3() As String
   E3 = GetAlpha(MemPos_E3)
End Function
Public Function E3_OFS(ByVal Offset As Long) As String
   E3_OFS = GetAlpha(MemPos_E3 + Offset)
End Function
Public Sub LET_E3(ByVal Value As String)
   LetAlpha MemPos_E3, Value
End Sub
Public Sub LET_E3_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E3 + Offset, Value
End Sub
Public Sub SPOOL_E3(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E3, Size, Value
End Sub
Public Sub SPOOL_E3_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E3 + Offset, Size, Value
End Sub

Public Function E4() As String
   E4 = GetAlpha(MemPos_E4)
End Function
Public Function E4_OFS(ByVal Offset As Long) As String
   E4_OFS = GetAlpha(MemPos_E4 + Offset)
End Function
Public Sub LET_E4(ByVal Value As String)
   LetAlpha MemPos_E4, Value
End Sub
Public Sub LET_E4_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E4 + Offset, Value
End Sub
Public Sub SPOOL_E4(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E4, Size, Value
End Sub
Public Sub SPOOL_E4_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E4 + Offset, Size, Value
End Sub

Public Function E5() As String
   E5 = GetAlpha(MemPos_E5)
End Function
Public Function E5_OFS(ByVal Offset As Long) As String
   E5_OFS = GetAlpha(MemPos_E5 + Offset)
End Function
Public Sub LET_E5(ByVal Value As String)
   LetAlpha MemPos_E5, Value
End Sub
Public Sub LET_E5_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E5 + Offset, Value
End Sub
Public Sub SPOOL_E5(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E5, Size, Value
End Sub
Public Sub SPOOL_E5_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E5 + Offset, Size, Value
End Sub

Public Function E6() As String
   E6 = GetAlpha(MemPos_E6)
End Function
Public Function E6_OFS(ByVal Offset As Long) As String
   E6_OFS = GetAlpha(MemPos_E6 + Offset)
End Function
Public Sub LET_E6(ByVal Value As String)
   LetAlpha MemPos_E6, Value
End Sub
Public Sub LET_E6_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E6 + Offset, Value
End Sub
Public Sub SPOOL_E6(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E6, Size, Value
End Sub
Public Sub SPOOL_E6_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E6 + Offset, Size, Value
End Sub

Public Function E7() As String
   E7 = GetAlpha(MemPos_E7)
End Function
Public Function E7_OFS(ByVal Offset As Long) As String
   E7_OFS = GetAlpha(MemPos_E7 + Offset)
End Function
Public Sub LET_E7(ByVal Value As String)
   LetAlpha MemPos_E7, Value
End Sub
Public Sub LET_E7_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E7 + Offset, Value
End Sub
Public Sub SPOOL_E7(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E7, Size, Value
End Sub
Public Sub SPOOL_E7_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E7 + Offset, Size, Value
End Sub

Public Function E8() As String
   E8 = GetAlpha(MemPos_E8)
End Function
Public Function E8_OFS(ByVal Offset As Long) As String
   E8_OFS = GetAlpha(MemPos_E8 + Offset)
End Function
Public Sub LET_E8(ByVal Value As String)
   LetAlpha MemPos_E8, Value
End Sub
Public Sub LET_E8_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E8 + Offset, Value
End Sub
Public Sub SPOOL_E8(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E8, Size, Value
End Sub
Public Sub SPOOL_E8_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E8 + Offset, Size, Value
End Sub

Public Function E9() As String
   E9 = GetAlpha(MemPos_E9)
End Function
Public Function E9_OFS(ByVal Offset As Long) As String
   E9_OFS = GetAlpha(MemPos_E9 + Offset)
End Function
Public Sub LET_E9(ByVal Value As String)
   LetAlpha MemPos_E9, Value
End Sub
Public Sub LET_E9_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlpha MemPos_E9 + Offset, Value
End Sub
Public Sub SPOOL_E9(ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E9, Size, Value
End Sub
Public Sub SPOOL_E9_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlpha MemPos_E9 + Offset, Size, Value
End Sub

' ---------------
' --- Buffers ---
' ---------------

' --- R ---

Public Function R(ByVal Size As Long) As Currency
   R = GetNumericBuffer(MemPos_RP, 0, Size)
End Function
Public Function R_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   R_OFS = GetNumericBuffer(MemPos_RP, Offset, Size)
End Function
Public Function R_A() As String
   R_A = GetAlphaBuffer(MemPos_RP, 0)
End Function
Public Function R_A_OFS(ByVal Offset As Long) As String
   R_A_OFS = GetAlphaBuffer(MemPos_RP, Offset)
End Function
Public Sub LET_R(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_RP, 0, Size, Value
End Sub
Public Sub LET_R_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_RP, Offset, Size, Value
End Sub
Public Sub LET_R_A(ByVal Value As String)
   LetAlphaBuffer MemPos_RP, 0, Value
End Sub
Public Sub LET_R_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_RP, Offset, Value
End Sub
Public Sub SPOOL_R_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_RP, 0, Size, Value
End Sub
Public Sub SPOOL_R_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_RP, Offset, Size, Value
End Sub

' --- Z ---

Public Function Z(ByVal Size As Long) As Currency
   Z = GetNumericBuffer(MemPos_ZP, 0, Size)
End Function
Public Function Z_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   Z_OFS = GetNumericBuffer(MemPos_ZP, Offset, Size)
End Function
Public Function Z_A() As String
   Z_A = GetAlphaBuffer(MemPos_ZP, 0)
End Function
Public Function Z_A_OFS(ByVal Offset As Long) As String
   Z_A_OFS = GetAlphaBuffer(MemPos_ZP, Offset)
End Function
Public Sub LET_Z(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_ZP, 0, Size, Value
End Sub
Public Sub LET_Z_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_ZP, Offset, Size, Value
End Sub
Public Sub LET_Z_A(ByVal Value As String)
   LetAlphaBuffer MemPos_ZP, 0, Value
End Sub
Public Sub LET_Z_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_ZP, Offset, Value
End Sub
Public Sub SPOOL_Z_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_ZP, 0, Size, Value
End Sub
Public Sub SPOOL_Z_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_ZP, Offset, Size, Value
End Sub

' --- X ---

Public Function X(ByVal Size As Long) As Currency
   X = GetNumericBuffer(MemPos_XP, 0, Size)
End Function
Public Function X_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   X_OFS = GetNumericBuffer(MemPos_XP, Offset, Size)
End Function
Public Function X_A() As String
   X_A = GetAlphaBuffer(MemPos_XP, 0)
End Function
Public Function X_A_OFS(ByVal Offset As Long) As String
   X_A_OFS = GetAlphaBuffer(MemPos_XP, Offset)
End Function
Public Sub LET_X(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_XP, 0, Size, Value
End Sub
Public Sub LET_X_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_XP, Offset, Size, Value
End Sub
Public Sub LET_X_A(ByVal Value As String)
   LetAlphaBuffer MemPos_XP, 0, Value
End Sub
Public Sub LET_X_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_XP, Offset, Value
End Sub
Public Sub SPOOL_X_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_XP, 0, Size, Value
End Sub
Public Sub SPOOL_X_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_XP, Offset, Size, Value
End Sub

' --- Y ---

Public Function Y(ByVal Size As Long) As Currency
   Y = GetNumericBuffer(MemPos_YP, 0, Size)
End Function
Public Function Y_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   Y_OFS = GetNumericBuffer(MemPos_YP, Offset, Size)
End Function
Public Function Y_A() As String
   Y_A = GetAlphaBuffer(MemPos_YP, 0)
End Function
Public Function Y_A_OFS(ByVal Offset As Long) As String
   Y_A_OFS = GetAlphaBuffer(MemPos_YP, Offset)
End Function
Public Sub LET_Y(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_YP, 0, Size, Value
End Sub
Public Sub LET_Y_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_YP, Offset, Size, Value
End Sub
Public Sub LET_Y_A(ByVal Value As String)
   LetAlphaBuffer MemPos_YP, 0, Value
End Sub
Public Sub LET_Y_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_YP, Offset, Value
End Sub
Public Sub SPOOL_Y_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_YP, 0, Size, Value
End Sub
Public Sub SPOOL_Y_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_YP, Offset, Size, Value
End Sub

' --- W ---

Public Function W(ByVal Size As Long) As Currency
   W = GetNumericBuffer(MemPos_WP, 0, Size)
End Function
Public Function W_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   W_OFS = GetNumericBuffer(MemPos_WP, Offset, Size)
End Function
Public Function W_A() As String
   W_A = GetAlphaBuffer(MemPos_WP, 0)
End Function
Public Function W_A_OFS(ByVal Offset As Long) As String
   W_A_OFS = GetAlphaBuffer(MemPos_WP, Offset)
End Function
Public Sub LET_W(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_WP, 0, Size, Value
End Sub
Public Sub LET_W_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_WP, Offset, Size, Value
End Sub
Public Sub LET_W_A(ByVal Value As String)
   LetAlphaBuffer MemPos_WP, 0, Value
End Sub
Public Sub LET_W_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_WP, Offset, Value
End Sub
Public Sub SPOOL_W_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_WP, 0, Size, Value
End Sub
Public Sub SPOOL_W_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_WP, Offset, Size, Value
End Sub

' --- S ---

Public Function S(ByVal Size As Long) As Currency
   S = GetNumericBuffer(MemPos_SP, 0, Size)
End Function
Public Function S_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   S_OFS = GetNumericBuffer(MemPos_SP, Offset, Size)
End Function
Public Function S_A() As String
   S_A = GetAlphaBuffer(MemPos_SP, 0)
End Function
Public Function S_A_OFS(ByVal Offset As Long) As String
   S_A_OFS = GetAlphaBuffer(MemPos_SP, Offset)
End Function
Public Sub LET_S(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_SP, 0, Size, Value
End Sub
Public Sub LET_S_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_SP, Offset, Size, Value
End Sub
Public Sub LET_S_A(ByVal Value As String)
   LetAlphaBuffer MemPos_SP, 0, Value
End Sub
Public Sub LET_S_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_SP, Offset, Value
End Sub
Public Sub SPOOL_S_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_SP, 0, Size, Value
End Sub
Public Sub SPOOL_S_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_SP, Offset, Size, Value
End Sub

' --- T ---

Public Function T(ByVal Size As Long) As Currency
   T = GetNumericBuffer(MemPos_TP, 0, Size)
End Function
Public Function T_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   T_OFS = GetNumericBuffer(MemPos_TP, Offset, Size)
End Function
Public Function T_A() As String
   T_A = GetAlphaBuffer(MemPos_TP, 0)
End Function
Public Function T_A_OFS(ByVal Offset As Long) As String
   T_A_OFS = GetAlphaBuffer(MemPos_TP, Offset)
End Function
Public Sub LET_T(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_TP, 0, Size, Value
End Sub
Public Sub LET_T_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_TP, Offset, Size, Value
End Sub
Public Sub LET_T_A(ByVal Value As String)
   LetAlphaBuffer MemPos_TP, 0, Value
End Sub
Public Sub LET_T_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_TP, Offset, Value
End Sub
Public Sub SPOOL_T_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_TP, 0, Size, Value
End Sub
Public Sub SPOOL_T_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_TP, Offset, Size, Value
End Sub

' --- U ---

Public Function U(ByVal Size As Long) As Currency
   U = GetNumericBuffer(MemPos_UP, 0, Size)
End Function
Public Function U_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   U_OFS = GetNumericBuffer(MemPos_UP, Offset, Size)
End Function
Public Function U_A() As String
   U_A = GetAlphaBuffer(MemPos_UP, 0)
End Function
Public Function U_A_OFS(ByVal Offset As Long) As String
   U_A_OFS = GetAlphaBuffer(MemPos_UP, Offset)
End Function
Public Sub LET_U(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_UP, 0, Size, Value
End Sub
Public Sub LET_U_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_UP, Offset, Size, Value
End Sub
Public Sub LET_U_A(ByVal Value As String)
   LetAlphaBuffer MemPos_UP, 0, Value
End Sub
Public Sub LET_U_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_UP, Offset, Value
End Sub
Public Sub SPOOL_U_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_UP, 0, Size, Value
End Sub
Public Sub SPOOL_U_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_UP, Offset, Size, Value
End Sub

' --- V ---

Public Function V(ByVal Size As Long) As Currency
   V = GetNumericBuffer(MemPos_VP, 0, Size)
End Function
Public Function V_OFS(ByVal Offset As Long, ByVal Size As Long) As Currency
   V_OFS = GetNumericBuffer(MemPos_VP, Offset, Size)
End Function
Public Function V_A() As String
   V_A = GetAlphaBuffer(MemPos_VP, 0)
End Function
Public Function V_A_OFS(ByVal Offset As Long) As String
   V_A_OFS = GetAlphaBuffer(MemPos_VP, Offset)
End Function
Public Sub LET_V(ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_VP, 0, Size, Value
End Sub
Public Sub LET_V_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   LetNumericBuffer MemPos_VP, Offset, Size, Value
End Sub
Public Sub LET_V_A(ByVal Value As String)
   LetAlphaBuffer MemPos_VP, 0, Value
End Sub
Public Sub LET_V_A_OFS(ByVal Offset As Long, ByVal Value As String)
   LetAlphaBuffer MemPos_VP, Offset, Value
End Sub
Public Sub SPOOL_V_A(ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_VP, 0, Size, Value
End Sub
Public Sub SPOOL_V_A_OFS(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   SpoolAlphaBuffer MemPos_VP, Offset, Size, Value
End Sub

' -------------------------------
' --- Page 0 System Variables ---
' -------------------------------

Public Function RP() As Currency
   RP = MEM(MemPos_RP)
End Function
Public Sub LET_RP(ByVal Value As Currency)
   LetByte MemPos_RP, Value
End Sub
Public Function RP2() As Currency
   RP2 = MEM(MemPos_RP2)
End Function
Public Sub LET_RP2(ByVal Value As Currency)
   LetByte MemPos_RP2, Value
End Sub
Public Function IRP() As Currency
   IRP = MEM(MemPos_IRP)
End Function
Public Sub LET_IRP(ByVal Value As Currency)
   LetByte MemPos_IRP, Value
End Sub
Public Function IRP2() As Currency
   IRP2 = MEM(MemPos_IRP2)
End Function
Public Sub LET_IRP2(ByVal Value As Currency)
   LetByte MemPos_IRP2, Value
End Sub

Public Function ZP() As Currency
   ZP = MEM(MemPos_ZP)
End Function
Public Sub LET_ZP(ByVal Value As Currency)
   LetByte MemPos_ZP, Value
End Sub
Public Function ZP2() As Currency
   ZP2 = MEM(MemPos_ZP2)
End Function
Public Sub LET_ZP2(ByVal Value As Currency)
   LetByte MemPos_ZP2, Value
End Sub
Public Function IZP() As Currency
   IZP = MEM(MemPos_IZP)
End Function
Public Sub LET_IZP(ByVal Value As Currency)
   LetByte MemPos_IZP, Value
End Sub
Public Function IZP2() As Currency
   IZP2 = MEM(MemPos_IZP2)
End Function
Public Sub LET_IZP2(ByVal Value As Currency)
   LetByte MemPos_IZP2, Value
End Sub

Public Function XP() As Currency
   XP = MEM(MemPos_XP)
End Function
Public Sub LET_XP(ByVal Value As Currency)
   LetByte MemPos_XP, Value
End Sub
Public Function XP2() As Currency
   XP2 = MEM(MemPos_XP2)
End Function
Public Sub LET_XP2(ByVal Value As Currency)
   LetByte MemPos_XP2, Value
End Sub
Public Function IXP() As Currency
   IXP = MEM(MemPos_IXP)
End Function
Public Sub LET_IXP(ByVal Value As Currency)
   LetByte MemPos_IXP, Value
End Sub
Public Function IXP2() As Currency
   IXP2 = MEM(MemPos_IXP2)
End Function
Public Sub LET_IXP2(ByVal Value As Currency)
   LetByte MemPos_IXP2, Value
End Sub

Public Function YP() As Currency
   YP = MEM(MemPos_YP)
End Function
Public Sub LET_YP(ByVal Value As Currency)
   LetByte MemPos_YP, Value
End Sub
Public Function YP2() As Currency
   YP2 = MEM(MemPos_YP2)
End Function
Public Sub LET_YP2(ByVal Value As Currency)
   LetByte MemPos_YP2, Value
End Sub
Public Function IYP() As Currency
   IYP = MEM(MemPos_IYP)
End Function
Public Sub LET_IYP(ByVal Value As Currency)
   LetByte MemPos_IYP, Value
End Sub
Public Function IYP2() As Currency
   IYP2 = MEM(MemPos_IYP2)
End Function
Public Sub LET_IYP2(ByVal Value As Currency)
   LetByte MemPos_IYP2, Value
End Sub

Public Function WP() As Currency
   WP = MEM(MemPos_WP)
End Function
Public Sub LET_WP(ByVal Value As Currency)
   LetByte MemPos_WP, Value
End Sub
Public Function WP2() As Currency
   WP2 = MEM(MemPos_WP2)
End Function
Public Sub LET_WP2(ByVal Value As Currency)
   LetByte MemPos_WP2, Value
End Sub
Public Function IWP() As Currency
   IWP = MEM(MemPos_IWP)
End Function
Public Sub LET_IWP(ByVal Value As Currency)
   LetByte MemPos_IWP, Value
End Sub
Public Function IWP2() As Currency
   IWP2 = MEM(MemPos_IWP2)
End Function
Public Sub LET_IWP2(ByVal Value As Currency)
   LetByte MemPos_IWP2, Value
End Sub

Public Function SP() As Currency
   SP = MEM(MemPos_SP)
End Function
Public Sub LET_SP(ByVal Value As Currency)
   LetByte MemPos_SP, Value
End Sub
Public Function SP2() As Currency
   SP2 = MEM(MemPos_SP2)
End Function
Public Sub LET_SP2(ByVal Value As Currency)
   LetByte MemPos_SP2, Value
End Sub
Public Function ISP() As Currency
   ISP = MEM(MemPos_ISP)
End Function
Public Sub LET_ISP(ByVal Value As Currency)
   LetByte MemPos_ISP, Value
End Sub
Public Function ISP2() As Currency
   ISP2 = MEM(MemPos_ISP2)
End Function
Public Sub LET_ISP2(ByVal Value As Currency)
   LetByte MemPos_ISP2, Value
End Sub

Public Function TP() As Currency
   TP = MEM(MemPos_TP)
End Function
Public Sub LET_TP(ByVal Value As Currency)
   LetByte MemPos_TP, Value
End Sub
Public Function TP2() As Currency
   TP2 = MEM(MemPos_TP2)
End Function
Public Sub LET_TP2(ByVal Value As Currency)
   LetByte MemPos_TP2, Value
End Sub
Public Function ITP() As Currency
   ITP = MEM(MemPos_ITP)
End Function
Public Sub LET_ITP(ByVal Value As Currency)
   LetByte MemPos_ITP, Value
End Sub
Public Function ITP2() As Currency
   ITP2 = MEM(MemPos_ITP2)
End Function
Public Sub LET_ITP2(ByVal Value As Currency)
   LetByte MemPos_ITP2, Value
End Sub

Public Function UP() As Currency
   UP = MEM(MemPos_UP)
End Function
Public Sub LET_UP(ByVal Value As Currency)
   LetByte MemPos_UP, Value
End Sub
Public Function UP2() As Currency
   UP2 = MEM(MemPos_UP2)
End Function
Public Sub LET_UP2(ByVal Value As Currency)
   LetByte MemPos_UP2, Value
End Sub
Public Function IUP() As Currency
   IUP = MEM(MemPos_IUP)
End Function
Public Sub LET_IUP(ByVal Value As Currency)
   LetByte MemPos_IUP, Value
End Sub
Public Function IUP2() As Currency
   IUP2 = MEM(MemPos_IUP2)
End Function
Public Sub LET_IUP2(ByVal Value As Currency)
   LetByte MemPos_IUP2, Value
End Sub

Public Function VP() As Currency
   VP = MEM(MemPos_VP)
End Function
Public Sub LET_VP(ByVal Value As Currency)
   LetByte MemPos_VP, Value
End Sub
Public Function VP2() As Currency
   VP2 = MEM(MemPos_VP2)
End Function
Public Sub LET_VP2(ByVal Value As Currency)
   LetByte MemPos_VP2, Value
End Sub
Public Function IVP() As Currency
   IVP = MEM(MemPos_IVP)
End Function
Public Sub LET_IVP(ByVal Value As Currency)
   LetByte MemPos_IVP, Value
End Sub
Public Function IVP2() As Currency
   IVP2 = MEM(MemPos_IVP2)
End Function
Public Sub LET_IVP2(ByVal Value As Currency)
   LetByte MemPos_IVP2, Value
End Sub

' --- System variables ---

Public Function LIB() As Currency
   LIB = MEM(MemPos_Lib)
End Function
Public Sub LET_LIB(ByVal Value As Currency)
   LetByte MemPos_Lib, Value
End Sub

Public Function PROG() As Long
   PROG = GetNumeric(MemPos_Prog, 2)
End Function
Public Sub LET_PROG(ByVal Value As Currency)
   If Value < -256 Or Value > 255 Then
      ThrowError "LET_PROG", "Invalid Program Number: " & Trim$(Str$(Value))
      Exit Sub
   End If
   LetNumeric MemPos_Prog, 2, Value
End Sub

Public Function PRIVG() As Currency
   PRIVG = MEM(MemPos_Privg)
End Function
Public Sub LET_PRIVG(ByVal Value As Currency)
   LetByte MemPos_Privg, Value
End Sub

Public Function CHAR() As Currency
   CHAR = MEM(MemPos_Char)
End Function
Public Sub LET_CHAR(ByVal Value As Currency)
   LetByte MemPos_Char, Value
End Sub

Public Function LENGTH() As Currency
   LENGTH = MEM(MemPos_Length)
End Function
Public Sub LET_LENGTH(ByVal Value As Currency)
   LetByte MemPos_Length, Value
End Sub

Public Function STATUS() As Currency
   STATUS = MEM(MemPos_Status)
End Function
Public Sub LET_STATUS(ByVal Value As Currency)
   LetByte MemPos_Status, Value
End Sub

Public Function ESCVAL() As Currency
   If EXITING Or MEMTF(MemPos_Background) Then
      ESCVAL = 0 ' force normal escape
   Else
      ESCVAL = MEM(MemPos_EscVal)
   End If
End Function
Public Sub LET_ESCVAL(ByVal Value As Currency)
   LetByte MemPos_EscVal, Value
End Sub

Public Function CANVAL() As Currency
   CANVAL = MEM(MemPos_CanVal)
End Function
Public Sub LET_CANVAL(ByVal Value As Currency)
   LetByte MemPos_CanVal, Value
End Sub

Public Function LOCKVAL() As Currency
   LOCKVAL = MEM(MemPos_LockVal)
End Function
Public Sub LET_LOCKVAL(ByVal Value As Currency)
   LetByte MemPos_LockVal, Value
End Sub

Public Function TCHAN() As Currency
   TCHAN = MEM(MemPos_TChan)
End Function
Public Sub LET_TCHAN(ByVal Value As Currency)
   LetByte MemPos_TChan, Value
End Sub

Public Function TERM() As Currency
   TERM = MEM(MemPos_Term)
End Function
Public Sub LET_TERM(ByVal Value As Currency)
   LetByte MemPos_Term, Value
End Sub

Public Function LANG() As Currency
   LANG = MEM(MemPos_Lang)
End Function
Public Sub LET_LANG(ByVal Value As Currency)
   LetByte MemPos_Lang, Value
End Sub

Public Function PRTNUM() As Currency
   PRTNUM = MEM(MemPos_PrtNum)
End Function
Public Sub LET_PRTNUM(ByVal Value As Currency)
   LetByte MemPos_PrtNum, Value
End Sub

Public Function TFA() As Currency
   TFA = MEM(MemPos_TFA)
End Function
Public Sub LET_TFA(ByVal Value As Currency)
   LetByte MemPos_TFA, Value
End Sub

Public Function VOL() As Currency
   VOL = MEM(MemPos_Vol)
End Function
Public Sub LET_VOL(ByVal Value As Currency)
   LetByte MemPos_Vol, Value
End Sub

Public Function PVOL() As Currency
   PVOL = MEM(MemPos_PVol)
End Function
Public Sub LET_PVOL(ByVal Value As Currency)
   LetByte MemPos_PVol, Value
End Sub

Public Function REQVOL() As Currency
   REQVOL = MEM(MemPos_ReqVol)
End Function
Public Sub LET_REQVOL(ByVal Value As Currency)
   LetByte MemPos_ReqVol, Value
End Sub

Public Function KBC() As Currency
   Dim bChar As Byte
   Dim Result As Currency
   ' --------------------
   If EXITING Then GoTo DoExit
   ' --- allow any pending messages to be processed ---
   If KBuff.Count = 0 Then
      Sleep 1 ' prevent high cpu usage
      DoEvents ' might have pending keystroke messages
      If EXITING Then GoTo DoExit
   End If
   ' --- process special characters now ---
NextSpecChar:
   If KBuff.Count > 1 Then
      If KBuff.Peek = 0 Then
         bChar = KBuff.GetChar ' zero
         bChar = KBuff.GetChar ' special character
         ' --- acknowledge characters used ---
         SendToServer "KEYBOARD" & vbTab & "USED" & vbTab & HexChar(0) & HexChar(bChar)
         ' --- handle special chars ---
         HandleSpecialChar bChar
         GoTo NextSpecChar
      End If
   End If
   ' --- get number of characters ---
   Result = KBuff.Count
   If Result > 254 Then Result = 254 ' max allowed
   ' --- done ---
   KBC = Result
   Exit Function
DoExit:
   ' --- When exiting, pretend KBuff has an Escape char.  ---
   ' --- The function GetKeyboardChar returns the Escape. ---
   LET_ESCVAL 0 ' force normal escapes
   KBC = 1 ' always one char in buffer
End Function
Public Function KBCX() As Currency
   ' --- this functions like KBC, but always returns zero if
   ' --- a keyboard script is running. It still processes any
   ' --- special characters and such. KBC is still needed as
   ' --- it is for items such as Policy Menus and such. Use
   ' --- this instead of KBC for any checking for Aborting.
   ' ---     IF KBCX#0 THEN ... ! ABORTING
   Dim bChar As Byte
   Dim Result As Currency
   ' --------------------
   If EXITING Then GoTo DoExit
   ' --- allow any pending messages to be processed ---
   If KBuff.Count = 0 Then
      Sleep 1 ' prevent high cpu usage
      DoEvents ' might have pending keystroke messages
      If EXITING Then GoTo DoExit
   End If
   ' --- process special characters now ---
NextSpecChar:
   If KBuff.Count > 1 Then
      If KBuff.Peek = 0 Then
         bChar = KBuff.GetChar ' zero
         bChar = KBuff.GetChar ' special character
         ' --- acknowledge characters used ---
         SendToServer "KEYBOARD" & vbTab & "USED" & vbTab & HexChar(0) & HexChar(bChar)
         ' --- handle special chars ---
         HandleSpecialChar bChar
         GoTo NextSpecChar
      End If
   End If
   ' --- get number of characters ---
   Result = KBuff.Count
   If Result > 254 Then Result = 254 ' max allowed
   ' --- always return a zero if running a script ---
   If MEMTF(MemPos_ScriptRunFlag) Then
      Result = 0 ' don't trigger an abort
   End If
   ' --- done ---
   KBCX = Result
   Exit Function
DoExit:
   ' --- When exiting, pretend KBuff has an Escape char.  ---
   ' --- The function GetKeyboardChar returns the Escape. ---
   LET_ESCVAL 0 ' force normal escapes
   KBCX = 1 ' always one char in buffer
End Function
Public Sub LET_KBC(ByVal Value As Currency)
   Dim bChar As Byte
   Dim bLastChar As Byte
   Dim lngCount As Long
   Dim lngRemoved As Long
   ' --------------------
   ' --- the only value allowed is Zero ---
   If Value <> 0 Then
      ThrowError "LET_KBC", "Cannot set KBC to a non-zero value"
      Exit Sub
   End If
   ' --- check if exiting ---
   If EXITING Then
      LET_ESCVAL 0 ' force normal escapes
      ' --- don't actually remove characters ---
      Exit Sub
   End If
   ' --- Don't clear if running a script. ---
   If MEMTF(MemPos_ScriptRunFlag) Then Exit Sub
   ' --- get number of chars currently in buffer ---
   lngCount = KBuff.Count
   If lngCount = 0 Then Exit Sub
   ' --- remove chars from buffer ---
   lngRemoved = 0
   bChar = 255
   Do While lngCount > 0
      bLastChar = bChar
      bChar = KBuff.Peek ' look but don't remove
      lngCount = lngCount - 1
      ' --- don't remove the last char if it is Zero ---
      If bChar <> 0 Or lngCount > 0 Then
         bChar = KBuff.GetChar ' actually remove it
         lngRemoved = lngRemoved + 1
         ' --- handle special characters ---
         If bLastChar = 0 Then
            HandleSpecialChar bChar
         End If
      End If
   Loop
   ' --- send clear message to client ---
   If lngRemoved > 0 Then
      SendToServer "KEYBOARD" & vbTab & "CLEARNUM" & vbTab & Trim$(Str$(lngRemoved))
   End If
End Sub
Public Sub LET_KBCX(ByVal Value As Currency)
   LET_KBC Value
End Sub

Public Function OPER() As Long
   OPER = GetNumeric(MemPos_Oper, 2)
End Function
Public Sub LET_OPER(ByVal Value As Currency)
   LetNumeric MemPos_Oper, 2, Value
End Sub

' --- Read-only values, so no LET_xxx is defined ---

Public Function USER() As Long
   USER = GetNumeric(MemPos_User, 2)
End Function

Public Function ORIG() As Long
   ORIG = GetNumeric(MemPos_Orig, 2)
End Function

Public Function MACHTYPE() As Currency
   MACHTYPE = MEM(MemPos_MachType)
End Function

Public Function SYSREL() As Currency
   SYSREL = MEM(MemPos_SysRel)
End Function

Public Function SYSREV() As Currency
   SYSREV = MEM(MemPos_SysRev)
End Function

' --- Flag registers ---

Public Function F() As Currency
   F = MEM(MemPos_F)
End Function
Public Sub LET_F(ByVal Value As Currency)
   LetByte MemPos_F, Value
End Sub

Public Function F1() As Currency
   F1 = MEM(MemPos_F1)
End Function
Public Sub LET_F1(ByVal Value As Currency)
   LetByte MemPos_F1, Value
End Sub

Public Function F2() As Currency
   F2 = MEM(MemPos_F2)
End Function
Public Sub LET_F2(ByVal Value As Currency)
   LetByte MemPos_F2, Value
End Sub

Public Function F3() As Currency
   F3 = MEM(MemPos_F3)
End Function
Public Sub LET_F3(ByVal Value As Currency)
   LetByte MemPos_F3, Value
End Sub

Public Function F4() As Currency
   F4 = MEM(MemPos_F4)
End Function
Public Sub LET_F4(ByVal Value As Currency)
   LetByte MemPos_F4, Value
End Sub

Public Function F5() As Currency
   F5 = MEM(MemPos_F5)
End Function
Public Sub LET_F5(ByVal Value As Currency)
   LetByte MemPos_F5, Value
End Sub

Public Function F6() As Currency
   F6 = MEM(MemPos_F6)
End Function
Public Sub LET_F6(ByVal Value As Currency)
   LetByte MemPos_F6, Value
End Sub

Public Function F7() As Currency
   F7 = MEM(MemPos_F7)
End Function
Public Sub LET_F7(ByVal Value As Currency)
   LetByte MemPos_F7, Value
End Sub

Public Function F8() As Currency
   F8 = MEM(MemPos_F8)
End Function
Public Sub LET_F8(ByVal Value As Currency)
   LetByte MemPos_F8, Value
End Sub

Public Function F9() As Currency
   F9 = MEM(MemPos_F9)
End Function
Public Sub LET_F9(ByVal Value As Currency)
   LetByte MemPos_F9, Value
End Sub

Public Function F10() As Currency
   F10 = MEM(MemPos_F10)
End Function
Public Sub LET_F10(ByVal Value As Currency)
   LetByte MemPos_F10, Value
End Sub

Public Function F11() As Currency
   F11 = MEM(MemPos_F11)
End Function
Public Sub LET_F11(ByVal Value As Currency)
   LetByte MemPos_F11, Value
End Sub

Public Function F12() As Currency
   F12 = MEM(MemPos_F12)
End Function
Public Sub LET_F12(ByVal Value As Currency)
   LetByte MemPos_F12, Value
End Sub

Public Function F13() As Currency
   F13 = MEM(MemPos_F13)
End Function
Public Sub LET_F13(ByVal Value As Currency)
   LetByte MemPos_F13, Value
End Sub

Public Function F14() As Currency
   F14 = MEM(MemPos_F14)
End Function
Public Sub LET_F14(ByVal Value As Currency)
   LetByte MemPos_F14, Value
End Sub

Public Function F15() As Currency
   F15 = MEM(MemPos_F15)
End Function
Public Sub LET_F15(ByVal Value As Currency)
   LetByte MemPos_F15, Value
End Sub

Public Function F16() As Currency
   F16 = MEM(MemPos_F16)
End Function
Public Sub LET_F16(ByVal Value As Currency)
   LetByte MemPos_F16, Value
End Sub

Public Function F17() As Currency
   F17 = MEM(MemPos_F17)
End Function
Public Sub LET_F17(ByVal Value As Currency)
   LetByte MemPos_F17, Value
End Sub

Public Function F18() As Currency
   F18 = MEM(MemPos_F18)
End Function
Public Sub LET_F18(ByVal Value As Currency)
   LetByte MemPos_F18, Value
End Sub

Public Function F19() As Currency
   F19 = MEM(MemPos_F19)
End Function
Public Sub LET_F19(ByVal Value As Currency)
   LetByte MemPos_F19, Value
End Sub

Public Function F20() As Currency
   F20 = MEM(MemPos_F20)
End Function
Public Sub LET_F20(ByVal Value As Currency)
   LetByte MemPos_F20, Value
End Sub

Public Function F21() As Currency
   F21 = MEM(MemPos_F21)
End Function
Public Sub LET_F21(ByVal Value As Currency)
   LetByte MemPos_F21, Value
End Sub

Public Function F22() As Currency
   F22 = MEM(MemPos_F22)
End Function
Public Sub LET_F22(ByVal Value As Currency)
   LetByte MemPos_F22, Value
End Sub

Public Function F23() As Currency
   F23 = MEM(MemPos_F23)
End Function
Public Sub LET_F23(ByVal Value As Currency)
   LetByte MemPos_F23, Value
End Sub

Public Function F24() As Currency
   F24 = MEM(MemPos_F24)
End Function
Public Sub LET_F24(ByVal Value As Currency)
   LetByte MemPos_F24, Value
End Sub

Public Function F25() As Currency
   F25 = MEM(MemPos_F25)
End Function
Public Sub LET_F25(ByVal Value As Currency)
   LetByte MemPos_F25, Value
End Sub

Public Function F26() As Currency
   F26 = MEM(MemPos_F26)
End Function
Public Sub LET_F26(ByVal Value As Currency)
   LetByte MemPos_F26, Value
End Sub

Public Function F27() As Currency
   F27 = MEM(MemPos_F27)
End Function
Public Sub LET_F27(ByVal Value As Currency)
   LetByte MemPos_F27, Value
End Sub

Public Function F28() As Currency
   F28 = MEM(MemPos_F28)
End Function
Public Sub LET_F28(ByVal Value As Currency)
   LetByte MemPos_F28, Value
End Sub

Public Function F29() As Currency
   F29 = MEM(MemPos_F29)
End Function
Public Sub LET_F29(ByVal Value As Currency)
   LetByte MemPos_F29, Value
End Sub

Public Function F30() As Currency
   F30 = MEM(MemPos_F30)
End Function
Public Sub LET_F30(ByVal Value As Currency)
   LetByte MemPos_F30, Value
End Sub

Public Function F31() As Currency
   F31 = MEM(MemPos_F31)
End Function
Public Sub LET_F31(ByVal Value As Currency)
   LetByte MemPos_F31, Value
End Sub

Public Function F32() As Currency
   F32 = MEM(MemPos_F32)
End Function
Public Sub LET_F32(ByVal Value As Currency)
   LetByte MemPos_F32, Value
End Sub

Public Function F33() As Currency
   F33 = MEM(MemPos_F33)
End Function
Public Sub LET_F33(ByVal Value As Currency)
   LetByte MemPos_F33, Value
End Sub

Public Function F34() As Currency
   F34 = MEM(MemPos_F34)
End Function
Public Sub LET_F34(ByVal Value As Currency)
   LetByte MemPos_F34, Value
End Sub

Public Function F35() As Currency
   F35 = MEM(MemPos_F35)
End Function
Public Sub LET_F35(ByVal Value As Currency)
   LetByte MemPos_F35, Value
End Sub

Public Function F36() As Currency
   F36 = MEM(MemPos_F36)
End Function
Public Sub LET_F36(ByVal Value As Currency)
   LetByte MemPos_F36, Value
End Sub

Public Function F37() As Currency
   F37 = MEM(MemPos_F37)
End Function
Public Sub LET_F37(ByVal Value As Currency)
   LetByte MemPos_F37, Value
End Sub

Public Function F38() As Currency
   F38 = MEM(MemPos_F38)
End Function
Public Sub LET_F38(ByVal Value As Currency)
   LetByte MemPos_F38, Value
End Sub

Public Function F39() As Currency
   F39 = MEM(MemPos_F39)
End Function
Public Sub LET_F39(ByVal Value As Currency)
   LetByte MemPos_F39, Value
End Sub

Public Function F40() As Currency
   F40 = MEM(MemPos_F40)
End Function
Public Sub LET_F40(ByVal Value As Currency)
   LetByte MemPos_F40, Value
End Sub

Public Function F41() As Currency
   F41 = MEM(MemPos_F41)
End Function
Public Sub LET_F41(ByVal Value As Currency)
   LetByte MemPos_F41, Value
End Sub

Public Function F42() As Currency
   F42 = MEM(MemPos_F42)
End Function
Public Sub LET_F42(ByVal Value As Currency)
   LetByte MemPos_F42, Value
End Sub

Public Function F43() As Currency
   F43 = MEM(MemPos_F43)
End Function
Public Sub LET_F43(ByVal Value As Currency)
   LetByte MemPos_F43, Value
End Sub

Public Function F44() As Currency
   F44 = MEM(MemPos_F44)
End Function
Public Sub LET_F44(ByVal Value As Currency)
   LetByte MemPos_F44, Value
End Sub

Public Function F45() As Currency
   F45 = MEM(MemPos_F45)
End Function
Public Sub LET_F45(ByVal Value As Currency)
   LetByte MemPos_F45, Value
End Sub

Public Function F46() As Currency
   F46 = MEM(MemPos_F46)
End Function
Public Sub LET_F46(ByVal Value As Currency)
   LetByte MemPos_F46, Value
End Sub

Public Function F47() As Currency
   F47 = MEM(MemPos_F47)
End Function
Public Sub LET_F47(ByVal Value As Currency)
   LetByte MemPos_F47, Value
End Sub

Public Function F48() As Currency
   F48 = MEM(MemPos_F48)
End Function
Public Sub LET_F48(ByVal Value As Currency)
   LetByte MemPos_F48, Value
End Sub

Public Function F49() As Currency
   F49 = MEM(MemPos_F49)
End Function
Public Sub LET_F49(ByVal Value As Currency)
   LetByte MemPos_F49, Value
End Sub

Public Function F50() As Currency
   F50 = MEM(MemPos_F50)
End Function
Public Sub LET_F50(ByVal Value As Currency)
   LetByte MemPos_F50, Value
End Sub

Public Function F51() As Currency
   F51 = MEM(MemPos_F51)
End Function
Public Sub LET_F51(ByVal Value As Currency)
   LetByte MemPos_F51, Value
End Sub

Public Function F52() As Currency
   F52 = MEM(MemPos_F52)
End Function
Public Sub LET_F52(ByVal Value As Currency)
   LetByte MemPos_F52, Value
End Sub

Public Function F53() As Currency
   F53 = MEM(MemPos_F53)
End Function
Public Sub LET_F53(ByVal Value As Currency)
   LetByte MemPos_F53, Value
End Sub

Public Function F54() As Currency
   F54 = MEM(MemPos_F54)
End Function
Public Sub LET_F54(ByVal Value As Currency)
   LetByte MemPos_F54, Value
End Sub

Public Function F55() As Currency
   F55 = MEM(MemPos_F55)
End Function
Public Sub LET_F55(ByVal Value As Currency)
   LetByte MemPos_F55, Value
End Sub

Public Function F56() As Currency
   F56 = MEM(MemPos_F56)
End Function
Public Sub LET_F56(ByVal Value As Currency)
   LetByte MemPos_F56, Value
End Sub

Public Function F57() As Currency
   F57 = MEM(MemPos_F57)
End Function
Public Sub LET_F57(ByVal Value As Currency)
   LetByte MemPos_F57, Value
End Sub

Public Function F58() As Currency
   F58 = MEM(MemPos_F58)
End Function
Public Sub LET_F58(ByVal Value As Currency)
   LetByte MemPos_F58, Value
End Sub

Public Function F59() As Currency
   F59 = MEM(MemPos_F59)
End Function
Public Sub LET_F59(ByVal Value As Currency)
   LetByte MemPos_F59, Value
End Sub

Public Function F60() As Currency
   F60 = MEM(MemPos_F60)
End Function
Public Sub LET_F60(ByVal Value As Currency)
   LetByte MemPos_F60, Value
End Sub

Public Function F61() As Currency
   F61 = MEM(MemPos_F61)
End Function
Public Sub LET_F61(ByVal Value As Currency)
   LetByte MemPos_F61, Value
End Sub

Public Function F62() As Currency
   F62 = MEM(MemPos_F62)
End Function
Public Sub LET_F62(ByVal Value As Currency)
   LetByte MemPos_F62, Value
End Sub

Public Function F63() As Currency
   F63 = MEM(MemPos_F63)
End Function
Public Sub LET_F63(ByVal Value As Currency)
   LetByte MemPos_F63, Value
End Sub

Public Function F64() As Currency
   F64 = MEM(MemPos_F64)
End Function
Public Sub LET_F64(ByVal Value As Currency)
   LetByte MemPos_F64, Value
End Sub

Public Function F65() As Currency
   F65 = MEM(MemPos_F65)
End Function
Public Sub LET_F65(ByVal Value As Currency)
   LetByte MemPos_F65, Value
End Sub

Public Function F66() As Currency
   F66 = MEM(MemPos_F66)
End Function
Public Sub LET_F66(ByVal Value As Currency)
   LetByte MemPos_F66, Value
End Sub

Public Function F67() As Currency
   F67 = MEM(MemPos_F67)
End Function
Public Sub LET_F67(ByVal Value As Currency)
   LetByte MemPos_F67, Value
End Sub

Public Function F68() As Currency
   F68 = MEM(MemPos_F68)
End Function
Public Sub LET_F68(ByVal Value As Currency)
   LetByte MemPos_F68, Value
End Sub

Public Function F69() As Currency
   F69 = MEM(MemPos_F69)
End Function
Public Sub LET_F69(ByVal Value As Currency)
   LetByte MemPos_F69, Value
End Sub

Public Function F70() As Currency
   F70 = MEM(MemPos_F70)
End Function
Public Sub LET_F70(ByVal Value As Currency)
   LetByte MemPos_F70, Value
End Sub

Public Function F71() As Currency
   F71 = MEM(MemPos_F71)
End Function
Public Sub LET_F71(ByVal Value As Currency)
   LetByte MemPos_F71, Value
End Sub

Public Function F72() As Currency
   F72 = MEM(MemPos_F72)
End Function
Public Sub LET_F72(ByVal Value As Currency)
   LetByte MemPos_F72, Value
End Sub

Public Function F73() As Currency
   F73 = MEM(MemPos_F73)
End Function
Public Sub LET_F73(ByVal Value As Currency)
   LetByte MemPos_F73, Value
End Sub

Public Function F74() As Currency
   F74 = MEM(MemPos_F74)
End Function
Public Sub LET_F74(ByVal Value As Currency)
   LetByte MemPos_F74, Value
End Sub

Public Function F75() As Currency
   F75 = MEM(MemPos_F75)
End Function
Public Sub LET_F75(ByVal Value As Currency)
   LetByte MemPos_F75, Value
End Sub

Public Function F76() As Currency
   F76 = MEM(MemPos_F76)
End Function
Public Sub LET_F76(ByVal Value As Currency)
   LetByte MemPos_F76, Value
End Sub

Public Function F77() As Currency
   F77 = MEM(MemPos_F77)
End Function
Public Sub LET_F77(ByVal Value As Currency)
   LetByte MemPos_F77, Value
End Sub

Public Function F78() As Currency
   F78 = MEM(MemPos_F78)
End Function
Public Sub LET_F78(ByVal Value As Currency)
   LetByte MemPos_F78, Value
End Sub

Public Function F79() As Currency
   F79 = MEM(MemPos_F79)
End Function
Public Sub LET_F79(ByVal Value As Currency)
   LetByte MemPos_F79, Value
End Sub

Public Function F80() As Currency
   F80 = MEM(MemPos_F80)
End Function
Public Sub LET_F80(ByVal Value As Currency)
   LetByte MemPos_F80, Value
End Sub

Public Function F81() As Currency
   F81 = MEM(MemPos_F81)
End Function
Public Sub LET_F81(ByVal Value As Currency)
   LetByte MemPos_F81, Value
End Sub

Public Function F82() As Currency
   F82 = MEM(MemPos_F82)
End Function
Public Sub LET_F82(ByVal Value As Currency)
   LetByte MemPos_F82, Value
End Sub

Public Function F83() As Currency
   F83 = MEM(MemPos_F83)
End Function
Public Sub LET_F83(ByVal Value As Currency)
   LetByte MemPos_F83, Value
End Sub

Public Function F84() As Currency
   F84 = MEM(MemPos_F84)
End Function
Public Sub LET_F84(ByVal Value As Currency)
   LetByte MemPos_F84, Value
End Sub

Public Function F85() As Currency
   F85 = MEM(MemPos_F85)
End Function
Public Sub LET_F85(ByVal Value As Currency)
   LetByte MemPos_F85, Value
End Sub

Public Function F86() As Currency
   F86 = MEM(MemPos_F86)
End Function
Public Sub LET_F86(ByVal Value As Currency)
   LetByte MemPos_F86, Value
End Sub

Public Function F87() As Currency
   F87 = MEM(MemPos_F87)
End Function
Public Sub LET_F87(ByVal Value As Currency)
   LetByte MemPos_F87, Value
End Sub

Public Function F88() As Currency
   F88 = MEM(MemPos_F88)
End Function
Public Sub LET_F88(ByVal Value As Currency)
   LetByte MemPos_F88, Value
End Sub

Public Function F89() As Currency
   F89 = MEM(MemPos_F89)
End Function
Public Sub LET_F89(ByVal Value As Currency)
   LetByte MemPos_F89, Value
End Sub

Public Function F90() As Currency
   F90 = MEM(MemPos_F90)
End Function
Public Sub LET_F90(ByVal Value As Currency)
   LetByte MemPos_F90, Value
End Sub

Public Function F91() As Currency
   F91 = MEM(MemPos_F91)
End Function
Public Sub LET_F91(ByVal Value As Currency)
   LetByte MemPos_F91, Value
End Sub

Public Function F92() As Currency
   F92 = MEM(MemPos_F92)
End Function
Public Sub LET_F92(ByVal Value As Currency)
   LetByte MemPos_F92, Value
End Sub

Public Function F93() As Currency
   F93 = MEM(MemPos_F93)
End Function
Public Sub LET_F93(ByVal Value As Currency)
   LetByte MemPos_F93, Value
End Sub

Public Function F94() As Currency
   F94 = MEM(MemPos_F94)
End Function
Public Sub LET_F94(ByVal Value As Currency)
   LetByte MemPos_F94, Value
End Sub

Public Function F95() As Currency
   F95 = MEM(MemPos_F95)
End Function
Public Sub LET_F95(ByVal Value As Currency)
   LetByte MemPos_F95, Value
End Sub

Public Function F96() As Currency
   F96 = MEM(MemPos_F96)
End Function
Public Sub LET_F96(ByVal Value As Currency)
   LetByte MemPos_F96, Value
End Sub

Public Function F97() As Currency
   F97 = MEM(MemPos_F97)
End Function
Public Sub LET_F97(ByVal Value As Currency)
   LetByte MemPos_F97, Value
End Sub

Public Function F98() As Currency
   F98 = MEM(MemPos_F98)
End Function
Public Sub LET_F98(ByVal Value As Currency)
   LetByte MemPos_F98, Value
End Sub

Public Function F99() As Currency
   F99 = MEM(MemPos_F99)
End Function
Public Sub LET_F99(ByVal Value As Currency)
   LetByte MemPos_F99, Value
End Sub

' --- handle flag registers from F to F99 using offsets ---

Public Function F_OFS(ByVal Offset As Long) As Currency
   If Offset < MinFLow Or Offset > MaxFHigh Then
      ThrowError "F_OFS", "F Offset can only be " & Trim$(Str$(MinFLow)) & " to " & Trim$(Str$(MaxFHigh))
      Exit Function
   End If
   If Offset < MinFHigh Then
      F_OFS = MEM(MemPos_F + Offset)
   Else
      F_OFS = MEM(MemPos_F10 + Offset - MinFHigh)
   End If
End Function
Public Sub LET_F_OFS(ByVal Offset As Long, ByVal Value As Currency)
   If Offset < MinFLow Or Offset > MaxFHigh Then
      ThrowError "LET_F_OFS", "F Offset can only be " & Trim$(Str$(MinFLow)) & " to " & Trim$(Str$(MaxFHigh))
      Exit Sub
   End If
   If Offset < MinFHigh Then
      LetByte MemPos_F + Offset, Value
   Else
      LetByte MemPos_F10 + Offset - MinFHigh, Value
   End If
End Sub

' --- Global Flag registers ---

Public Function G() As Long
   G = GetGReg("G")
End Function
Public Sub LET_G(ByVal Value As Currency)
   ' --- This is a special version only used for the G register ---
   Dim strSQL As String
   ' ------------------
   If Value <> 255 And Value <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G", "G can only be set to 255 or the current USER number"
      Exit Sub
   End If
   strSQL = "SELECT * FROM [%GREGS] WHERE [NAME] = 'G'"
   With rsGRegs
      ' --- must be adUseServer for concurrency issues ---
      .CursorLocation = adUseServer
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      ' --- only do the update if G=255 (not in use) or G=USER (in use by me) ---
      If .Fields("VALUE") = 255 Or .Fields("VALUE") = GetNumeric(MemPos_User, 2) Then
         .Fields("VALUE") = Value
         .Update
      End If
      .Close
   End With
   Exit Sub
ConnError:
   ThrowError "LET_G", "SQL Connection Error:"
   Exit Sub
End Sub

Public Function G1() As Currency
   G1 = GetGReg("G1")
End Function
Public Sub LET_G1(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G1", "G1 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G1", Value
End Sub

Public Function G2() As Currency
   G2 = GetGReg("G2")
End Function
Public Sub LET_G2(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G2", "G2 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G2", Value
End Sub

Public Function G3() As Currency
   G3 = GetGReg("G3")
End Function
Public Sub LET_G3(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G3", "G3 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G3", Value
End Sub

Public Function G4() As Currency
   G4 = GetGReg("G4")
End Function
Public Sub LET_G4(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G4", "G4 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G4", Value
End Sub

Public Function G5() As Currency
   G5 = GetGReg("G5")
End Function
Public Sub LET_G5(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G5", "G5 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G5", Value
End Sub

Public Function G6() As Currency
   G6 = GetGReg("G6")
End Function
Public Sub LET_G6(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G6", "G6 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G6", Value
End Sub

Public Function G7() As Currency
   G7 = GetGReg("G7")
End Function
Public Sub LET_G7(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G7", "G7 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G7", Value
End Sub

Public Function G8() As Currency
   G8 = GetGReg("G8")
End Function
Public Sub LET_G8(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G8", "G8 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G8", Value
End Sub

Public Function G9() As Currency
   G9 = GetGReg("G9")
End Function
Public Sub LET_G9(ByVal Value As Currency)
   If GetGReg("G") <> GetNumeric(MemPos_User, 2) Then
      ThrowError "LET_G9", "G9 can only be changed when G equals the current USER number"
      Exit Sub
   End If
   LetGReg "G9", Value
End Sub

Public Function G_OFS(ByVal Offset As Long) As Currency
   Select Case Offset
      Case 0: G_OFS = G
      Case 1: G_OFS = G1
      Case 2: G_OFS = G2
      Case 3: G_OFS = G3
      Case 4: G_OFS = G4
      Case 5: G_OFS = G5
      Case 6: G_OFS = G6
      Case 7: G_OFS = G7
      Case 8: G_OFS = G8
      Case 9: G_OFS = G9
      Case Else
         If Offset < 0 Or Offset > 9 Then
            ThrowError "G_OFS", "G Offset can only be 0 to 9"
            Exit Function
         End If
   End Select
End Function
Public Sub LET_G_OFS(ByVal Offset As Long, ByVal Value As Currency)
   Select Case Offset
      Case 0: LET_G Value
      Case 1: LET_G1 Value
      Case 2: LET_G2 Value
      Case 3: LET_G3 Value
      Case 4: LET_G4 Value
      Case 5: LET_G5 Value
      Case 6: LET_G6 Value
      Case 7: LET_G7 Value
      Case 8: LET_G8 Value
      Case 9: LET_G9 Value
      Case Else
         If Offset < 0 Or Offset > 9 Then
            ThrowError "LET_G_OFS", "G Offset can only be 0 to 9"
            Exit Sub
         End If
   End Select
End Sub
