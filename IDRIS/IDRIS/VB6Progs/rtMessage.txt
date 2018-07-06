Attribute VB_Name = "rtMessage"
' ------------------------------
' --- rtMessage - 09/16/2008 ---
' ------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 09/16/2008 - SBAKKER - URD 11164
'            - Made changes recommended by CodeAdvisor.
' ------------------------------------------------------------------------------

Public Function PackMsg(ByVal Value As String) As String
   Dim Result As String
   Dim intLoop As Integer
   Dim intChar As Integer
   Dim lngCheckSum As Long
   ' -------------------------
   lngCheckSum = 0
   ' --- build packed string ---
   For intLoop = 1 To Len(Value)
      intChar = Asc(Mid$(Value, intLoop, 1))
      Result = Result & Right$("00" & Hex$(intChar), 2)
      lngCheckSum = lngCheckSum + intChar
      If lngCheckSum >= 32768 Then
         lngCheckSum = lngCheckSum Mod 32768
      End If
   Next
   ' --- add checksum ---
   Result = Result & Right$("0000" & Hex$(lngCheckSum), 4)
   ' --- return result ---
   PackMsg = Result
End Function

Public Function UnpackMsg(ByVal Value As String) As String
   Dim Result As String
   Dim intLoop As Integer
   Dim intChar As Integer
   Dim lngCheckSum As Long
   ' ---------------------
   lngCheckSum = 0
   ' --- build unpacked string ---
   Result = ""
   For intLoop = 1 To Len(Value) - 4 Step 2
      intChar = Val("&H" & Mid$(Value, intLoop, 2))
      If intChar < 0 Then
         intChar = intChar - (Int(intChar / 256) * 256)
      End If
      Result = Result & Chr$(intChar)
      lngCheckSum = lngCheckSum + intChar
      If lngCheckSum >= 32768 Then
         lngCheckSum = lngCheckSum Mod 32768
      End If
   Next
   ' --- validate against checksum ---
   If lngCheckSum <> Val("&H" & Right$(Value, 4)) Then
      Result = Value ' return original value if error
   End If
   ' --- return result ---
   UnpackMsg = Result
End Function
