Attribute VB_Name = "rtMemory"
' -----------------------------
' --- rtMemory - 10/01/2008 ---
' -----------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 10/01/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
'            - Made changes recommended by CodeAdvisor.
' 06/16/2008 - SBAKKER - URD 11118
'            - Added SERVER to Memory Dump for debugging purposes.
' 12/12/2006 - Removed unused local variables.
' 02/07/2006 - Load and save ProgILCode in Memory file at the end.
' 02/06/2006 - Load and save Breakpoint stack in Memory file. Added function
'              GetBreakpointsText to print out the Breakpoint list.
' 02/06/2006 - Clear GosubStack in LoadMemory, in case it already has entries.
' 02/06/2006 - Added GetMemPageLabels function to include column and row labels.
' 02/06/2006 - Moved To_Byte, AlphaLen, InBufferSpace, GetGosubStackText, and
'              MEMTF to rtMemory so they are available in IDRISClient.
' 02/03/2006 - Save and load the CurrIP value. This is needed for debugging, but
'              not for running programs. The IP value should match the values in
'              CurrDevNum, CurrVolName, CurrLibName, PROG, and CurrJumpPoint,
'              but this hasn't been verified. The "/START" parameter is still
'              used to set these values.
' 01/31/2006 - Fixed LoadMemory to check for valid Hex lines, not for "=" char.
' 01/30/2006 - Save current values of G registers into memory file (for
'              debugging purposes only).
' 01/20/2006 - Save and load N64-N99 variables into/out of memory pages.
' ------------------------------------------------------------------------------

Public Function SizeMax(ByVal Size As Long) As Currency
   Select Case Size
      Case 0: SizeMax = 1@
      Case 1: SizeMax = 256@
      Case 2: SizeMax = 65536@
      Case 3: SizeMax = 16777216@
      Case 4: SizeMax = 4294967296@
      Case 5: SizeMax = 1099511627776@
      Case 6: SizeMax = 281474976710656@
      Case Else
         ThrowError "SizeMax", "Invalid size: " & Trim$(Str$(Size))
   End Select
End Function

Public Function HalfSizeMax(ByVal Size As Long) As Currency
   Select Case Size
      Case 0: HalfSizeMax = 1@
      Case 1: HalfSizeMax = 128@
      Case 2: HalfSizeMax = 32768@
      Case 3: HalfSizeMax = 8388608@
      Case 4: HalfSizeMax = 2147483648@
      Case 5: HalfSizeMax = 549755813888@
      Case 6: HalfSizeMax = 140737488355328@
      Case Else
         ThrowError "HalfSizeMax", "Invalid size: " & Trim$(Str$(Size))
   End Select
End Function

Public Sub LetByte(ByVal Offset As Long, ByVal Value As Long)
   ' --- check if invalid value ---
   If Value < -256 Or Value > 256 Then
      ' --- check if Privg set for "ignore numeric overflow" ---
      If (MEM(MemPos_Privg) And 2) = 0 Then ' not set
         ThrowError "LetByte(" & Trim$(Str$(Offset)) & ")", "Numeric overflow: " & Trim$(Str$(Value))
         Exit Sub
      End If
   End If
   ' --- shift value into 0-255 range ---
   If Value < 0 Or Value > 255 Then
      Value = Value - (Int(Value / 256) * 256)
   End If
   ' --- save value to memory ---
   MEM(Offset) = Value
End Sub

Public Function GetAlpha(ByVal Offset As Long) As String
   Dim bTemp As Byte
   Dim bChar As Byte
   Dim lngLoop As Long
   Dim strResult As String
   ' ---------------------
   ' --- build string one char at a time ---
   strResult = ""
   lngLoop = 0
   Do
      bTemp = MEM(Offset + lngLoop)
      bChar = bTemp Mod 128
      If bChar <> 0 Then ' ignore nulls
         strResult = strResult & Chr$(bChar)
      End If
      lngLoop = lngLoop + 1
   Loop Until bTemp < 128
   ' --- get memory value ---
   GetAlpha = strResult
End Function

Public Sub LetAlpha(ByVal Offset As Long, ByVal Value As String)
   ' --- Note: FREEZE_LENGTH is used to prevent LENGTH from being updated here.     ---
   ' --- Some commands (ie ENTERALPHA) need to return 0 for a null string's length. ---
   ' --- If FREEZE_LENGTH wasn't used, a null string would set LENGTH to 1 here.    ---
   ' --- FREEZE_LENGTH is only good for one alpha assignment (LET_A ALPHA_RESULT).  ---
   Dim bChar As Byte
   Dim lngLen As Long
   Dim lngLoop As Long
   ' --------------------
   ' --- check if string too long ---
   If Len(Value) > 256 Then
      ThrowError "LetAlpha(" & Trim$(Str$(Offset)) & ")", "String longer then 256 chars"
      Exit Sub
   End If
   ' --- store string ---
   If Value = "" Then
      MEM(Offset) = 0 ' null string
      If Not FREEZE_LENGTH Then
         MEM(MemPos_Length) = 1 ' null = one byte
      End If
   Else
      lngLen = Len(Value)
      For lngLoop = 0 To lngLen - 1
         bChar = Asc(Mid$(Value, lngLoop + 1, 1))
         ' --- store alpha character ---
         If lngLoop < lngLen - 1 Then
            MEM(Offset + lngLoop) = bChar + 128
         Else
            MEM(Offset + lngLoop) = bChar
         End If
      Next lngLoop
      If Not FREEZE_LENGTH Then
         MEM(MemPos_Length) = To_Byte(lngLen)
      End If
   End If
   FREEZE_LENGTH = False
End Sub

Public Sub SpoolAlpha(ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   Dim lngLen As Long
   Dim lngLoop As Long
   ' -----------------
   If Size < 0 Or Size > 256 Then
      ThrowError "SpoolAlpha", "Invalid size: " & Trim$(Str$(Size))
      Exit Sub
   End If
   ' --- store string ---
   lngLen = Len(Value)
   For lngLoop = 0 To Size - 1
      If lngLoop >= lngLen Then
         MEM(Offset + lngLoop) = 32 + 128 ' space
      Else
         MEM(Offset + lngLoop) = Asc(Mid$(Value, lngLoop + 1, 1)) + 128
      End If
   Next lngLoop
End Sub

Public Function GetNumeric(ByVal Offset As Long, ByVal Size As Long) As Currency
   Dim lngLoop As Long
   Dim curResult As Currency
   ' -----------------------
   ' --- check if invalid size ---
   If Size < 1 Or Size > 6 Then
      ThrowError "GetNumeric(" & Trim$(Str$(Offset)) & ")", "Invalid Size: " & Trim$(Str$(Size))
      Exit Function
   End If
   ' --- get numeric value ---
   curResult = MEM(Offset)
   For lngLoop = 1 To Size - 1
      curResult = (curResult * 256) + MEM(Offset + lngLoop)
   Next lngLoop
   ' --- switch to negative if needed ---
   If curResult >= HalfSizeMax(Size) Then
      curResult = curResult - SizeMax(Size)
   End If
   ' --- return value ---
   GetNumeric = curResult
End Function

Public Sub LetNumeric(ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   Dim bTemp As Byte
   Dim lngLoop As Long
   Dim curValue As Currency
   ' ----------------------
   ' --- check if invalid size ---
   If Size < 1 Or Size > 6 Then
      ThrowError "LetNumeric(" & Trim$(Str$(Offset)) & ")", "Invalid Size: " & Trim$(Str$(Size))
      Exit Sub
   End If
   ' --- save value for adjusting ---
   curValue = Value
   ' --- check if value is too big for specified size ---
   If curValue >= HalfSizeMax(Size) Or curValue < -HalfSizeMax(Size) Then
      ' --- check if Privg set for "ignore numeric overflow" ---
      If (MEM(MemPos_Privg) And 2) = 0 Then ' not set
         ThrowError "LetNumeric(" & Trim$(Str$(Offset)) & ")", "Numeric overflow: " & Trim$(Str$(Value))
         Exit Sub
      End If
      ' --- adjust value into range 0 to SizeMax(Size)-1 ---
      curValue = curValue - (Int(curValue / SizeMax(Size)) * SizeMax(Size))
   ElseIf curValue < 0 Then
      ' --- adjust value into range 0 to SizeMax(Size)-1 ---
      curValue = curValue + SizeMax(Size)
   End If
   ' --- store value backwards, from end to beginning ---
   For lngLoop = 0 To Size - 1
      bTemp = Int(curValue / SizeMax(Size - lngLoop - 1))
      MEM(Offset + lngLoop) = bTemp
      curValue = curValue - (bTemp * SizeMax(Size - lngLoop - 1))
   Next lngLoop
End Sub

Public Sub MoveMem(ByVal FromOfs As Long, ByVal ToOfs As Long, ByVal Size As Long)
   Dim lngLoop As Long
   ' -----------------
   For lngLoop = 0 To Size - 1
      MEM(ToOfs + lngLoop) = MEM(FromOfs + lngLoop)
   Next lngLoop
End Sub

' ------------------------------
' --- Buffer Access Routines ---
' ------------------------------

Public Function GetNumericBuffer(ByVal BuffPtr As Long, ByVal Offset As Long, ByVal Size As Long) As Currency
   Dim lngPos As Long
   Dim curResult As Currency
   ' -----------------------
   lngPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
   If Not InBufferSpace(lngPos, Size) Then
      ThrowError "GetNumericBuffer", "*** Buffer Overflow ***"
      Exit Function
   End If
   ' --- get value ---
   curResult = GetNumeric(lngPos, Size)
   ' --- update buffer pointers ---
   lngPos = lngPos + Size
   MEM(BuffPtr + 1) = lngPos \ 256
   MEM(BuffPtr) = lngPos Mod 256
   ' --- return result ---
   GetNumericBuffer = curResult
End Function

Public Sub LetNumericBuffer(ByVal BuffPtr As Long, ByVal Offset As Long, ByVal Size As Long, ByVal Value As Currency)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
   If Not InBufferSpace(lngPos, Size) Then
      ThrowError "LetNumericBuffer", "*** Buffer Overflow ***"
      Exit Sub
   End If
   ' --- store value ---
   LetNumeric lngPos, Size, Value
   ' --- update buffer pointers ---
   lngPos = lngPos + Size
   MEM(BuffPtr + 1) = lngPos \ 256
   MEM(BuffPtr) = lngPos Mod 256
End Sub

Public Function GetAlphaBuffer(ByVal BuffPtr As Long, ByVal Offset As Long) As String
   Dim lngPos As Long
   Dim strResult As String
   ' ---------------------
   lngPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
   If Not InBufferSpace(lngPos, 1) Then
      ThrowError "GetAlphaBuffer", "*** Buffer Overflow ***"
      Exit Function
   End If
   ' --- store value ---
   strResult = GetAlpha(lngPos)
   If Not InBufferSpace(lngPos, AlphaLen(strResult)) Then
      ThrowError "GetAlphaBuffer", "*** Buffer Overflow ***"
      Exit Function
   End If
   ' --- update buffer pointers ---
   lngPos = lngPos + AlphaLen(strResult)
   MEM(BuffPtr + 1) = lngPos \ 256
   MEM(BuffPtr) = lngPos Mod 256
   ' --- return result ---
   GetAlphaBuffer = strResult
End Function

Public Sub LetAlphaBuffer(ByVal BuffPtr As Long, ByVal Offset As Long, ByVal Value As String)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
   If Not InBufferSpace(lngPos, AlphaLen(Value)) Then
      ThrowError "LetAlphaBuffer", "*** Buffer Overflow ***"
      Exit Sub
   End If
   ' --- store value ---
   LetAlpha lngPos, Value
   ' --- update buffer pointers ---
   lngPos = lngPos + AlphaLen(Value)
   MEM(BuffPtr + 1) = lngPos \ 256
   MEM(BuffPtr) = lngPos Mod 256
End Sub

Public Sub SpoolAlphaBuffer(ByVal BuffPtr As Long, ByVal Offset As Long, ByVal Size As Long, ByVal Value As String)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
   If Not InBufferSpace(lngPos, Size) Then
      ThrowError "SpoolAlphaBuffer", "*** Buffer Overflow ***"
      Exit Sub
   End If
   ' --- store value ---
   SpoolAlpha lngPos, Size, Value
   ' --- update buffer pointers ---
   lngPos = lngPos + Size
   MEM(BuffPtr + 1) = lngPos \ 256
   MEM(BuffPtr) = lngPos Mod 256
End Sub

Public Sub SaveMemory(ByVal FileName As String)
   Dim bTemp As Byte
   Dim lngLoop As Long
   Dim lngMemPos As Long
   Dim lngSaveFile As Long
   Dim strLine As String
   Dim strAscii As String
   Dim objItem As rtStackEntry
   ' -------------------------
   ' --- store numeric variables into memory ---
   For lngLoop = MinNLow To MaxNLow
      LetNumeric MemPos_N + (lngLoop * NumSlotSize), NumSlotSize, N_OFS(lngLoop)
   Next lngLoop
   For lngLoop = MinNHigh To MaxNHigh
      LetNumeric MemPos_N64 + ((lngLoop - MinNHigh) * NumSlotSize), NumSlotSize, N_OFS(lngLoop)
   Next lngLoop
   LetNumeric MemPos_Rec, NumSlotSize, REC
   LetNumeric MemPos_RemVal, NumSlotSize, REMVAL
   ' --- store current values of G registers (for debugging only) ---
   ' --- for now, return default values, not actual values ---
   LetNumeric MemPos_G, 2, 255 ' GetGReg("G")
   LetByte MemPos_G1, 0 ' GetGReg("G1")
   LetByte MemPos_G2, 0 ' GetGReg("G2")
   LetByte MemPos_G3, 0 ' GetGReg("G3")
   LetByte MemPos_G4, 0 ' GetGReg("G4")
   LetByte MemPos_G5, 0 ' GetGReg("G5")
   LetByte MemPos_G6, 0 ' GetGReg("G6")
   LetByte MemPos_G7, 0 ' GetGReg("G7")
   LetByte MemPos_G8, 0 ' GetGReg("G8")
   LetByte MemPos_G9, 0 ' GetGReg("G9")
   ' --- output memory block ---
   lngSaveFile = FreeFile
   Open FileName For Output As #lngSaveFile
   Close #lngSaveFile
   DoEvents
   lngSaveFile = FreeFile
   Open FileName For Output As #lngSaveFile
   strLine = ""
   strAscii = ""
   For lngMemPos = 0 To TotalMemSize - 1
      bTemp = MEM(lngMemPos)
      ' --- output in hex ---
      strLine = strLine & HexChar(bTemp)
      ' --- output in ascii for debugging ---
      bTemp = bTemp Mod 128
      If bTemp < 32 Or bTemp > 126 Then bTemp = Asc(".")
      strAscii = strAscii & Chr$(bTemp)
      ' --- print out every 16 bytes ---
      If lngMemPos Mod 16 = 15 Then
         Print #lngSaveFile, strLine & "   " & strAscii
         strLine = ""
         strAscii = ""
      End If
      ' --- add a blank line every 256 bytes (16 lines) ---
      If lngMemPos Mod 256 = 255 Then
         Print #lngSaveFile, ' add a blank line between pages
      End If
   Next lngMemPos
   ' --- save current instruction pointer ---
   strLine = "CURRIP=" & _
             Trim$(Str$(CurrDevNum)) & ":" & _
             CurrVolName & ":" & _
             CurrLibName & ":" & _
             Trim$(Str$(PROG)) & ":" & _
             Trim$(Str$(CurrJumpPoint))
   Print #lngSaveFile, strLine
   ' --- save gosub stack table ---
   For lngLoop = 1 To GosubStack.Count
      Set objItem = GosubStack.Item(lngLoop)
      Print #lngSaveFile, "GOSUBSTACK_" & Format$(lngLoop, "000") & "=" & objItem.ToString
   Next lngLoop
   ' --- save breakpoint table ---
   For lngLoop = 1 To Breakpoints.Count
      Set objItem = Breakpoints.Item(lngLoop)
      Print #lngSaveFile, "BREAKPOINT_" & Format$(lngLoop, "000") & "=" & objItem.ToString
   Next lngLoop
   ' --- save channel path table ---
   For lngLoop = 0 To MaxChannel
      If ChannelPaths(lngLoop) <> "" Then
         Print #lngSaveFile, "CHANPATH_" & Format$(lngLoop, "000") & "=" & ChannelPaths(lngLoop)
      End If
   Next lngLoop
   ' --- save keyboard buffer ---
   If KBuff.Count > 0 Then
      Print #lngSaveFile, "KBDBUFFER=" & KBuff.ToString
   End If
   ' --- save various string values ---
   If LoginID <> "" Then
      Print #lngSaveFile, "LOGINID=" & LCase$(LoginID)
   End If
   If MachineName <> "" Then
      Print #lngSaveFile, "MACHINENAME=" & LCase$(MachineName)
   End If
   If Environ("computername") <> "" Then
      Print #lngSaveFile, "SERVER=" & UCase$(Environ("computername"))
   End If
   If PrinterFileName <> "" Then
      Print #lngSaveFile, "PRINTERFILENAME=" & PrinterFileName
   End If
   If PrinterType <> "" Then
      Print #lngSaveFile, "PRINTERTYPE=" & PrinterType
   End If
   If PrinterDeviceName <> "" Then
      Print #lngSaveFile, "PRINTERDEVICENAME=" & PrinterDeviceName
   End If
   If PrinterParameters <> "" Then
      Print #lngSaveFile, "PRINTERPARAMETERS=" & PrinterParameters
   End If
   If SortFileName <> "" Then
      Print #lngSaveFile, "SORTFILENAME=" & SortFileName
   End If
   If SQLSubQuery <> "" Then
      Print #lngSaveFile, "SQLSUBQUERY=" & SQLSubQuery
   End If
   If SQLSubQueryFile <> "" Then
      Print #lngSaveFile, "SQLSUBQUERYFILE=" & SQLSubQueryFile
   End If
   If ClientList <> "" Then
      Print #lngSaveFile, "CLIENTLIST=" & ClientList
   End If
   If ReadOnly Then
      Print #lngSaveFile, "READONLY=TRUE"
   End If
   ' --- save various numeric values ---
   Print #lngSaveFile, "SORTTAGSIZE=" & Trim$(Str$(SortTagSize))
   Print #lngSaveFile, "SORTLINECOUNT=" & Trim$(Str$(SortLineCount))
   Print #lngSaveFile, "FETCHLINECOUNT=" & Trim$(Str$(FetchLineCount))
   ' --- save memory sort data ---
   If MEM(MemPos_SortState) <> 0 And SortFileName = "" Then
      For lngLoop = 1 To SortLineCount
         Print #lngSaveFile, "SORTTAGS_" & Format$(lngLoop, "000") & "=" & SortTags(lngLoop)
         Print #lngSaveFile, "SORTINDEX_" & Format$(lngLoop, "000") & "=" & Trim$(Str$(SortIndex(lngLoop)))
      Next lngLoop
   End If
   ' --- save IL code (must be last in file) ---
   If ProgILCode <> "" Then
      Print #lngSaveFile,
      Print #lngSaveFile, "ILCODE"
      Print #lngSaveFile, ProgILCode;
   End If
   ' --- done ---
   Close #lngSaveFile
End Sub

Public Function LoadMemory(ByVal FileName As String) As Boolean
   Dim bTemp As Byte
   Dim lngPos As Long
   Dim lngLoop As Long
   Dim lngValue As Long
   Dim lngMemPos As Long
   Dim strLine As String
   Dim strItem As String
   Dim strValue As String
   Dim lngSaveFile As Long
   Dim objItem As rtStackEntry
   Dim blnSkipRead As Boolean
   ' -------------------------
   If FileName = "" Then GoTo ErrorFound
   If Dir$(FileName) = "" Then
      ThrowError "LoadMemory", "File not found: " & FileName
      GoTo ErrorFound
   End If
   ' --- open memory file ---
   On Error GoTo ErrorFound
   lngSaveFile = FreeFile
   Open FileName For Input As #lngSaveFile
   ' --- load all old memory bytes ---
   blnSkipRead = False
   For lngMemPos = 0 To TotalMemSize - 1
      If lngMemPos Mod 16 = 0 Then
         Do
            Line Input #lngSaveFile, strLine
         Loop Until strLine <> ""
      End If
      ' --- check if reading an older memory file with a newer program ---
      If Not IsHexLine(strLine) Then
         blnSkipRead = True
         Exit For
      End If
      bTemp = CByte("&H" & Mid$(strLine, ((lngMemPos Mod 16) * 2) + 1, 2))
      MEM(lngMemPos) = bTemp
   Next lngMemPos
   ' --- get numeric variables from memory ---
   For lngLoop = MinNLow To MaxNLow
      LET_N_OFS lngLoop, GetNumeric(MemPos_N + (lngLoop * NumSlotSize), NumSlotSize)
   Next lngLoop
   For lngLoop = MinNHigh To MaxNHigh
      LET_N_OFS lngLoop, GetNumeric(MemPos_N64 + ((lngLoop - MinNHigh) * NumSlotSize), NumSlotSize)
   Next lngLoop
   REC = GetNumeric(MemPos_Rec, NumSlotSize)
   REMVAL = GetNumeric(MemPos_RemVal, NumSlotSize)
   ' --- clear stacks and variables ---
   Do While GosubStack.Count > 0
      GosubStack.Remove 1
   Loop
   Do While Breakpoints.Count > 0
      Breakpoints.Remove 1
   Loop
   ProgILCode = ""
   ' --- check for any names/paths/stack entries ---
   Do While Not EOF(lngSaveFile)
      If Not blnSkipRead Then
         Line Input #lngSaveFile, strLine
      End If
      blnSkipRead = False
      lngPos = InStr(strLine, "=") ' position of first equal sign
      If lngPos = 0 Then
         If UCase$(strLine) = "ILCODE" Then
            ' --- everything else in the file is IL code ---
            GoTo GetILCode
         End If
         GoTo NextItem
      End If
      strItem = UCase$(Left$(strLine, lngPos - 1)) ' line's item
      strValue = Mid$(strLine, lngPos + 1) ' line's value
      ' --- load current instruction pointer (but don't use it yet) ---
      If strItem = "CURRIP" Then
         CurrIP = strValue
         GoTo NextItem
      End If
      ' --- gosub stack entry ---
      If Left$(strItem, 11) = "GOSUBSTACK_" Then
         Set objItem = New rtStackEntry
         objItem.FromString strValue
         GosubStack.Add objItem
         Set objItem = Nothing
         GoTo NextItem
      End If
      ' --- breakpoint entry ---
      If Left$(strItem, 11) = "BREAKPOINT_" Then
         Set objItem = New rtStackEntry
         objItem.FromString strValue
         Breakpoints.Add objItem
         Set objItem = Nothing
         GoTo NextItem
      End If
      ' --- channel path ---
      If Left$(strItem, 9) = "CHANPATH_" Then
         lngValue = Val(Mid$(strItem, 10, 3))
         ChannelPaths(lngValue) = strValue
         GoTo NextItem
      End If
      ' --- keyboard buffer ---
      If strItem = "KBDBUFFER" Then
         KBuff.FromString strValue
         GoTo NextItem
      End If
      ' --- load various string values ---
      If strItem = "LOGINID" Then
         LoginID = LCase$(strValue)
         GoTo NextItem
      End If
      If strItem = "MACHINENAME" Then
         MachineName = LCase$(strValue)
         GoTo NextItem
      End If
      If strItem = "PRINTERFILENAME" Then
         PrinterFileName = strValue
         GoTo NextItem
      End If
      If strItem = "PRINTERTYPE" Then
         PrinterType = strValue
         GoTo NextItem
      End If
      If strItem = "PRINTERDEVICENAME" Then
         PrinterDeviceName = strValue
         GoTo NextItem
      End If
      If strItem = "PRINTERPARAMETERS" Then
         PrinterParameters = strValue
         GoTo NextItem
      End If
      If strItem = "SORTFILENAME" Then
         SortFileName = strValue
         GoTo NextItem
      End If
      If strItem = "SQLSUBQUERY" Then
         SQLSubQuery = strValue
         GoTo NextItem
      End If
      If strItem = "SQLSUBQUERYFILE" Then
         SQLSubQueryFile = strValue
         GoTo NextItem
      End If
      If strItem = "CLIENTLIST" Then
         If ClientList = "" Then
            ClientList = strValue
            If ClientList <> "" Then
               DebugMessage "ClientList = """ & ClientList & """"
               PrepareClientLists
            End If
         End If
         GoTo NextItem
      End If
      If strItem = "READONLY" Then
         If Not ReadOnly Then
            ReadOnly = (UCase$(strValue) = "TRUE")
            DebugMessage "ReadOnly = """ & UCase$(strValue) & """"
         End If
      End If
      ' --- load various numeric values ---
      If strItem = "SORTTAGSIZE" Then
         SortTagSize = Val(strValue)
         GoTo NextItem
      End If
      If strItem = "SORTLINECOUNT" Then
         SortLineCount = Val(strValue)
         GoTo NextItem
      End If
      If strItem = "FETCHLINECOUNT" Then
         FetchLineCount = Val(strValue)
         GoTo NextItem
      End If
      ' --- load memory sort data ---
      If Left$(strItem, 9) = "SORTTAGS_" Then
         lngValue = Val(Mid$(strItem, 10, 3))
         SortTags(lngValue) = strValue
         GoTo NextItem
      End If
      If Left$(strItem, 10) = "SORTINDEX_" Then
         lngValue = Val(Mid$(strItem, 11, 3))
         SortIndex(lngValue) = Val(strValue)
         GoTo NextItem
      End If
      ' --- get next item ---
NextItem:
   Loop
   GoTo Done
   ' --- get IL code ---
GetILCode:
   ProgILCode = ""
   Do While Not EOF(lngSaveFile)
      Line Input #lngSaveFile, strLine
      ProgILCode = ProgILCode & strLine & vbCrLf
   Loop
   GoTo Done
   ' --- done ---
Done:
   Close #lngSaveFile
   DoEvents
   ' --- delete file after loaded ---
   If Not InsideIDE And Not DebugFlag Then
      On Error Resume Next
      Kill FileName
      On Error GoTo 0
   End If
   ' --- done ---
   LoadMemory = True
   Exit Function
ErrorFound:
   ' --- error ---
   On Error Resume Next
   Close #lngSaveFile
   DoEvents
   On Error GoTo 0
   LoadMemory = False
End Function

Public Function GetMemPage(ByVal Offset As Long) As String
   Dim LoopX As Long
   Dim LoopY As Long
   Dim bChar As Byte
   Dim strResult As String
   ' ---------------------
   ' --- change page number to byte number ---
   If Offset < 256 Then Offset = Offset * 256
   ' --- build page map ---
   strResult = ""
   For LoopY = 0 To 15
      For LoopX = 0 To 15
         bChar = MEM(Offset + (LoopY * 16) + LoopX)
         strResult = strResult & HexChar(bChar)
         strResult = strResult & " "
      Next LoopX
      strResult = strResult & "  "
      For LoopX = 0 To 15
         bChar = MEM(Offset + (LoopY * 16) + LoopX)
         bChar = ModPos(bChar, 128)
         If bChar < 32 Or bChar > 126 Then bChar = 46 ' period
         strResult = strResult & Chr$(bChar)
      Next LoopX
      strResult = strResult & vbCrLf
   Next LoopY
   GetMemPage = strResult
End Function

Public Function GetMemPageLabels(ByVal Offset As Long) As String
   Dim LoopX As Long
   Dim LoopY As Long
   Dim bChar As Byte
   Dim strResult As String
   ' ---------------------
   ' --- change page number to byte number ---
   If Offset < 256 Then Offset = Offset * 256
   ' --- build page map ---
   strResult = "    "
   For LoopX = 0 To 15
      strResult = strResult & " " & HexChar(LoopX)
   Next LoopX
   strResult = strResult & vbCrLf & "     " & String$(47, "-") & vbCrLf
   For LoopY = 0 To 15
      strResult = strResult & HexChar(LoopY * 16) & " | "
      For LoopX = 0 To 15
         bChar = MEM(Offset + (LoopY * 16) + LoopX)
         strResult = strResult & HexChar(bChar)
         strResult = strResult & " "
      Next LoopX
      strResult = strResult & "  "
      For LoopX = 0 To 15
         bChar = MEM(Offset + (LoopY * 16) + LoopX)
         bChar = ModPos(bChar, 128)
         If bChar < 32 Or bChar > 126 Then bChar = 46 ' period
         strResult = strResult & Chr$(bChar)
      Next LoopX
      strResult = strResult & vbCrLf
   Next LoopY
   GetMemPageLabels = strResult
End Function

Public Function To_Byte(ByVal Value As Currency) As Byte
   If Value >= 0 And Value < 256 Then
      To_Byte = Value
   Else
      To_Byte = Value - (Int(Value / 256) * 256)
   End If
End Function

Public Function AlphaLen(ByVal Value As String) As Long
   If Value = "" Then
      AlphaLen = 1 ' one byte
   Else
      AlphaLen = Len(Value)
   End If
End Function

Public Function InBufferSpace(ByVal Position As Long, ByVal Size As Long) As Boolean
   ' --- check if completely outside memory ---
   If Position < 0 Then GoTo ErrorFound
   If Not MEMTF(MemPos_TBAlloc) Then
      If Position + Size > MemPos_TrackBuffer Then GoTo ErrorFound
   Else
      If Position + Size > MemPos_TrackBuffer + (32 * 256) Then GoTo ErrorFound
   End If
   ' --- if #2 bit of PRIVG is not set, must be within a buffer ---
   If MEM(MemPos_Privg) Mod 2 = 0 Then
      If Position < MemPos_R Then GoTo ErrorFound
      If Position + Size > MemPos_W + 256 And Position < MemPos_S Then GoTo ErrorFound
      If Position + Size > MemPos_V + 256 Then GoTo ErrorFound
   End If
   ' --- are within buffer space ---
   InBufferSpace = True
   Exit Function
ErrorFound:
   InBufferSpace = False
End Function

Public Sub LET_MEMTF(ByVal Offset As Long, ByVal Value As Boolean)
   If Value Then
      MEM(Offset) = TRUEVAL
   Else
      MEM(Offset) = FALSEVAL
   End If
End Sub

Public Function MEMTF(ByVal Offset As Long) As Boolean
   MEMTF = (MEM(Offset) <> FALSEVAL)
End Function

Public Function IsHexLine(ByVal Value As String) As Boolean
   Dim lngLoop As Long
   Dim strChar As String
   ' -------------------
   IsHexLine = False
   If Len(Value) <> 32 And Len(Value) <> 51 Then Exit Function
   For lngLoop = 0 To 31
      strChar = Mid$(Value, lngLoop + 1, 1)
      If strChar < "0" Or strChar > "9" Then
         If strChar < "A" Or strChar > "F" Then
            Exit Function
         End If
      End If
   Next
   IsHexLine = True
End Function

Public Function GetHexAlpha(ByVal MemPos As Long, ByVal HexLen As Long) As String
   Dim lngLoop As Integer
   Dim strResult As String
   ' ---------------------
   strResult = ""
   For lngLoop = MemPos To MemPos + HexLen - 1
      strResult = strResult & HexChar(MEM(lngLoop))
   Next
   GetHexAlpha = strResult
End Function

Public Function GetGosubStackText() As String
   Dim lngLoop As Long
   Dim strTemp As String
   Dim strResult As String
   Dim objItem As rtStackEntry
   ' -------------------------
   strResult = ""
   lngLoop = 1
   Do While lngLoop <= GosubStack.Count
      Set objItem = GosubStack.Item(lngLoop)
      strTemp = objItem.ToString
      If strResult <> "" And strTemp <> "" Then strResult = strResult & vbCrLf
      strResult = strResult & strTemp
      lngLoop = lngLoop + 1
   Loop
   GetGosubStackText = strResult
End Function

Public Function GetBreakpointsText() As String
   Dim lngLoop As Long
   Dim strTemp As String
   Dim strResult As String
   Dim objItem As rtStackEntry
   ' -------------------------
   strResult = ""
   lngLoop = 1
   Do While lngLoop <= Breakpoints.Count
      Set objItem = Breakpoints.Item(lngLoop)
      strTemp = objItem.ToString
      If strResult <> "" And strTemp <> "" Then strResult = strResult & vbCrLf
      strResult = strResult & strTemp
      lngLoop = lngLoop + 1
   Loop
   GetBreakpointsText = strResult
End Function
