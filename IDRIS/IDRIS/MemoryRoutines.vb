' --------------------------------------
' --- MemoryRoutines.vb - 10/26/2016 ---
' --------------------------------------

' ------------------------------------------------------------------------------
' 09/17/2013 - SBakker
'            - Added additional error information.
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

Imports System.Text

Module MemoryRoutines

    Public Function SizeMax(ByVal Size As Integer) As Int64
        Select Case Size
            Case 0 : Return 1L
            Case 1 : Return 256L
            Case 2 : Return 65536L
            Case 3 : Return 16777216L
            Case 4 : Return 4294967296L
            Case 5 : Return 1099511627776L
            Case 6 : Return 281474976710656L
        End Select
        Throw New SystemException($"SizeMax: Invalid size: {Size}")
    End Function

    Public Function HalfSizeMax(ByVal Size As Integer) As Int64
        Select Case Size
            Case 0 : Return 1L
            Case 1 : Return 128L
            Case 2 : Return 32768L
            Case 3 : Return 8388608L
            Case 4 : Return 2147483648L
            Case 5 : Return 549755813888L
            Case 6 : Return 140737488355328L
        End Select
        Throw New SystemException($"HalfSizeMax: Invalid size: {Size}")
    End Function

    Public Sub LetByte(ByVal Offset As Integer, ByVal Value As Int64)
        ' --- check if invalid value ---
        If Value < -256 Or Value > 256 Then
            ' --- check if Privg set for "ignore numeric overflow" ---
            If (MEM(MemPos_Privg) And 2) = 0 Then ' not set
                Throw New SystemException($"LetByte({Offset}): Numeric overflow: {Value}")
            End If
        End If
        ' --- shift value into 0-255 range ---
        If Value < 0 Or Value > 255 Then
            Value = Value - (CInt(Value / 256) * 256)
        End If
        ' --- save value to memory ---
        MEM(Offset) = CByte(Value)
    End Sub

    Public Function GetAlpha(ByVal Offset As Integer) As String
        ' --- Check for null string ---
        If MEM(Offset) = 0 Then Return ""
        ' --- build string one char at a time ---
        Dim Result As New StringBuilder
        Dim OrigChar As Integer
        Dim FixedChar As Integer
        Dim LoopNum As Integer = 0
        Do
            OrigChar = MEM(Offset + LoopNum)
            LoopNum += 1
            FixedChar = OrigChar Mod 128
            If FixedChar <> 0 Then
                Result.Append(Chr(FixedChar))
            End If
        Loop Until OrigChar < 128
        Return Result.ToString
    End Function

    Public Sub LetAlpha(ByVal Offset As Integer, ByVal Value As String)
        ' --- Note: FREEZE_LENGTH is used to prevent LENGTH from being updated here.     ---
        ' --- Some commands (ie ENTERALPHA) need to return 0 for a null string's length. ---
        ' --- If FREEZE_LENGTH wasn't used, a null string would set LENGTH to 1 here.    ---
        ' --- FREEZE_LENGTH is only good for one alpha assignment (LET_A ALPHA_RESULT).  ---
        Dim TempChar As Byte
        Dim TempLen As Integer
        Dim LoopNum As Integer
        ' --------------------
        ' --- check if string too long ---
        If Value.Length > 256 Then
            Throw New SystemException($"LetAlpha({Offset}): String longer then 256 chars")
        End If
        ' --- store string ---
        If Value = "" Then
            MEM(Offset) = 0 ' null string
            If Not FREEZE_LENGTH Then
                MEM(MemPos_Length) = 1 ' null = one byte
            End If
        Else
            TempLen = Value.Length
            For LoopNum = 0 To TempLen - 1
                TempChar = CByte(Asc(Value(LoopNum)))
                ' --- store alpha character ---
                If LoopNum < TempLen - 1 Then
                    MEM(Offset + LoopNum) = CByte(TempChar + 128)
                Else
                    MEM(Offset + LoopNum) = TempChar
                End If
            Next LoopNum
            If Not FREEZE_LENGTH Then
                LetByte(MemPos_Length, TempLen)
            End If
        End If
        FREEZE_LENGTH = False
    End Sub

    Public Sub SpoolAlpha(ByVal Offset As Integer, ByVal Size As Integer, ByVal Value As String)
        Dim TempLen As Integer
        Dim LoopNum As Integer
        ' --------------------
        If Size < 0 Or Size > 256 Then
            Throw New SystemException($"SpoolAlpha: Invalid size: {Size}")
        End If
        ' --- store string ---
        TempLen = Value.Length
        For LoopNum = 0 To Size - 1
            If LoopNum >= TempLen Then
                MEM(Offset + LoopNum) = 32 + 128 ' space
            Else
                MEM(Offset + LoopNum) = CByte(Asc(Value(LoopNum)) + 128)
            End If
        Next LoopNum
    End Sub

    Public Function GetNumeric(ByVal Offset As Integer, ByVal Size As Integer) As Int64
        Dim LoopNum As Integer
        Dim Result As Int64
        ' --------------------
        ' --- check if invalid size ---
        If Size < 1 Or Size > 6 Then
            Throw New SystemException($"GetNumeric({Offset}): Invalid Size: {Size}")
        End If
        ' --- get numeric value ---
        Result = MEM(Offset)
        For LoopNum = 1 To Size - 1
            Result = (Result * 256) + MEM(Offset + LoopNum)
        Next LoopNum
        ' --- switch to negative if needed ---
        If Result >= HalfSizeMax(Size) Then
            Result = Result - SizeMax(Size)
        End If
        ' --- return value ---
        Return Result
    End Function

    Public Sub LetNumeric(ByVal Offset As Integer, ByVal Size As Integer, ByVal Value As Int64)
        Dim TempByte As Byte
        Dim LoopNum As Integer
        Dim TempValue As Int64
        ' --------------------
        ' --- check if invalid size ---
        If Size < 1 Or Size > 6 Then
            Throw New SystemException($"LetNumeric({Offset}): Invalid Size: {Size}")
        End If
        ' --- save value for adjusting ---
        TempValue = Value
        ' --- check if value is too big for specified size ---
        If TempValue >= HalfSizeMax(Size) Or TempValue < -HalfSizeMax(Size) Then
            ' --- check if Privg set for "ignore numeric overflow" ---
            If (MEM(MemPos_Privg) And 2) = 0 Then ' not set
                Throw New SystemException($"LetNumeric({Offset}): Numeric overflow: {Value}")
            End If
            ' --- adjust value into range 0 to SizeMax(Size)-1 ---
            TempValue = TempValue - CLng(Int(TempValue / SizeMax(Size))) * SizeMax(Size)
        ElseIf TempValue < 0 Then
            ' --- adjust value into range 0 to SizeMax(Size)-1 ---
            TempValue = TempValue + SizeMax(Size)
        End If
        ' --- store value backwards, from end to beginning ---
        For LoopNum = 0 To Size - 1
            TempByte = CByte(Int(TempValue / SizeMax(Size - LoopNum - 1)))
            MEM(Offset + LoopNum) = TempByte
            TempValue = TempValue - (TempByte * SizeMax(Size - LoopNum - 1))
        Next LoopNum
    End Sub

    Public Sub MoveMem(ByVal FromOfs As Integer, ByVal ToOfs As Integer, ByVal Size As Integer)
        Dim lngLoop As Integer
        ' --------------------
        For lngLoop = 0 To Size - 1
            MEM(ToOfs + lngLoop) = MEM(FromOfs + lngLoop)
        Next lngLoop
    End Sub

    ' ------------------------------
    ' --- Buffer Access Routines ---
    ' ------------------------------

    Public Function GetNumericBuffer(ByVal BuffPtr As Integer, ByVal Offset As Integer, ByVal Size As Integer) As Int64
        Dim TempPos As Integer
        Dim Result As Int64
        ' --------------------
        TempPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
        If Not InBufferSpace(TempPos, Size) Then
            Throw New SystemException("GetNumericBuffer: *** Buffer Overflow ***")
        End If
        ' --- get value ---
        Result = GetNumeric(TempPos, Size)
        ' --- update buffer pointers ---
        TempPos = TempPos + Size
        MEM(BuffPtr + 1) = CByte(TempPos \ 256)
        MEM(BuffPtr) = CByte(TempPos Mod 256)
        ' --- return result ---
        Return Result
    End Function

    Public Sub LetNumericBuffer(ByVal BuffPtr As Integer, ByVal Offset As Integer, ByVal Size As Integer, ByVal Value As Int64)
        Dim TempPos As Integer
        ' --------------------
        TempPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
        If Not InBufferSpace(TempPos, Size) Then
            Throw New SystemException("LetNumericBuffer: *** Buffer Overflow ***")
        End If
        ' --- store value ---
        LetNumeric(TempPos, Size, Value)
        ' --- update buffer pointers ---
        TempPos = TempPos + Size
        MEM(BuffPtr + 1) = CByte(TempPos \ 256)
        MEM(BuffPtr) = CByte(TempPos Mod 256)
    End Sub

    Public Function GetAlphaBuffer(ByVal BuffPtr As Integer, ByVal Offset As Integer) As String
        Dim TempPos As Integer
        Dim Result As String
        ' --------------------
        TempPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
        If Not InBufferSpace(TempPos, 1) Then
            Throw New SystemException("GetAlphaBuffer: *** Buffer Overflow ***")
        End If
        ' --- store value ---
        Result = GetAlpha(TempPos)
        If Not InBufferSpace(TempPos, AlphaLen(Result)) Then
            Throw New SystemException("GetAlphaBuffer: *** Buffer Overflow ***")
        End If
        ' --- update buffer pointers ---
        TempPos = TempPos + AlphaLen(Result)
        MEM(BuffPtr + 1) = CByte(TempPos \ 256)
        MEM(BuffPtr) = CByte(TempPos Mod 256)
        ' --- return result ---
        Return Result
    End Function

    Public Sub LetAlphaBuffer(ByVal BuffPtr As Integer, ByVal Offset As Integer, ByVal Value As String)
        Dim TempPos As Integer
        ' --------------------
        TempPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
        If Not InBufferSpace(TempPos, AlphaLen(Value)) Then
            Throw New SystemException("LetAlphaBuffer: *** Buffer Overflow ***")
        End If
        ' --- store value ---
        LetAlpha(TempPos, Value)
        ' --- update buffer pointers ---
        TempPos = TempPos + AlphaLen(Value)
        MEM(BuffPtr + 1) = CByte(TempPos \ 256)
        MEM(BuffPtr) = CByte(TempPos Mod 256)
    End Sub

    Public Sub SpoolAlphaBuffer(ByVal BuffPtr As Integer, ByVal Offset As Integer, ByVal Size As Integer, ByVal Value As String)
        Dim TempPos As Integer
        ' --------------------
        TempPos = (MEM(BuffPtr + 1) * 256) + MEM(BuffPtr) + Offset
        If Not InBufferSpace(TempPos, Size) Then
            Throw New SystemException("SpoolAlphaBuffer: *** Buffer Overflow ***")
        End If
        ' --- store value ---
        SpoolAlpha(TempPos, Size, Value)
        ' --- update buffer pointers ---
        TempPos = TempPos + Size
        MEM(BuffPtr + 1) = CByte(TempPos \ 256)
        MEM(BuffPtr) = CByte(TempPos Mod 256)
    End Sub

    ' Public Sub SaveMemory(ByVal FileName As String)
    '     Dim bTemp As Byte
    '     Dim lngLoop As Integer
    '     Dim lngMemPos As Integer
    '     Dim lngSaveFile As Integer
    '     Dim strLine As String
    '     Dim strAscii As String
    '     Dim objItem As rtStackEntry
    '     ' -------------------------
    '     ' --- store numeric variables into memory ---
    '     For lngLoop = MinNLow To MaxNLow
    '         LetNumeric(MemPos_N + (lngLoop * NumSlotSize), NumSlotSize, N_OFS(lngLoop))
    '     Next lngLoop
    '     For lngLoop = MinNHigh To MaxNHigh
    '         LetNumeric(MemPos_N64 + ((lngLoop - MinNHigh) * NumSlotSize), NumSlotSize, N_OFS(lngLoop))
    '     Next lngLoop
    '     LetNumeric(MemPos_Rec, NumSlotSize, REC)
    '     LetNumeric(MemPos_RemVal, NumSlotSize, REMVAL)
    '     ' --- store current values of G registers (for debugging only) ---
    '     ' --- for now, return default values, not actual values ---
    '     LetNumeric(MemPos_G, 2, 255) ' GetGReg("G")
    '     LetByte(MemPos_G1, 0) ' GetGReg("G1")
    '     LetByte(MemPos_G2, 0) ' GetGReg("G2")
    '     LetByte(MemPos_G3, 0) ' GetGReg("G3")
    '     LetByte(MemPos_G4, 0) ' GetGReg("G4")
    '     LetByte(MemPos_G5, 0) ' GetGReg("G5")
    '     LetByte(MemPos_G6, 0) ' GetGReg("G6")
    '     LetByte(MemPos_G7, 0) ' GetGReg("G7")
    '     LetByte(MemPos_G8, 0) ' GetGReg("G8")
    '     LetByte(MemPos_G9, 0) ' GetGReg("G9")
    '     ' --- output memory block ---
    '     lngSaveFile = FreeFile()
    'Open FileName For Output As #lngSaveFile
    'Close #lngSaveFile
    '     DoEvents()
    '     lngSaveFile = FreeFile()
    'Open FileName For Output As #lngSaveFile
    '     strLine = ""
    '     strAscii = ""
    '     For lngMemPos = 0 To TotalMemSize - 1
    '         bTemp = MEM(lngMemPos)
    '         ' --- output in hex ---
    '         strLine = strLine & HexChar(bTemp)
    '         ' --- output in ascii for debugging ---
    '         bTemp = bTemp Mod 128
    '         If bTemp < 32 Or bTemp > 126 Then bTemp = Asc(".")
    '         strAscii = strAscii & Chr$(bTemp)
    '         ' --- print out every 16 bytes ---
    '         If lngMemPos Mod 16 = 15 Then
    '      Print #lngSaveFile, strLine & "   " & strAscii
    '             strLine = ""
    '             strAscii = ""
    '         End If
    '         ' --- add a blank line every 256 bytes (16 lines) ---
    '         If lngMemPos Mod 256 = 255 Then
    '      Print #lngSaveFile, ' add a blank line between pages
    '         End If
    '     Next lngMemPos
    '     ' --- save current instruction pointer ---
    '     strLine = "CURRIP=" & _
    '               Trim$(Str$(CurrDevNum.tostring +":" & _
    '               CurrVolName & ":" & _
    '               CurrLibName & ":" & _
    '               Trim$(Str$(PROG.tostring +":" & _
    '               Trim$(Str$(CurrJumpPoint))
    'Print #lngSaveFile, strLine
    '     ' --- save gosub stack table ---
    '     For lngLoop = 1 To GosubStack.Count
    '         objItem = GosubStack.Item(lngLoop)
    '   Print #lngSaveFile, "GOSUBSTACK_" & Format$(lngLoop, "000") & "=" & objItem.ToString
    '     Next lngLoop
    '     ' --- save breakpoint table ---
    '     For lngLoop = 1 To Breakpoints.Count
    '         objItem = Breakpoints.Item(lngLoop)
    '   Print #lngSaveFile, "BREAKPOINT_" & Format$(lngLoop, "000") & "=" & objItem.ToString
    '     Next lngLoop
    '     ' --- save channel path table ---
    '     For lngLoop = 0 To MaxChannel
    '         If ChannelPaths(lngLoop) <> "" Then
    '      Print #lngSaveFile, "CHANPATH_" & Format$(lngLoop, "000") & "=" & ChannelPaths(lngLoop)
    '         End If
    '     Next lngLoop
    '     ' --- save keyboard buffer ---
    '     If KBuff.Count > 0 Then
    '   Print #lngSaveFile, "KBDBUFFER=" & KBuff.ToString
    '     End If
    '     ' --- save various string values ---
    '     If LoginID <> "" Then
    '   Print #lngSaveFile, "LOGINID=" & LCase$(LoginID)
    '     End If
    '     If MachineName <> "" Then
    '   Print #lngSaveFile, "MACHINENAME=" & LCase$(MachineName)
    '     End If
    '     If Environ("computername") <> "" Then
    '   Print #lngSaveFile, "SERVER=" & UCase$(Environ("computername"))
    '     End If
    '     If PrinterFileName <> "" Then
    '   Print #lngSaveFile, "PRINTERFILENAME=" & PrinterFileName
    '     End If
    '     If PrinterType <> "" Then
    '   Print #lngSaveFile, "PRINTERTYPE=" & PrinterType
    '     End If
    '     If PrinterDeviceName <> "" Then
    '   Print #lngSaveFile, "PRINTERDEVICENAME=" & PrinterDeviceName
    '     End If
    '     If PrinterParameters <> "" Then
    '   Print #lngSaveFile, "PRINTERPARAMETERS=" & PrinterParameters
    '     End If
    '     If SortFileName <> "" Then
    '   Print #lngSaveFile, "SORTFILENAME=" & SortFileName
    '     End If
    '     If SQLSubQuery <> "" Then
    '   Print #lngSaveFile, "SQLSUBQUERY=" & SQLSubQuery
    '     End If
    '     If SQLSubQueryFile <> "" Then
    '   Print #lngSaveFile, "SQLSUBQUERYFILE=" & SQLSubQueryFile
    '     End If
    '     If ClientList <> "" Then
    '   Print #lngSaveFile, "CLIENTLIST=" & ClientList
    '     End If
    'If ReadOnly Then
    '   Print #lngSaveFile, "READONLY=TRUE"
    '     End If
    '     ' --- save various numeric values ---
    'Print #lngSaveFile, "SORTTAGSIZE="  + SortTagSize))
    'Print #lngSaveFile, "SORTLINECOUNT="  + SortLineCount))
    'Print #lngSaveFile, "FETCHLINECOUNT="  + FetchLineCount))
    '     ' --- save memory sort data ---
    '     If MEM(MemPos_SortState) <> 0 And SortFileName = "" Then
    '         For lngLoop = 1 To SortLineCount
    '      Print #lngSaveFile, "SORTTAGS_" & Format$(lngLoop, "000") & "=" & SortTags(lngLoop)
    '      Print #lngSaveFile, "SORTINDEX_" & Format$(lngLoop, "000") & "="  + SortIndex(lngLoop)))
    '         Next lngLoop
    '     End If
    '     ' --- save IL code (must be last in file) ---
    '     If ProgILCode <> "" Then
    '   Print #lngSaveFile,
    '   Print #lngSaveFile, "ILCODE"
    '   Print #lngSaveFile, ProgILCode;
    '     End If
    '     ' --- done ---
    'Close #lngSaveFile
    ' End Sub

    '    Public Function LoadMemory(ByVal FileName As String) As Boolean
    '        Dim bTemp As Byte
    '        Dim lngPos As Integer
    '        Dim lngLoop As Integer
    '        Dim lngValue As Integer
    '        Dim lngMemPos As Integer
    '        Dim strLine As String
    '        Dim strItem As String
    '        Dim strValue As String
    '        Dim lngSaveFile As Integer
    '        Dim objItem As rtStackEntry
    '        Dim blnSkipRead As Boolean
    '        ' -------------------------
    '        If FileName = "" Then GoTo ErrorFound
    '        If Dir$(FileName) = "" Then
    '            Throw New SystemException("LoadMemory", "File not found: " & FileName)
    '            GoTo ErrorFound
    '        End If
    '        ' --- open memory file ---
    '        On Error GoTo ErrorFound
    '        lngSaveFile = FreeFile()
    '   Open FileName For Input As #lngSaveFile
    '        ' --- load all old memory bytes ---
    '        blnSkipRead = False
    '        For lngMemPos = 0 To TotalMemSize - 1
    '            If lngMemPos Mod 16 = 0 Then
    '                Do
    '            Line Input #lngSaveFile, strLine
    '                Loop Until strLine <> ""
    '            End If
    '            ' --- check if reading an older memory file with a newer program ---
    '            If Not IsHexLine(strLine) Then
    '                blnSkipRead = True
    '                Exit For
    '            End If
    '            bTemp = CByte("&H" & Mid$(strLine, ((lngMemPos Mod 16) * 2) + 1, 2))
    '            MEM(lngMemPos) = bTemp
    '        Next lngMemPos
    '        ' --- get numeric variables from memory ---
    '        For lngLoop = MinNLow To MaxNLow
    '            LET_N_OFS(lngLoop, GetNumeric(MemPos_N + (lngLoop * NumSlotSize), NumSlotSize))
    '        Next lngLoop
    '        For lngLoop = MinNHigh To MaxNHigh
    '            LET_N_OFS(lngLoop, GetNumeric(MemPos_N64 + ((lngLoop - MinNHigh) * NumSlotSize), NumSlotSize))
    '        Next lngLoop
    '        REC = GetNumeric(MemPos_Rec, NumSlotSize)
    '        REMVAL = GetNumeric(MemPos_RemVal, NumSlotSize)
    '        ' --- clear stacks and variables ---
    '        Do While GosubStack.Count > 0
    '            GosubStack.Remove(1)
    '        Loop
    '        Do While Breakpoints.Count > 0
    '            Breakpoints.Remove(1)
    '        Loop
    '        ProgILCode = ""
    '        ' --- check for any names/paths/stack entries ---
    '        Do While Not EOF(lngSaveFile)
    '            If Not blnSkipRead Then
    '         Line Input #lngSaveFile, strLine
    '            End If
    '            blnSkipRead = False
    '            lngPos = InStr(strLine, "=") ' position of first equal sign
    '            If lngPos = 0 Then
    '                If UCase$(strLine) = "ILCODE" Then
    '                    ' --- everything else in the file is IL code ---
    '                    GoTo GetILCode
    '                End If
    '                GoTo NextItem
    '            End If
    '            strItem = UCase$(Left$(strLine, lngPos - 1)) ' line's item
    '            strValue = Mid$(strLine, lngPos + 1) ' line's value
    '            ' --- load current instruction pointer (but don't use it yet) ---
    '            If strItem = "CURRIP" Then
    '                CurrIP = strValue
    '                GoTo NextItem
    '            End If
    '            ' --- gosub stack entry ---
    '            If Left$(strItem, 11) = "GOSUBSTACK_" Then
    '                objItem = New rtStackEntry
    '                objItem.FromString(strValue)
    '                GosubStack.Add(objItem)
    '                objItem = Nothing
    '                GoTo NextItem
    '            End If
    '            ' --- breakpoint entry ---
    '            If Left$(strItem, 11) = "BREAKPOINT_" Then
    '                objItem = New rtStackEntry
    '                objItem.FromString(strValue)
    '                Breakpoints.Add(objItem)
    '                objItem = Nothing
    '                GoTo NextItem
    '            End If
    '            ' --- channel path ---
    '            If Left$(strItem, 9) = "CHANPATH_" Then
    '                lngValue = Val(Mid$(strItem, 10, 3))
    '                ChannelPaths(lngValue) = strValue
    '                GoTo NextItem
    '            End If
    '            ' --- keyboard buffer ---
    '            If strItem = "KBDBUFFER" Then
    '                KBuff.FromString(strValue)
    '                GoTo NextItem
    '            End If
    '            ' --- load various string values ---
    '            If strItem = "LOGINID" Then
    '                LoginID = LCase$(strValue)
    '                GoTo NextItem
    '            End If
    '            If strItem = "MACHINENAME" Then
    '                MachineName = LCase$(strValue)
    '                GoTo NextItem
    '            End If
    '            If strItem = "PRINTERFILENAME" Then
    '                PrinterFileName = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "PRINTERTYPE" Then
    '                PrinterType = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "PRINTERDEVICENAME" Then
    '                PrinterDeviceName = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "PRINTERPARAMETERS" Then
    '                PrinterParameters = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "SORTFILENAME" Then
    '                SortFileName = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "SQLSUBQUERY" Then
    '                SQLSubQuery = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "SQLSUBQUERYFILE" Then
    '                SQLSubQueryFile = strValue
    '                GoTo NextItem
    '            End If
    '            If strItem = "CLIENTLIST" Then
    '                If ClientList = "" Then
    '                    ClientList = strValue
    '                    If ClientList <> "" Then
    '                        DebugMessage("ClientList = """ & ClientList & """")
    '                        PrepareClientLists()
    '                    End If
    '                End If
    '                GoTo NextItem
    '            End If
    '            If strItem = "READONLY" Then
    '         If Not ReadOnly Then
    '            ReadOnly = (UCase$(strValue) = "TRUE")
    '                    DebugMessage("ReadOnly = """ & UCase$(strValue) & """")
    '                End If
    '            End If
    '            ' --- load various numeric values ---
    '            If strItem = "SORTTAGSIZE" Then
    '                SortTagSize = Val(strValue)
    '                GoTo NextItem
    '            End If
    '            If strItem = "SORTLINECOUNT" Then
    '                SortLineCount = Val(strValue)
    '                GoTo NextItem
    '            End If
    '            If strItem = "FETCHLINECOUNT" Then
    '                FetchLineCount = Val(strValue)
    '                GoTo NextItem
    '            End If
    '            ' --- load memory sort data ---
    '            If Left$(strItem, 9) = "SORTTAGS_" Then
    '                lngValue = Val(Mid$(strItem, 10, 3))
    '                SortTags(lngValue) = strValue
    '                GoTo NextItem
    '            End If
    '            If Left$(strItem, 10) = "SORTINDEX_" Then
    '                lngValue = Val(Mid$(strItem, 11, 3))
    '                SortIndex(lngValue) = Val(strValue)
    '                GoTo NextItem
    '            End If
    '            ' --- get next item ---
    'NextItem:
    '        Loop
    '        GoTo Done
    '        ' --- get IL code ---
    'GetILCode:
    '        ProgILCode = ""
    '        Do While Not EOF(lngSaveFile)
    '      Line Input #lngSaveFile, strLine
    '            ProgILCode = ProgILCode & strLine & vbCrLf
    '        Loop
    '        GoTo Done
    '        ' --- done ---
    'Done:
    '   Close #lngSaveFile
    '        DoEvents()
    '        ' --- delete file after loaded ---
    '        If Not InsideIDE And Not DebugFlag Then
    '            On Error Resume Next
    '            Kill(FileName)
    '            On Error GoTo 0
    '        End If
    '        ' --- done ---
    '        LoadMemory = True
    '        Exit Function
    'ErrorFound:
    '        ' --- error ---
    '        On Error Resume Next
    '   Close #lngSaveFile
    '        DoEvents()
    '        On Error GoTo 0
    '        LoadMemory = False
    '    End Function

    '    Public Function GetMemPage(ByVal Offset As Integer) As String
    '        Dim LoopX As Integer
    '        Dim LoopY As Integer
    '        Dim bChar As Byte
    '        Dim strResult As String
    '        ' ---------------------
    '        ' --- change page number to byte number ---
    '        If Offset < 256 Then Offset = Offset * 256
    '        ' --- build page map ---
    '        strResult = ""
    '        For LoopY = 0 To 15
    '            For LoopX = 0 To 15
    '                bChar = MEM(Offset + (LoopY * 16) + LoopX)
    '                strResult = strResult & HexChar(bChar)
    '                strResult = strResult & " "
    '            Next LoopX
    '            strResult = strResult & "  "
    '            For LoopX = 0 To 15
    '                bChar = MEM(Offset + (LoopY * 16) + LoopX)
    '                bChar = ModPos(bChar, 128)
    '                If bChar < 32 Or bChar > 126 Then bChar = 46 ' period
    '                strResult = strResult & Chr$(bChar)
    '            Next LoopX
    '            strResult = strResult & vbCrLf
    '        Next LoopY
    '        GetMemPage = strResult
    '    End Function

    '    Public Function GetMemPageLabels(ByVal Offset As Integer) As String
    '        Dim LoopX As Integer
    '        Dim LoopY As Integer
    '        Dim bChar As Byte
    '        Dim strResult As String
    '        ' ---------------------
    '        ' --- change page number to byte number ---
    '        If Offset < 256 Then Offset = Offset * 256
    '        ' --- build page map ---
    '        strResult = "    "
    '        For LoopX = 0 To 15
    '            strResult = strResult & " " & HexChar(LoopX)
    '        Next LoopX
    '        strResult = strResult & vbCrLf & "     " & String$(47, "-") & vbCrLf
    '        For LoopY = 0 To 15
    '            strResult = strResult & HexChar(LoopY * 16) & " | "
    '            For LoopX = 0 To 15
    '                bChar = MEM(Offset + (LoopY * 16) + LoopX)
    '                strResult = strResult & HexChar(bChar)
    '                strResult = strResult & " "
    '            Next LoopX
    '            strResult = strResult & "  "
    '            For LoopX = 0 To 15
    '                bChar = MEM(Offset + (LoopY * 16) + LoopX)
    '                bChar = ModPos(bChar, 128)
    '                If bChar < 32 Or bChar > 126 Then bChar = 46 ' period
    '                strResult = strResult & Chr$(bChar)
    '            Next LoopX
    '            strResult = strResult & vbCrLf
    '        Next LoopY
    '        GetMemPageLabels = strResult
    '    End Function

    'Public Function To_Byte(ByVal Value As Int64) As Byte
    '    If Value >= 0 And Value < 256 Then
    '        Return CByte(Value)
    '    Else
    '        Return CByte(Value - (Int(Value / 256) * 256))
    '    End If
    'End Function

    Public Function AlphaLen(ByVal Value As String) As Integer
        If Value = "" Then Return 1 ' one byte
        Return Value.Length
    End Function

    Public Function InBufferSpace(ByVal Position As Integer, ByVal Size As Integer) As Boolean
        ' --- check if completely outside memory ---
        If Position < 0 Then Return False
        If Not MEMTF(MemPos_TBAlloc) Then
            If Position + Size > MemPos_TrackBuffer Then Return False
        Else
            If Position + Size > MemPos_TrackBuffer + (32 * 256) Then Return False
        End If
        ' --- if #2 bit of PRIVG is not set, must be within a buffer ---
        If MEM(MemPos_Privg) Mod 2 = 0 Then
            If Position < MemPos_R Then Return False
            If Position + Size > MemPos_W + 256 And Position < MemPos_S Then Return False
            If Position + Size > MemPos_V + 256 Then Return False
        End If
        ' --- are within buffer space ---
        Return True
    End Function

    Public Sub LET_MEMTF(ByVal Offset As Integer, ByVal Value As Boolean)
        If Value Then
            MEM(Offset) = TRUEVAL
        Else
            MEM(Offset) = FALSEVAL
        End If
    End Sub

    Public Function MEMTF(ByVal Offset As Integer) As Boolean
        MEMTF = (MEM(Offset) <> FALSEVAL)
    End Function

    '    Public Function IsHexLine(ByVal Value As String) As Boolean
    '        Dim lngLoop As Integer
    '        Dim strChar As String
    '        ' -------------------
    '        IsHexLine = False
    '        If Len(Value) <> 32 And Len(Value) <> 51 Then Exit Function
    '        For lngLoop = 0 To 31
    '            strChar = Mid$(Value, lngLoop + 1, 1)
    '            If strChar < "0" Or strChar > "9" Then
    '                If strChar < "A" Or strChar > "F" Then
    '                    Exit Function
    '                End If
    '            End If
    '        Next
    '        IsHexLine = True
    '    End Function

    '    Public Function GetHexAlpha(ByVal MemPos As Integer, ByVal HexLen As Integer) As String
    '        Dim lngLoop As Integer
    '        Dim strResult As String
    '        ' ---------------------
    '        strResult = ""
    '        For lngLoop = MemPos To MemPos + HexLen - 1
    '            strResult = strResult & HexChar(MEM(lngLoop))
    '        Next
    '        GetHexAlpha = strResult
    '    End Function

    '    Public Function GetGosubStackText() As String
    '        Dim lngLoop As Integer
    '        Dim strTemp As String
    '        Dim strResult As String
    '        Dim objItem As rtStackEntry
    '        ' -------------------------
    '        strResult = ""
    '        lngLoop = 1
    '        Do While lngLoop <= GosubStack.Count
    '            objItem = GosubStack.Item(lngLoop)
    '            strTemp = objItem.ToString
    '            If strResult <> "" And strTemp <> "" Then strResult = strResult & vbCrLf
    '            strResult = strResult & strTemp
    '            lngLoop = lngLoop + 1
    '        Loop
    '        GetGosubStackText = strResult
    '    End Function

    '    Public Function GetBreakpointsText() As String
    '        Dim lngLoop As Integer
    '        Dim strTemp As String
    '        Dim strResult As String
    '        Dim objItem As rtStackEntry
    '        ' -------------------------
    '        strResult = ""
    '        lngLoop = 1
    '        Do While lngLoop <= Breakpoints.Count
    '            objItem = Breakpoints.Item(lngLoop)
    '            strTemp = objItem.ToString
    '            If strResult <> "" And strTemp <> "" Then strResult = strResult & vbCrLf
    '            strResult = strResult & strTemp
    '            lngLoop = lngLoop + 1
    '        Loop
    '        GetBreakpointsText = strResult
    '    End Function

End Module
