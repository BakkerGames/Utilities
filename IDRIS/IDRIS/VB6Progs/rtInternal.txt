Attribute VB_Name = "rtInternal"
' -------------------------------
' --- rtInternal - 02/09/2010 ---
' -------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 02/09/2010 - SBakker - URD 11076
'            - Make sure .CursorLocation is always the first property set.
' 08/19/2009 - SBakker - URD 10076
'            - Ignore multiple ReleaseAppLock errors.
' 08/14/2009 - SBakker - URD 11076
'            - Changed locked records to be read with adUseServer. Adding the
'              handling of History files throws errors using adUseClient.
' 10/13/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
'            - Made changes recommended by CodeAdvisor.
' 07/28/2008 - SBAKKER - URD 11127
'            - During file initialize, truncate the table if it's empty after
'              deleting all the matching records.
' 06/16/2008 - SBAKKER - URD 11118
'            - Added USER and SERVER to ThrowError for debugging purposes.
' 04/24/2008 - SBAKKER - URD 11077
'            - Added enhanced error messages for invalid dates during Unpack.
' 10/15/2007 - SBAKKER - URD 11036
'            - Clear SQL errors before trying again to get a User Number. This
'              will prevent "Violation of Primary Key" errors on [%USER].
' 07/18/2007 - SBAKKER - URD 10969
'            - Change UpdateUserInfo to check LOGINID to be NULL or equal to the
'              current LoginID. Otherwise it will fail, and a new user number
'              will be assigned. Reuse the same user number if possible.
' 06/06/2007 - SBAKKER - URD 9739
'            - Added extra check for General network errors. Changed text to
'              indicate a SQL connection error instead.
' 02/08/2007 - SBAKKER - URD 10898
'            - Added UserAuthorized function to prevent users from accessing
'              IDRIS during EOM Calc time.
' 01/22/2007 - SBAKKER - URD 9739
'            - Added error checking on the cnSQL object before using it, to
'              prevent hanging runtime processes.
' 07/13/2006 - Added check for "APPLICATION SERVEROK". This should only happen
'              if a runtime tries to run without being called, so it now is
'              treated just like "APPLICATION END".
'            - Added "ON ERROR RESUME NEXT" inside ThrowError. This lets it
'              continue processing even if something goes wrong. (Had problems
'              with writing to the Event Log.)
' 06/29/2006 - Enhanced MakeLibExe to check for executables within an UPDATE
'              directory that are newer than those in the regular directory. Run
'              whichever one is newer.
' 02/08/2006 - Added "DEBUG | BREAKPOINT" to list of defined commands. It will
'              add a breakpoint to the list, but will not start the program
'              running.
' 02/06/2006 - Moved To_Byte, AlphaLen, InBufferSpace, GetGosubStackText, and
'              MEMTF to rtMemory so they are available in IDRISClient.
' 02/03/2006 - Updated SysRel/SysRev to 5.1, to indicate new registers and new
'              memory map.
' 01/30/2006 - Added clearing of Debug variables. Added processing of DEBUG
'              messages from the client program. Added "AddBreakPoint" routine.
' 01/30/2006 - Removed "DebugMessage '*** SWITCHING...'" messages. The problem
'              these were tracking has been corrected and is not needed anymore.
' 01/20/2006 - Added better handling in LetGReg. Checking for G=USER is now done
'              in LET_Gx instead of in LetGReg.
' 11/22/2005 - Clear internal vars properly in InitRuntime and ExecuteEscape.
' 10/28/2005 - VOL=255 allows the file to be opened across all volumes. Added
'              support for parts of SQL WHERE clause not existing (such as
'              DEVICE and VOLUME).
' ------------------------------------------------------------------------------

Public Sub InitRuntime()
   Dim lngLoop As Long
   ' -----------------
   ' --- run each time a runtime is started ---
   ReadyToRun = False
   MUSTEXIT = False
   EXITING = False
   SWITCHING = False
   WAITTOEXIT = False
   ERRORTHROWN = False
   NUMERIC_RESULT = 0
   ALPHA_RESULT = ""
   UPDATE_VALUE = 0
   FREEZE_LENGTH = False
   HasLockedRec = False
   LockFlag = False
   LockedResource = ""
   LockedFileNum = -1
   LockedRecNum = -1
   LockedRecLen = -1
   SQLSubQuery = ""
   SQLSubQueryFile = ""
   UpdateRecField = False
   Set LockedCadolXref = Nothing
   Set GosubStack = New Collection
   Set KBuff = New rtKbdBuffer
   For lngLoop = 0 To MaxFile
      Set Files(lngLoop) = New rtCadolFile
      Set Files(lngLoop).RecSet = New ADODB.Recordset
   Next lngLoop
   Set rsLockedRec = New ADODB.Recordset
   Set rsGRegs = New ADODB.Recordset
   Set rsRecord = New ADODB.Recordset
   ClientList = ""
   ReadOnly = False
   ' --- debug variables ---
   InBreakMode = False
   BreakFilename = ""
   LastILCode = ""
   DebugOneStep = False
   DebugStepOver = False
   Set Breakpoints = New Collection
   Set Watchpoints = New Collection
End Sub

Public Sub InitMemory()
   ' --- used when starting a new session ---
   MEM(MemPos_MachType) = 20 ' IDRIS
   MEM(MemPos_SysRel) = 5 ' new system release level 5
   MEM(MemPos_SysRev) = 1 ' revision 1
   MEM(MemPos_PrintDev) = 255 ' no printer assigned
   LET_MEMTF MemPos_PrintOn, False
   MEM(MemPos_LocalEdit) = 0               ' not doing local edit
   LET_MEMTF MemPos_ScriptRunFlag, False   ' not running a keyboard script
   LET_MEMTF MemPos_ScriptWriteFlag, False ' not writing a keyboard script
   ' --- initialize all buffer pointers ---
   INIT_R
   INIT_IR
   INIT_Z
   INIT_IZ
   INIT_X
   INIT_IX
   INIT_Y
   INIT_IY
   INIT_W
   INIT_IW
   INIT_S
   INIT_IS
   INIT_T
   INIT_IT
   INIT_U
   INIT_IU
   INIT_V
   INIT_IV
   ' --- fill date with "mm/dd/yyyy" and yyyymmdd ---
   LET_DATEVAL GETDATESTR
   INIT_Z
   LET_Z 4, GETDATEVAL
   INIT_Z
   MOVE 4, MemPos_Z, 0, MemPos_DateVal, 10
   INIT_Z
   LET_Z 4, 0 ' clear out z buffer again
   INIT_Z
End Sub

Public Sub ThrowError(ByVal Source As String, ByVal Msg As String)
   Dim strTemp As String
   Dim lngLoop As Long
   Dim lngErrFile As Long
   Dim strMemDump As String
   Dim strExtraInfo As String
   ' ------------------------
   ' --- prevent multiple errors ---
   If ERRORTHROWN Then Exit Sub
   On Error Resume Next
   ' --- add description from standard error handler ---
   If Err.Description <> "" Then
      Msg = Msg & vbCrLf & Err.Description
   End If
   ' --- add descriptions from sql errors ---
   If Not (cnSQL Is Nothing) Then
      If cnSQL.Errors.Count > 0 Then
         For lngLoop = 0 To cnSQL.Errors.Count - 1
            Msg = Msg & vbCrLf & cnSQL.Errors.Item(lngLoop)
         Next lngLoop
      End If
   End If
   ' --- check if SQL connection error ---
   If InStr(Msg, "General network error") > 0 Then
      Msg = "SQL Connection seems to have errors or has been closed."
   End If
   ' --- add useful info for tracking down error ---
   strExtraInfo = "Current Program: " & _
                  Trim$(Str$(CurrDevNum)) & ":" & _
                  CurrVolName & ":" & _
                  CurrLibName & ":" & _
                  Trim$(Str$(PROG)) & ":" & _
                  Trim$(Str$(CurrJumpPoint))
   strTemp = GetGosubStackText
   If strTemp <> "" Then
      strExtraInfo = strExtraInfo & vbCrLf & "Gosub Stack:"
      strExtraInfo = strExtraInfo & vbCrLf & strTemp
   End If
   strExtraInfo = strExtraInfo & vbCrLf & "KEY = '" & KEY & "', REC = " & Trim$(Str$(REC))
   strExtraInfo = strExtraInfo & vbCrLf & "USER = '" & LoginID & "', SERVER = '" & UCase$(Environ("computername")) & "'"
   ' --- write out error to error log ---
   If ErrorFilename <> "" Then
      ' --- add error to error log ---
      lngErrFile = FreeFile
      Open ErrorFilename For Append As #lngErrFile
      Print #lngErrFile, Format$(Now, "yyyy.mm.dd hh:mm:ss") & " - *** Error: " & _
                         Source & " - " & Msg & " ***"
      Print #lngErrFile, strExtraInfo & vbCrLf
      Close #lngErrFile
      DoEvents
      ' --- also save a memory dump when an error occurs ---
      strMemDump = Left$(ErrorFilename, InStrRev(ErrorFilename, "\")) & _
                   Format$(Now, "yyyymmdd_hhmmss") & "_" & FormatNum("z5", USER) & ".mem"
      DebugMessage "Saving to Memory File """ & strMemDump & """"
      SaveMemory strMemDump
   End If
   ' --- show on debug screen ---
   DebugMessage "*** Error: " & Source & " - " & Msg & " ***"
   DebugMessage strExtraInfo
   ' --- also report the error ---
   App.LogEvent Source & " - " & Msg & vbCrLf & strExtraInfo, vbLogEventTypeError
   DoEvents
   ' --- check if connected to the server ---
   If rtFormMainLoaded Then
      If rtFormMain.wsToServer.State = sckConnected Then
         ' --- send error to server which will forward it to client ---
         If Not MEMTF(MemPos_Background) Then
            ' --- set flag to wait for APPLICATION END message ---
            WAITTOEXIT = True
            SendToServer "APPLICATION" & vbTab & "ERROR" & vbTab & Source & " - " & Msg & vbCrLf & strExtraInfo
         Else
            SendToServer "SERVER" & vbTab & "ERROR" & vbTab & Source & " - " & Msg & vbCrLf & strExtraInfo
         End If
         DoEvents
      Else
         WAITTOEXIT = False
      End If
   Else
      WAITTOEXIT = False
   End If
   ' --- check if debugging in VB6 IDE ---
   If InsideIDE Then
      MsgBox Msg & vbCrLf & strExtraInfo, vbCritical, Source
      If Not DebugFlag Then Debug.Assert False
   End If
   ' --- make sure records are unlocked ---
   UNLOCKREC
   ' --- adjust proper internal flags ---
   MUSTEXIT = True
   EXITING = True
   SWITCHING = False
   ERRORTHROWN = True
   ' --- force an escape ---
   LET_ESCVAL 0 ' perform normal escape
   On Error GoTo 0
End Sub

Public Sub AddGosubStack(ByVal ItemType As Long, ByVal ProgNum As Long, ByVal JumpNum As Long)
   ' --- remember, only one When vector of each type can be on the stack at a time ---
   Dim lngLoop As Long
   Dim objItem As rtStackEntry
   ' -------------------------
   ' --- remove any others with same ItemType ---
   If ItemType <> GOSUB_TYPEVAL Then
      lngLoop = 1
      Do While lngLoop <= GosubStack.Count
         Set objItem = GosubStack.Item(lngLoop)
         If objItem.ItemType = ItemType Then
            GosubStack.Remove lngLoop
         Else
            lngLoop = lngLoop + 1
         End If
      Loop
   End If
   ' --- check for WHEN ERROR TRAP sending JumpNum = -1 ---
   If JumpNum < 0 Then Exit Sub
   ' --- add item to gosub stack ---
   Set objItem = New rtStackEntry
   With objItem
      .ItemType = ItemType
      ' --- check for a call to /USERLIB - it only exists in 0:/SYSVOL ---
      If CurrLibName = "/USERLIB" Then
         .DevNum = 0
         .VolName = "/SYSVOL"
      Else
         .DevNum = CurrDevNum
         .VolName = CurrVolName
      End If
      .LibName = CurrLibName
      .ProgNum = ProgNum
      .JumpNum = JumpNum
   End With
   GosubStack.Add objItem
End Sub

Public Sub ClearGosubStack()
   Do While GosubStack.Count > 0
      GosubStack.Remove 1
   Loop
End Sub

Public Sub ClearGosubStackTo(ByVal ItemType As Long)
   Dim lngLoop As Long
   Dim objItem As rtStackEntry
   ' -------------------------
   lngLoop = GosubStack.Count
   Do While lngLoop >= 1
      Set objItem = GosubStack.Item(lngLoop)
      ' --- look for proper item type ---
      If objItem.ItemType = ItemType Then
         ' --- change to gosub type so it will be used ---
         objItem.ItemType = GOSUB_TYPEVAL
         Exit Sub
      End If
      ' --- wrong type, just remove ---
      GosubStack.Remove lngLoop
      lngLoop = lngLoop - 1
   Loop
End Sub

Public Function HasGosubStackItem(ByVal ItemType As Long) As Boolean
   Dim lngLoop As Long
   Dim objItem As rtStackEntry
   ' -------------------------
   lngLoop = GosubStack.Count
   Do While lngLoop >= 1
      Set objItem = GosubStack.Item(lngLoop)
      If objItem.ItemType = ItemType Then GoTo FoundItem
      lngLoop = lngLoop - 1
   Loop
   HasGosubStackItem = False
   Exit Function
FoundItem:
   HasGosubStackItem = True
End Function

Public Sub SendToServer(ByVal Value As String)
   If Value = "" Then Exit Sub
   ' --- show outgoing messages ---
   If DebugFlag And DebugFlagLevel > 0 Then
      DebugMessage "QUEUED: " & Replace(Value, vbTab, " ")
   End If
   If MEMTF(MemPos_Background) Then
      If UCase$(Left$(Value, 7)) <> "SERVER" & vbTab Then
         ThrowError "SendToServer", "Invalid Background Message: " & Replace(Value, vbTab, " ")
         Exit Sub
      End If
   End If
   ' --- accumulate and send later ---
   PendingOutput = PendingOutput & PackMsg(Value) & vbCrLf
   rtFormMain.TickTimer.Enabled = True
   DoEvents
   If (PendingOutput <> "") And (Not EXITING) Then
      rtFormMain.TickTimer.Enabled = True
      DoEvents
   End If
End Sub

Public Function ModPos(ByVal Value1 As Currency, ByVal Value2 As Currency) As Currency
   ' ----------------------------
   ' --- 0 <= result < value2 ---
   ' ----------------------------
   If Value2 <= 0 Then ' modpos must have positive divisors
      ThrowError "ModPos", "Invalid divisor: " & Trim$(Str$(Value2))
      Exit Function
   End If
   ModPos = Value1 - (Int(Value1 / Value2) * Value2)
End Function

Public Function DivRem(ByVal Value1 As Currency, ByVal Value2 As Currency) As Currency
   Dim curResult As Currency
   ' -----------------------
   If Value2 = 0 Then ' there is no division by zero error
      curResult = 0
      REMVAL = 0
   Else
      curResult = Fix(Value1 / Value2)
      REMVAL = Abs(Value1) - (Abs(curResult) * Abs(Value2))
   End If
   DivRem = curResult
End Function

Public Function FormatNum(ByVal DisplayFmt As String, ByVal Value As Currency) As String

   Dim Result As String
   Dim TempChar As String
   Dim lngLoop As Long
   
   Dim FrontNeg As Boolean
   Dim FrontParen As Boolean
   Dim ZeroFill As Boolean
   Dim StarFill As Boolean
   Dim Comma As Boolean
   Dim DigitsFound As Boolean
   Dim DigitsAbove As Long
   Dim DecimalPoint As Boolean
   Dim DigitsBelow As Long
   Dim RearNeg As Boolean
   Dim RearParen As Boolean
   ' -------------------------
   
   If Value <> Int(Value) Then GoTo ValueError
   If Value > 140737488355327@ Then GoTo ValueError
   If Value < -140737488355328@ Then GoTo ValueError
   
   ' --- clear flags ---
   FrontNeg = False
   FrontParen = False
   ZeroFill = False
   StarFill = False
   Comma = False
   DigitsFound = False
   DigitsAbove = 0
   DecimalPoint = False
   DigitsBelow = 0
   RearNeg = False
   RearParen = False
   
   Result = ""
   
   ' --- parse display format into flags ---
   For lngLoop = 1 To Len(DisplayFmt)
      TempChar = Mid$(DisplayFmt, lngLoop, 1)
      Select Case TempChar
         Case "("
            If DigitsFound Or FrontParen Or FrontNeg Then GoTo ErrorFound
            FrontParen = True
         Case "-"
            If Not DigitsFound Then
               If FrontParen Or FrontNeg Then GoTo ErrorFound
               FrontNeg = True
            Else
               If RearNeg Or RearParen Then GoTo ErrorFound
               RearNeg = True
            End If
         Case "*"
            If DigitsFound Or StarFill Or ZeroFill Then GoTo ErrorFound
            StarFill = True
         Case "z", "Z"
            If DigitsFound Or StarFill Or ZeroFill Then GoTo ErrorFound
            ZeroFill = True
         Case ","
            If DigitsFound Or Comma Then GoTo ErrorFound
            Comma = True
         Case "0" To "9"
            If Not DecimalPoint Then
               DigitsAbove = (DigitsAbove * 10) + Val(TempChar)
            Else
               If RearNeg Or RearParen Then GoTo ErrorFound
               DigitsBelow = (DigitsBelow * 10) + Val(TempChar)
            End If
            DigitsFound = True
         Case "."
            If Not DigitsFound Then GoTo ErrorFound
            If DecimalPoint Then GoTo ErrorFound
            DecimalPoint = True
         Case ")"
            If Not DigitsFound Then GoTo ErrorFound
            If RearNeg Or RearParen Then GoTo ErrorFound
            RearParen = True
         Case Else
            GoTo ErrorFound
      End Select
   Next lngLoop
   
   ' --- check for digit overflows ---
   If DigitsAbove + DigitsBelow > 14 Then GoTo ErrorFound
   If DigitsBelow > 7 Then GoTo ErrorFound
   
   ' --- clean up decimal point ---
   If DigitsBelow = 0 Then DecimalPoint = False
   
   ' --- clean up negative display ---
   If FrontParen Or RearParen Then
      FrontParen = True
      RearParen = True
      FrontNeg = False
      RearNeg = False
   End If
   If RearNeg Then
      FrontNeg = False
   End If
   
   ' --- build result ---
   Result = Trim$(Str$(Int(Abs(Value))))
   ' --- chop off if too long ---
   If Len(Result) > DigitsAbove + DigitsBelow Then
      Result = Right$(Result, DigitsAbove + DigitsBelow)
   End If
   ' --- add spaces if too short ---
   If Len(Result) < DigitsAbove + DigitsBelow Then
      Result = Space(DigitsAbove + DigitsBelow - Len(Result)) & Result
   End If
   ' --- fill in zeros to right of decimal point ---
   If DecimalPoint Then
      If Len(Trim$(Right$(Result, DigitsBelow))) < DigitsBelow Then
         Result = Left$(Result, DigitsAbove) & _
                  Right$("0000000" & Trim$(Right$(Result, DigitsBelow)), DigitsBelow)
      End If
   End If
   If Comma Then
      If DigitsAbove > 3 Then
         Result = Left$(Result, DigitsAbove - 3) & "," & Mid$(Result, DigitsAbove - 2)
      End If
      If DigitsAbove > 6 Then
         Result = Left$(Result, DigitsAbove - 6) & "," & Mid$(Result, DigitsAbove - 5)
      End If
      If DigitsAbove > 9 Then
         Result = Left$(Result, DigitsAbove - 9) & "," & Mid$(Result, DigitsAbove - 8)
      End If
      If DigitsAbove > 12 Then
         Result = Left$(Result, DigitsAbove - 12) & "," & Mid$(Result, DigitsAbove - 11)
      End If
      Result = Replace(Result, " ,", "  ")
      DigitsAbove = DigitsAbove + ((DigitsAbove - 1) \ 3) ' add for the commas
   End If
   If ZeroFill Then
      Result = Right$(String$(DigitsAbove, "0") & Trim$(Result), DigitsBelow + DigitsAbove)
   End If
   If StarFill Then
      Result = Right$(String$(DigitsAbove, "*") & Trim$(Result), DigitsBelow + DigitsAbove)
   End If
   If DecimalPoint Then
      Result = Left$(Result, DigitsAbove) & "." & Right$(Result, DigitsBelow)
      DigitsBelow = DigitsBelow + 1 ' add for the decimal point
   End If
   If FrontNeg Or FrontParen Then
      If Value < 0 Then
         If FrontNeg Then
            Result = "-" & Trim$(Result)
         Else
            Result = "(" & Trim$(Result)
         End If
         DigitsAbove = DigitsAbove + 1
         Result = Space(DigitsAbove + DigitsBelow - Len(Result)) & Result
      Else
         Result = " " & Result
         DigitsAbove = DigitsAbove + 1
      End If
   End If
   If RearNeg Or RearParen Then
      If Value < 0 Then
         If RearNeg Then
            Result = Result & "-"
         Else
            Result = Result & ")"
         End If
         DigitsBelow = DigitsBelow + 1
      Else
         Result = Result & " "
         DigitsBelow = DigitsBelow + 1
      End If
   End If
   
   FormatNum = Result
   
   Exit Function
   
ErrorFound:

   ThrowError "FormatNum", "Invalid display format: """ & DisplayFmt & """"
   Exit Function
   
ValueError:

   ThrowError "FormatNum", "Invalid numeric value: " & Trim$(Str$(Value))
   Exit Function
   
End Function

Public Sub CheckFormFeed()
   ' --- print a formfeed if one is pending ---
   If MEMTF(MemPos_FFPending) Then
      ' --- finish last line ---
      If MEMTF(MemPos_LineHasData) Then
         Print #PrinterFileNum,
      End If
      Print #PrinterFileNum, Chr$(12); ' Ctrl-L
      LET_MEMTF MemPos_FFPending, False
      LET_MEMTF MemPos_PageHasData, False
      LET_MEMTF MemPos_LineHasData, False
   End If
End Sub

Public Sub ExecuteEscape()
   UNLOCKREC
   RELEASEDEVICE
   PRINTOFF
   GRAPHOFF
   CloseAllChannels
   CloseSortFile
   ClearGosubStack
   LET_MEMTF MemPos_TBAlloc, False
   MEM(MemPos_EscVal) = 0
   MEM(MemPos_CanVal) = 0
   MEM(MemPos_Status) = 0
   SQLSubQuery = ""
   SQLSubQueryFile = ""
   UpdateRecField = False
   If MEMTF(MemPos_Background) Then
      EXITRUNTIME
   End If
End Sub

Public Sub CloseSortFile()
   ' --- close any open sort file ---
   If MEM(MemPos_SortState) > 1 And MEM(MemPos_SortState) < 6 Then
      Close #SortFileNum
      DoEvents
      SortFileNum = 0
   End If
   ' --- delete sort file ---
   If SortFileName <> "" Then
      On Error Resume Next
      Kill SortFileName
      On Error GoTo 0
   End If
   ' --- reset variables ---
   SortFileName = ""
   SortFileNum = 0
   SortTagSize = 0
   SortLineCount = 0
   FetchLineCount = 0
   MEM(MemPos_SortState) = 0 ' not sorting
End Sub

Public Sub CloseAllChannels()
   Dim lngLoop As Long
   ' -----------------
   For lngLoop = 0 To MaxChannel
      CLOSECHANNEL lngLoop
   Next lngLoop
   MEM(MemPos_Status) = 0 ' ok
End Sub

Public Function ConvertNum(ByVal EnterFmt As String, ByVal Value As String, ByRef Result As Currency) As Boolean

   Dim lngLoop As Long
   Dim lngLen As Long
   Dim blnNegSign As Boolean
   Dim lngDigitsAbove As Long
   Dim lngDigitsBelow As Long
   Dim blnFoundNeg As Boolean
   Dim curValAbove As Currency
   Dim blnFoundPoint As Boolean
   Dim curValBelow As Currency
   Dim lngTempDigitsAbove As Long
   Dim lngTempDigitsBelow As Long
   Dim curResult As Currency
   ' -------------------------------
   
   ' --- split format into pieces ---
   If Not SplitEnterFmt(EnterFmt, blnNegSign, lngDigitsAbove, lngDigitsBelow) Then
      GoTo InvalidVal
   End If
   
   ' --- prepare parts ---
   blnFoundNeg = False
   curValAbove = 0
   blnFoundPoint = False
   curValBelow = 0
   
   ' --- find length of string to convert ---
   lngLen = lngDigitsAbove + lngDigitsBelow        ' get total digits above plus below
   If blnNegSign Then lngLen = lngLen + 1          ' add one for negative sign
   If lngDigitsBelow > 0 Then lngLen = lngLen + 1  ' add one for decimal point
   If lngLen > Len(Value) Then lngLen = Len(Value) ' make less if string is shorter
   
   ' --- save number of digits for countdown check ---
   lngTempDigitsAbove = lngDigitsAbove
   lngTempDigitsBelow = lngDigitsBelow
   
   ' --- split out value into parts ---
   For lngLoop = 1 To lngLen
      Select Case Mid$(Value, lngLoop, 1)
         Case " "
            ' --- ignore spaces ---
         Case ","
            ' --- ignore commas ---
         Case "-"
            If Not blnNegSign Then GoTo InvalidVal
            If blnFoundNeg Then GoTo InvalidVal
            blnFoundNeg = True
         Case "."
            If blnFoundPoint Then GoTo InvalidVal
            If lngDigitsBelow = 0 Then GoTo InvalidVal
            blnFoundPoint = True
         Case "0" To "9"
            If Not blnFoundPoint Then
               curValAbove = (curValAbove * 10) + Val(Mid$(Value, lngLoop, 1))
               lngTempDigitsAbove = lngTempDigitsAbove - 1
               If lngTempDigitsAbove < 0 Then GoTo InvalidVal
            Else
               curValBelow = (curValBelow * 10) + Val(Mid$(Value, lngLoop, 1))
               lngTempDigitsBelow = lngTempDigitsBelow - 1
               If lngTempDigitsBelow < 0 Then GoTo InvalidVal
            End If
         Case Else
            GoTo InvalidVal
      End Select
   Next lngLoop

   ' --- shift digits below by unused number of places ---
   For lngLoop = 1 To lngTempDigitsBelow
      curValBelow = curValBelow * 10
   Next lngLoop
   
   ' --- get actual result value ---
   curResult = curValAbove
   For lngLoop = 1 To lngDigitsBelow
      curResult = curResult * 10
   Next lngLoop
   curResult = curResult + curValBelow
   If blnFoundNeg Then curResult = -curResult
   
   ' --- put value into target ---
   Result = curResult
   
   ConvertNum = True
   Exit Function
   
InvalidVal:

   ConvertNum = False

End Function

Public Function SplitEnterFmt(ByVal EnterFmt As String, _
                              ByRef NegSign As Boolean, _
                              ByRef DigitsAbove As Long, _
                              ByRef DigitsBelow As Long) As Boolean

   Dim lngLoop As Long
   Dim blnFoundPoint As Boolean
   ' --------------------------
   
   NegSign = False
   DigitsAbove = 0
   DigitsBelow = 0
   blnFoundPoint = False
   
   For lngLoop = 1 To Len(EnterFmt)
      Select Case Mid$(EnterFmt, lngLoop, 1)
         Case "-"
            If lngLoop <> 1 Then GoTo InvalidFmt
            If NegSign Then GoTo InvalidFmt
            NegSign = True
         Case "."
            If blnFoundPoint Then GoTo InvalidFmt
            blnFoundPoint = True
         Case "0" To "9"
            If Not blnFoundPoint Then
               DigitsAbove = (DigitsAbove * 10) + Val(Mid$(EnterFmt, lngLoop, 1))
            Else
               DigitsBelow = (DigitsBelow * 10) + Val(Mid$(EnterFmt, lngLoop, 1))
            End If
         Case Else
            GoTo InvalidFmt
      End Select
   Next lngLoop

   If DigitsAbove + DigitsBelow > 14 Then GoTo InvalidFmt
   If DigitsBelow > 7 Then GoTo InvalidFmt
   
   SplitEnterFmt = True
   Exit Function
   
InvalidFmt:

   SplitEnterFmt = False

End Function

Public Function AdjustFilenameWindows(ByVal Value As String) As String
   ' --- Don't use Percent (%) in Windows world ---
   AdjustFilenameWindows = AdjustFilenameSQL(Replace(Value, "/", "_"))
End Function

Public Function AdjustFilenameSQL(ByVal Value As String) As String
   Dim lngLoop As Long
   Dim strResult As String
   ' ---------------------
   strResult = ""
   For lngLoop = 1 To Len(Value)
      Select Case Mid$(Value, lngLoop, 1)
         ' --- change slash to underline. percent causes problems. ---
         Case "/"
            strResult = strResult & "%"
         ' --- change double-quote to single-quote ---
         Case """"
            strResult = strResult & "'"
         ' --- some special chars must be converted to underline ---
         Case "\", ":", "*", "?", "<", ">", "|"
            strResult = strResult & "_"
         ' --- all others allowed ---
         Case Else
            strResult = strResult & Mid$(Value, lngLoop, 1)
      End Select
   Next lngLoop
   ' --- return adjusted name ---
   AdjustFilenameSQL = Trim$(strResult)
End Function

Public Sub AddSortTag(ByVal SortTag As String)
   Dim lngLoop As Long
   ' -----------------
   ' --- check if wrong sort state ---
   If MEM(MemPos_SortState) < 2 Or MEM(MemPos_SortState) > 3 Then
      SYSERROR "SORT NOT INITIALIZED"
      Exit Sub
   End If
   ' --- check if wrong sort tag length ---
   If MEM(MemPos_SortState) = 2 Then
      SortTagSize = Len(SortTag)
   ElseIf SortTagSize <> Len(SortTag) Then
      SYSERROR "TAG SIZE NOT CONSISTANT"
      Exit Sub
   End If
   ' --- check if switching from memory sort to file sort ---
   If SortFileName = "" And SortLineCount >= MAXMEMTAGS Then
      SortFileName = GetTempFile(TempPath, "SMP")
      SortFileNum = FreeFile
      Open SortFileName For Output As #SortFileNum
      For lngLoop = 1 To SortLineCount
         Print #SortFileNum, SortTags(lngLoop)
      Next lngLoop
   End If
   ' --- add sort tag ---
   If SortFileName = "" Then
      ' --- memory sort ---
      SortTags(SortLineCount + 1) = SortTag & FormatNum("z14", REC)
      SortIndex(SortLineCount + 1) = SortLineCount + 1
   Else
      ' --- file sort ---
      Print #SortFileNum, SortTag & FormatNum("z14", REC)
   End If
   ' --- done ---
   SortLineCount = SortLineCount + 1
   MEM(MemPos_SortState) = 3 ' have sorted at least one tag
End Sub

Public Function InvalidFilename(ByVal Value As String) As Boolean
   ' --- this checks for all invalid characters for a filename --
   Dim lngLoop As Long
   Dim strChar As String
   ' --------------------
   If Trim$(Value) = "" Then GoTo ErrorFound ' filenames can't be blank
   For lngLoop = 1 To Len(Value)
      strChar = Mid$(Value, lngLoop, 1)
      If Asc(strChar) < 32 Or Asc(strChar) = 127 Then GoTo ErrorFound ' control chars
      If strChar = ":" Then GoTo ErrorFound
      If strChar = "\" Then GoTo ErrorFound
      If strChar = "<" Then GoTo ErrorFound
      If strChar = ">" Then GoTo ErrorFound
      If strChar = "|" Then GoTo ErrorFound
      If strChar = "?" Then GoTo ErrorFound
      If strChar = "*" Then GoTo ErrorFound
   Next lngLoop
   InvalidFilename = False
   Exit Function
ErrorFound:
   InvalidFilename = True
End Function

Public Sub InitBuffer(ByVal BuffStart As Long)
   ' --- move pointer to beginning of buffer ---
   Select Case BuffStart
      Case MemPos_R: INIT_R: INIT_IR
      Case MemPos_Z: INIT_Z: INIT_IZ
      Case MemPos_X: INIT_X: INIT_IX
      Case MemPos_Y: INIT_Y: INIT_IY
      Case MemPos_W: INIT_W: INIT_IW
      Case MemPos_S: INIT_S: INIT_IS
      Case MemPos_T: INIT_T: INIT_IT
      Case MemPos_U: INIT_U: INIT_IU
      Case MemPos_V: INIT_V: INIT_IV
      Case Else
         ThrowError "InitBuffer", "Invalid Buffer: " & Trim$(Str$(BuffStart))
         Exit Sub
   End Select
End Sub

Public Function BufferPos(ByVal BuffStart As Long) As Long
   ' --- move pointer to beginning of buffer ---
   Select Case BuffStart
      Case MemPos_R: BufferPos = (MEM(MemPos_RP2) * 256) + MEM(MemPos_RP)
      Case MemPos_Z: BufferPos = (MEM(MemPos_ZP2) * 256) + MEM(MemPos_ZP)
      Case MemPos_X: BufferPos = (MEM(MemPos_XP2) * 256) + MEM(MemPos_XP)
      Case MemPos_Y: BufferPos = (MEM(MemPos_YP2) * 256) + MEM(MemPos_YP)
      Case MemPos_W: BufferPos = (MEM(MemPos_WP2) * 256) + MEM(MemPos_WP)
      Case MemPos_S: BufferPos = (MEM(MemPos_SP2) * 256) + MEM(MemPos_SP)
      Case MemPos_T: BufferPos = (MEM(MemPos_TP2) * 256) + MEM(MemPos_TP)
      Case MemPos_U: BufferPos = (MEM(MemPos_UP2) * 256) + MEM(MemPos_UP)
      Case MemPos_V: BufferPos = (MEM(MemPos_VP2) * 256) + MEM(MemPos_VP)
      Case Else
         ThrowError "BufferPos", "Invalid Buffer: " & Trim$(Str$(BuffStart))
         Exit Function
   End Select
End Function

Public Function HexString(ByVal Value As String) As String
   Dim strHex As String
   Dim lngLoop As Long
   ' --------------------
   strHex = ""
   For lngLoop = 1 To Len(Value)
      strHex = strHex & HexChar(Asc(Mid$(Value, lngLoop, 1)))
   Next lngLoop
   HexString = strHex
End Function

Public Function HexChar(ByVal Value As Long) As String
   HexChar = Right$("00" & Hex$(Value), 2)
End Function

Public Function LenFormat(ByVal EnterFmt As String) As Long
   Dim lngLoop As Long
   Dim blnFoundPoint As Boolean
   Dim lngResult As Long
   Dim blnNegSign As Boolean
   Dim lngDigitsAbove As Long
   Dim lngDigitsBelow As Long
   Dim blnHasParen As Boolean
   Dim blnHasComma As Boolean
   ' ---------------------------
   blnNegSign = False
   lngDigitsAbove = 0
   lngDigitsBelow = 0
   blnFoundPoint = False
   blnHasParen = False
   blnHasComma = False
   For lngLoop = 1 To Len(EnterFmt)
      Select Case Mid$(EnterFmt, lngLoop, 1)
         Case "(", ")"
            blnHasParen = True
         Case ","
            blnHasComma = True
         Case "-"
            If lngLoop <> 1 Then GoTo InvalidFmt
            If blnNegSign Then GoTo InvalidFmt
            blnNegSign = True
         Case "."
            If blnFoundPoint Then GoTo InvalidFmt
            blnFoundPoint = True
         Case "0" To "9"
            If Not blnFoundPoint Then
               lngDigitsAbove = (lngDigitsAbove * 10) + Val(Mid$(EnterFmt, lngLoop, 1))
            Else
               lngDigitsBelow = (lngDigitsBelow * 10) + Val(Mid$(EnterFmt, lngLoop, 1))
            End If
         Case Else
            GoTo InvalidFmt
      End Select
   Next lngLoop
   If lngDigitsAbove + lngDigitsBelow > 14 Then GoTo InvalidFmt
   If lngDigitsBelow > 7 Then GoTo InvalidFmt
   ' --- find length of string to convert ---
   lngResult = lngDigitsAbove + lngDigitsBelow          ' get total digits above plus below
   If blnNegSign Then lngResult = lngResult + 1         ' add one for negative sign
   If lngDigitsBelow > 0 Then lngResult = lngResult + 1 ' add one for decimal point
   If blnHasParen Then lngResult = lngResult + 2        ' add for both before and after paren
   If blnHasComma And lngDigitsAbove >= 4 Then          ' add commas
      lngResult = lngResult + ((lngDigitsAbove - 1) \ 3)
   End If
   ' --- done ---
   LenFormat = lngResult
   Exit Function
InvalidFmt:
   LenFormat = -1 ' invalid
End Function

Public Function GetCadolXrefs(ByVal FileName As String) As Collection
   ' --- returns all the cadol xrefs for the specified filename ---
   Dim strSQL As String
   Dim oCadolXref As rtCadolXref
   Dim rsCSX As ADODB.Recordset
   Dim MyCadolXrefs As Collection
   ' ----------------------------
   ' --- open and fill recordset ---
   Set rsCSX = New ADODB.Recordset
   Set MyCadolXrefs = New Collection
   strSQL = "SELECT * FROM [%CADOL_SQL_XREF] "
   strSQL = strSQL & "WHERE DATFILENAME = '" & FileName & "' "
   strSQL = strSQL & "ORDER BY CADOLKEY DESC, CADOLBYTE DESC, CADOLVALUE DESC "
   With rsCSX
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
   End With
   ' --- load cadol/sql xrefs ---
   Do While Not rsCSX.EOF
      Set oCadolXref = New rtCadolXref
      With oCadolXref
         .SQLTableName = UCase$(rsCSX!SQLTableName)
         If IsNull(rsCSX!CadolKey) Then
            .CadolKey = ""
         Else
            .CadolKey = rsCSX!CadolKey ' do NOT uppercase this! must be case-sensitive.
         End If
         If IsNull(rsCSX!CadolByte) Then
            .CadolByte = -1 ' nothing
         Else
            .CadolByte = Val(rsCSX!CadolByte)
         End If
         If IsNull(rsCSX!CadolLength) Then
            .CadolLength = -1 ' nothing
         Else
            .CadolLength = Val(rsCSX!CadolLength)
         End If
         If IsNull(rsCSX!CadolValue) Then
            .CadolValue = -1 ' nothing
         Else
            .CadolValue = Val(rsCSX!CadolValue)
         End If
         If IsNull(rsCSX!MultiFile) Then
            .Multiple = False
         Else
            .Multiple = (UCase$(rsCSX!MultiFile) = "Y")
         End If
         Set .DataFormats = GetDataFormats(rsCSX!FormatFileName)
      End With
      ' --- only add here if not a multiple file format ---
      If Not oCadolXref.Multiple Then
         MyCadolXrefs.Add oCadolXref
      End If
      Set oCadolXref = Nothing
      rsCSX.MoveNext
   Loop
   ' --- if found a match, then done ---
   If MyCadolXrefs.Count > 0 Then GoTo Done
   ' --- check for multiple file formats ---
   rsCSX.Close
   strSQL = "SELECT * FROM [%CADOL_SQL_XREF] "
   strSQL = strSQL & "WHERE MULTIFILE = 'Y' "
   strSQL = strSQL & "ORDER BY CADOLKEY DESC, CADOLBYTE DESC, CADOLVALUE DESC "
   With rsCSX
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
   End With
   ' --- load cadol/sql xrefs ---
   Do While Not rsCSX.EOF
      If WildcardMatch(rsCSX!DATFileName, FileName) Then
         Set oCadolXref = New rtCadolXref
         With oCadolXref
            .SQLTableName = UCase$(rsCSX!SQLTableName)
            If IsNull(rsCSX!CadolKey) Then
               .CadolKey = ""
            Else
               .CadolKey = rsCSX!CadolKey ' do NOT uppercase this! must be case-sensitive.
            End If
            If IsNull(rsCSX!CadolByte) Then
               .CadolByte = -1 ' nothing
            Else
               .CadolByte = Val(rsCSX!CadolByte)
            End If
            If IsNull(rsCSX!CadolLength) Then
               .CadolLength = -1 ' nothing
            Else
               .CadolLength = Val(rsCSX!CadolLength)
            End If
            If IsNull(rsCSX!CadolValue) Then
               .CadolValue = -1 ' nothing
            Else
               .CadolValue = Val(rsCSX!CadolValue)
            End If
            If IsNull(rsCSX!MultiFile) Then
               .Multiple = False
            Else
               .Multiple = (UCase$(rsCSX!MultiFile) = "Y")
            End If
            Set .DataFormats = GetDataFormats(rsCSX!FormatFileName)
         End With
         ' --- only add here if it is a multiple file format ---
         If oCadolXref.Multiple Then
            MyCadolXrefs.Add oCadolXref
         End If
         Set oCadolXref = Nothing
      End If
      rsCSX.MoveNext
   Loop
   ' --- done ---
Done:
   rsCSX.Close
   Set rsCSX = Nothing
   If MyCadolXrefs.Count = 0 Then
      Set MyCadolXrefs = Nothing
   End If
   Set GetCadolXrefs = MyCadolXrefs
   Exit Function
ConnError:
   ThrowError "GetCadolXrefs", "SQL Connection Error:"
   Set MyCadolXrefs = Nothing
End Function

Public Function GetDataFormats(ByVal FormatName As String) As Collection
   ' --- returns all the dataformats for the specified filename and sqltable ---
   Dim strSQL As String
   Dim rsDF As ADODB.Recordset
   Dim oDataFormat As rtDataFormat
   Dim MyDataFormats As Collection
   ' -----------------------------
   ' --- open and fill recordset ---
   Set rsDF = New ADODB.Recordset
   Set MyDataFormats = New Collection
   strSQL = "SELECT * FROM [%DATAFORMAT] "
   strSQL = strSQL & "WHERE TABLENAME = '" & FormatName & "' "
   strSQL = strSQL & "ORDER BY FIELDNUMBER "
   With rsDF
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
   End With
   ' --- load data formats ---
   Do While Not rsDF.EOF
      Set oDataFormat = New rtDataFormat
      With oDataFormat
         .FieldNumber = Val(rsDF!FieldNumber)
         .FieldName = UCase$(rsDF!FieldName)
         .CadolType = UCase$(rsDF!CadolType)
         .CadolLength = Val(rsDF!CadolLength)
         If IsNull(rsDF!CadolScale) Then
            .CadolScale = 0
         Else
            .CadolScale = Val(rsDF!CadolScale)
            If .CadolScale < 0 Or .CadolScale > 4 Then
               ThrowError "GetDataFormats", "Invalid Cadol Scale: " & Trim$(Str$(.CadolScale))
               Exit Function
            End If
         End If
         If IsNull(rsDF!CadolValue) Then
            .CadolValue = ""
         Else
            .CadolValue = Trim$(Str$(rsDF!CadolValue))
         End If
      End With
      ' --- add to collection ---
      MyDataFormats.Add oDataFormat
      Set oDataFormat = Nothing
      rsDF.MoveNext
   Loop
   ' --- done ---
   rsDF.Close
   Set rsDF = Nothing
   If MyDataFormats.Count = 0 Then
      Set MyDataFormats = Nothing
   End If
   Set GetDataFormats = MyDataFormats
   Exit Function
ConnError:
   ThrowError "GetDataFormats", "SQL Connection Error:"
   Set GetDataFormats = Nothing
End Function

Public Function FixSqlStr(ByVal Value As String) As String
   ' --- used to replace single quotes inside a sql string ---
   FixSqlStr = Replace(Value, "'", "''")
End Function

Public Function HasLetters(ByVal Value As String) As Boolean
   Dim lngLoop As Long
   Dim bTemp() As Byte
   ' -----------------
   bTemp = Value ' convert to bytes
   For lngLoop = 0 To UBound(bTemp) Step 2
      If (bTemp(lngLoop) >= 65 And bTemp(lngLoop) <= 90) Or _
         (bTemp(lngLoop) >= 97 And bTemp(lngLoop) <= 122) Then
         HasLetters = True
         Exit Function
      End If
   Next lngLoop
   HasLetters = False
End Function

Public Function ReadLockedRec(ByVal Value As Long, ByVal RecNum As Long, ByRef oCadolXref As rtCadolXref) As Boolean
   Dim aData() As Byte
   Dim lngLen As Long
   Dim lngLoop As Long
   Dim strSQL As String
   Dim strWhere As String
   Dim rsResult As ADODB.Recordset
   ' -----------------------------
   ' --- create an application lock on the specified record ---
   LockedResource = Trim$(Str$(Files(Value).Device)) & ":" & _
                    Files(Value).Volume & ":" & _
                    Files(Value).FileName & ":" & _
                    Trim$(Str$(RecNum))
   If Not GetAppLock(LockedResource) Then
      GoTo ErrorFound
   End If
   ' --- build query for specified record ---
   strSQL = ""
   strSQL = strSQL & "SELECT * "
   strSQL = strSQL & "FROM [" & oCadolXref.SQLTableName & "] "
   strSQL = strSQL & "WITH (ROWLOCK) "
   With Files(Value)
      strWhere = "WHERE"
      If .Device <> 255 Then
         strSQL = strSQL & strWhere & " DEVICE = " & Trim$(Str$(.Device)) & " "
         strWhere = "AND"
      End If
      If .AdjVolume <> "" Then
         strSQL = strSQL & strWhere & " VOLUME = '" & .AdjVolume & "' "
         strWhere = "AND"
      End If
      If oCadolXref.Multiple Then
         strSQL = strSQL & strWhere & " FILENAME = '" & .FileName & "' "
         strWhere = "AND"
      End If
      strSQL = strSQL & strWhere & " REC = " & Trim$(Str$(RecNum))
      strWhere = "AND"
   End With
   ' --- read and lock the record ---
   With rsLockedRec
      On Error GoTo SQLError
      ' --- must be adUseServer for trigger issues ---
      .CursorLocation = adUseServer
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .CacheSize = 1
      .MaxRecords = 1
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      On Error GoTo 0
      ' --- check if record unexpectedly missing ---
      If .EOF Then
         .Close
         GoTo ErrorDoRelease
      End If
      ' --- save the locked record's info ---
      REC = .Fields("REC")
      LET_KEY .Fields("KEY")
      aData = .Fields("PACKED_DATA")
      lngLen = UBound(aData) + 1
      LET_LENGTH lngLen
      LockedFileNum = Value
      LockedRecNum = REC
      LockedRecLen = lngLen
      Set LockedCadolXref = oCadolXref
      ' --- init r buffer pointers ---
      INIT_R
      INIT_IR
      ' --- move record to R buffer ---
      For lngLoop = 0 To lngLen - 1
         MEM(MemPos_R + lngLoop) = aData(lngLoop)
      Next lngLoop
   End With
   ' --- indicate that we have a locked record ---
   HasLockedRec = True
   ' --- turn off lock flag, as we have read a locked record ---
   LockFlag = False
   ' --- done ---
   ReadLockedRec = True
   Exit Function
ConnError:
   ThrowError "READLOCKEDREC", "SQL Connection Error:"
   GoTo ErrorDoRelease
SQLError:
   ThrowError "READLOCKEDREC", "Cannot execute SQL query: " & vbCrLf & strSQL
   GoTo ErrorDoRelease
ErrorDoRelease:
   ReleaseAppLock LockedResource
   LockedResource = ""
ErrorFound:
   ReadLockedRec = False
End Function

Public Function InitializeFile(ByVal Value As Long) As Boolean
   Dim strSQL As String
   Dim strWhere As String
   Dim oCadolXref As rtCadolXref
   ' ---------------------------
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      FATALERROR "Invalid File Number: " & Trim$(Str$(Value))
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      FATALERROR "File not open: " & Trim$(Str$(Value))
      GoTo ErrorFound
   End If
   ' --- check if locked record is in the current file ---
   If HasLockedRec And LockedFileNum = Value Then
      UNLOCKREC
   End If
   ' --- build sql query ---
   strSQL = ""
   With Files(Value)
      For Each oCadolXref In .CadolXrefs
         ' --- delete all matching records ---
         strSQL = strSQL & "DELETE FROM [" & oCadolXref.SQLTableName & "] "
         strWhere = "WHERE"
         If .Device <> 255 Then
            strSQL = strSQL & strWhere & " DEVICE = " & Trim$(Str$(.Device)) & " "
            strWhere = "AND"
         End If
         If .AdjVolume <> "" Then
            strSQL = strSQL & strWhere & " VOLUME = '" & .AdjVolume & "' "
            strWhere = "AND"
         End If
         If oCadolXref.Multiple Then
            strSQL = strSQL & strWhere & " FILENAME = '" & .FileName & "' "
            strWhere = "AND"
         End If
         ' --- if no records left, truncate table ---
         strSQL = strSQL & "IF (SELECT COUNT(*) FROM [" & oCadolXref.SQLTableName & "]) = 0 "
         strSQL = strSQL & "TRUNCATE TABLE [" & oCadolXref.SQLTableName & "] "
      Next
   End With
   ' --- execute the sql query ---
   If strSQL <> "" Then
      On Error GoTo ErrorFound
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      cnSQL.Execute strSQL, , adCmdText
      On Error GoTo 0
   End If
   ' --- done ---
   InitializeFile = True
   Exit Function
ConnError:
   ThrowError "InitializeFile", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   InitializeFile = False
End Function

Public Function FixPath(ByVal Value As String) As String
   Dim strResult As String
   ' ---------------------
   strResult = Value
   If strResult <> "" And Right$(strResult, 1) <> "\" Then
      strResult = strResult & "\"
   End If
   FixPath = strResult
End Function

Public Function UnpackFields(ByRef rsData As ADODB.Recordset, _
                             ByRef oCadolXref As rtCadolXref, _
                             ByRef StartPos As Long, _
                             ByRef DataLen As Long) As Boolean
   Dim CurrValue As Variant
   Dim CurrPos As Long
   Dim SavePos As Long
   Dim lngLoop As Long
   Dim lngYear As Long
   Dim lngMonth As Long
   Dim lngDay As Long
   Dim lngMaxDay As Long
   Dim bArray() As Byte
   Dim MaxLen As Long
   Dim HasPackedData As Boolean
   Dim oDataFormat As rtDataFormat
   ' -----------------------------
   HasPackedData = False
   On Error GoTo ErrorFound
   CurrPos = StartPos
   For Each oDataFormat In oCadolXref.DataFormats
      With oDataFormat
         Select Case .CadolType
            Case "N"
               CurrValue = GetNumeric(CurrPos, .CadolLength)
               If CurrValue <> 0 And .CadolScale <> 0 Then
                  If .CadolScale = 1 Then CurrValue = CurrValue / 10
                  If .CadolScale = 2 Then CurrValue = CurrValue / 100
                  If .CadolScale = 3 Then CurrValue = CurrValue / 1000
                  If .CadolScale = 4 Then CurrValue = CurrValue / 10000
               End If
               CurrPos = CurrPos + .CadolLength
               rsData.Fields(.FieldName) = CurrValue
            Case "A"
               CurrValue = GetAlpha(CurrPos)
               CurrPos = CurrPos + AlphaLen(CurrValue)
               If Len(CurrValue) > .CadolLength Then
                  ThrowError "UnpackFields", "Alpha value exceeds field length: " & _
                             .FieldName & " - """ & CurrValue & """"
                  GoTo ErrorFound
               End If
               rsData.Fields(.FieldName) = CurrValue
            Case "D", "DC"
               CurrValue = GetNumeric(CurrPos, .CadolLength)
               CurrPos = CurrPos + .CadolLength
               If CurrValue = 0 Then
                  CurrValue = Null
               Else
                  ' --- check for dates needing centuries ---
                  If CurrValue > 0 And CurrValue <= 1999999 Then
                     If .CadolType <> "DC" And CurrValue < 550000@ Then
                        CurrValue = CurrValue + 20000000@
                     Else
                        CurrValue = CurrValue + 19000000@
                     End If
                  End If
                  ' --- split value into year, month, day ---
                  lngYear = CurrValue \ 10000
                  lngMonth = (CurrValue \ 100) - (lngYear * 100)
                  lngDay = CurrValue - (lngYear * 10000) - (lngMonth * 100)
                  ' --- check for invalid values ---
                  If lngYear < 1900 Or lngYear > 9999 Then GoTo InvalidDate
                  If lngMonth < 1 Or lngMonth > 12 Then GoTo InvalidDate
                  If lngDay < 1 Or lngDay > 31 Then GoTo InvalidDate
                  ' --- check if day is after end of month ---
                  If lngDay > 28 Then
                     lngMaxDay = 31
                     If lngMonth = 4 Or lngMonth = 6 Or lngMonth = 9 Or lngMonth = 11 Then
                        lngMaxDay = 30
                     End If
                     If lngMonth = 2 Then
                        If lngDay > 29 Then GoTo InvalidDate
                        If (lngYear \ 400) * 400 = lngYear Then
                           lngMaxDay = 29
                        ElseIf (lngYear \ 100) * 100 = lngYear Then
                           lngMaxDay = 28
                        ElseIf (lngYear \ 4) * 4 = lngYear Then
                           lngMaxDay = 29
                        Else
                           lngMaxDay = 28
                        End If
                     End If
                     If lngDay > lngMaxDay Then GoTo InvalidDate
                  End If
                  ' --- build date string ---
                  CurrValue = CDate(Trim$(Str$(lngMonth)) & "/" & Trim$(Str$(lngDay)) & "/" & Trim$(Str$(lngYear)))
               End If
               rsData.Fields(.FieldName) = CurrValue
            Case "B"
               If UCase$(.FieldName) = "PACKED_DATA" Then
                  HasPackedData = True
               End If
               MaxLen = DataLen - (CurrPos - StartPos)
               If MaxLen <= 0 Then
                  CurrValue = Null
               Else
                  ReDim bArray(MaxLen - 1)
                  For lngLoop = 0 To MaxLen - 1
                     bArray(lngLoop) = MEM(CurrPos)
                     CurrPos = CurrPos + 1
                  Next lngLoop
                  CurrValue = bArray ' convert array to variant
               End If
               rsData.Fields(.FieldName) = CurrValue
            Case "U"
               CurrValue = ""
               MaxLen = .CadolLength - 1
               For lngLoop = 0 To MaxLen
                  CurrValue = CurrValue & Chr$(MEM(CurrPos + lngLoop) - 128)
               Next lngLoop
               CurrPos = CurrPos + .CadolLength
               rsData.Fields(.FieldName) = CurrValue
            Case "X"
               SavePos = CurrPos
               CurrValue = GetAlpha(CurrPos)
               CurrPos = SavePos + .CadolLength
               rsData.Fields(.FieldName) = CurrValue
            ' --- these types don't store actual data in SQL ---
            Case "C", "FN"
               CurrPos = CurrPos + .CadolLength ' move pointer forward without get
            Case "FA"
               CurrValue = GetAlpha(CurrPos) ' move pointer forward by getting alpha
               CurrPos = CurrPos + AlphaLen(CurrValue)
         End Select
      End With
   Next
   ' --- store packed data here without using SQL functions ---
   If Not HasPackedData Then
      CurrPos = StartPos
      ReDim bArray(DataLen - 1)
      For lngLoop = 0 To DataLen - 1
         bArray(lngLoop) = MEM(CurrPos)
         CurrPos = CurrPos + 1
      Next lngLoop
      CurrValue = bArray ' convert array to variant
      rsData.Fields("PACKED_DATA") = CurrValue
   End If
   ' --- done ---
Done:
   UnpackFields = True
   Exit Function
InvalidDate:
   ThrowError "UnpackFields", "Invalid Date: " & oDataFormat.FieldName & " - """ & CurrValue & """"
ErrorFound:
   UnpackFields = False
End Function

Public Function MatchKeyPattern(ByVal KeyValue As String, ByVal Pattern As String) As Boolean
   Dim lngLoop As Long
   ' --------------------
   ' --- check for null pattern first ---
   If Pattern = "" Then GoTo Matches
   ' --- check the length ---
   If Len(KeyValue) > Len(Pattern) Then GoTo NoMatch
   ' --- check char by char against format ---
   KeyValue = UCase$(KeyValue)
   Pattern = UCase$(Pattern)
   For lngLoop = 1 To Len(Pattern)
      Select Case Mid$(Pattern, lngLoop, 1)
         Case "^" ' required character
            If lngLoop > Len(KeyValue) Then GoTo NoMatch
         Case "?" ' optional char
            ' --- don't need to do anything here ---
         Case Else
            If Mid$(KeyValue, lngLoop, 1) <> Mid$(Pattern, lngLoop, 1) Then GoTo NoMatch
      End Select
   Next lngLoop
   ' --- done ---
Matches:
   MatchKeyPattern = True
   Exit Function
NoMatch:
   MatchKeyPattern = False
End Function

Public Function OpenAllFiles() As Boolean
   ' --- this builds Files() objects for all files marked as open in memory. ---
   ' --- it is ONLY used after loading a new runtime from a memory image.    ---
   Dim FileNum As Long
   Dim lngFileType As Long
   Dim lngSaveVol As Long
   Dim strSaveKey As String
   ' ----------------------
   lngSaveVol = VOL
   strSaveKey = KEY
   For FileNum = 0 To MaxFile
      ' --- check if file is open ---
      If MEMTF(MemPos_FileTable + (FileNum * FileEntrySize)) Then
         LET_VOL MEM(MemPos_FileTable + (FileNum * FileEntrySize) + 1)
         lngFileType = MEM(MemPos_FileTable + (FileNum * FileEntrySize) + 2)
         LET_KEY GetAlpha(MemPos_FileTable + (FileNum * FileEntrySize) + 3)
         ' --- open the file access object ---
         If lngFileType = 2 Then
            If Not OPENDATA(FileNum) Then GoTo ErrorFound
         ElseIf lngFileType = 5 Then
            If Not OPENDIRECTORY(FileNum) Then GoTo ErrorFound
         End If
         LET_VOL lngSaveVol ' restore volume number
      End If
   Next
   LET_VOL lngSaveVol ' restore volume number
   LET_KEY strSaveKey ' restore key value
   OpenAllFiles = True
   Exit Function
ErrorFound:
   OpenAllFiles = False
End Function

Public Sub MemorySort(ByVal L As Long, ByVal R As Long)
   ' --- this is an implementation of the QuickSort algorithm ---
   Dim i As Long
   Dim j As Long
   Dim X As String
   Dim Y As Long
   ' -------------
   i = L
   j = R
   X = SortTags(SortIndex((L + R) / 2))
   Do While (i <= j)
      Do While (SortTags(SortIndex(i)) < X And i < R)
         i = i + 1
         If i = R Then Exit Do
      Loop
      Do While (X < SortTags(SortIndex(j)) And j > L)
         j = j - 1
         If j = L Then Exit Do
      Loop
      If (i < j) Then
         Y = SortIndex(i)
         SortIndex(i) = SortIndex(j)
         SortIndex(j) = Y
      End If
      If (i <= j) Then
         i = i + 1
         j = j - 1
      End If
   Loop
   If (L < j) Then MemorySort L, j
   If (i < R) Then MemorySort i, R
End Sub

Public Sub ParseRuntimeCommand()

   ' --- This will parse commands sent to the Runtime from the outside ---
   
   Dim lngPos As Long
   Dim strLine As String
   Dim Tokens() As String
   Dim strTemp As String
   Dim strTarget As String
   Dim lngValue As Long
   Dim lngLastVal As Long
   Dim lngLoop As Long
   Dim strShellFlag As String
   Dim lngResult As Long
   Dim strCommand As String
   Dim strLibName As String
   Dim strParams As String
   Dim strMemFilename As String
   Dim blnAdded As Boolean
   Dim oStackEntry As rtStackEntry
   ' -----------------------------
   
   Do While InStr(PendingInput, vbCrLf) > 0
      
      lngPos = InStr(PendingInput, vbCrLf)
      strLine = Left$(PendingInput, lngPos - 1)
      PendingInput = Mid$(PendingInput, lngPos + 2)
      
      If strLine <> "" Then ' ignore null commands
         
         strLine = UnpackMsg(strLine)
         Tokens = Split(strLine, vbTab) ' tokens separated by tabs
         ' --- show incoming messages ---
         If DebugFlag And DebugFlagLevel > 0 Then
            If InStr(Replace(strLine, vbTab, " "), "APPLICATION NOP") = 0 Then
               DebugMessage "RCVD: " & Replace(strLine, vbTab, " ")
            End If
         End If
         
         Select Case UCase$(Tokens(0))
            
            Case "APPLICATION"
               If UBound(Tokens) < 1 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               Select Case UCase$(Tokens(1))
                  Case "NOP"
                     ' --- does nothing, used to keep connection alive ---
                  Case "START"
                     ' --- starting point at which runtime can function ---
                     ReadyToRun = True
                  Case "END"
                     MUSTEXIT = True
                     EXITING = True
                     DebugMessage "Exiting from 'Application End'"
                     SWITCHING = False ' allows runtime to exit
                     WAITTOEXIT = False ' allows runtime to exit
                     ' --- force an escape ---
                     LET_ESCVAL 0 ' perform normal escape
                     Exit Sub
                  Case "SWITCHREADY"
                     If Not SWITCHING Then
                        ThrowError "APPLICATION:SWITCHREADY", "*** Unrequested SwitchReady message ***"
                        Exit Sub
                     End If
                     ' --- create memory file ---
                     strMemFilename = GetTempFile(TempPath, "MEM")
                     DebugMessage "Saving to Memory File """ & strMemFilename & """"
                     SaveMemory strMemFilename
                     ' --- get target library info ---
                     Set oStackEntry = New rtStackEntry
                     oStackEntry.FromString SpawnTarget
                     ' --- check if target library exists ---
                     If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundNewLibrary
                     ' --- check for library in PROG_VOL next ---
                     If oStackEntry.VolName <> "/SYSVOL" And oStackEntry.VolName <> "PROG_VOL" Then
                        oStackEntry.DevNum = 0
                        oStackEntry.VolName = "PROG_VOL"
                        If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundNewLibrary
                     End If
                     ' --- finally check in /SYSVOL ---
                     If oStackEntry.VolName <> "/SYSVOL" Then
                        oStackEntry.DevNum = 0
                        oStackEntry.VolName = "/SYSVOL"
                        If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundNewLibrary
                     End If
                     ' --- library not found ---
                     ThrowError "System Error", "Library not found: " & SpawnTarget
                     Exit Sub
                     ' --- switch to another runtime executable ---
FoundNewLibrary:
                     strCommand = """" & LibraryPath & MakeLibExe(oStackEntry.ToString) & """"
                     strParams = "/START=" & oStackEntry.ToString
                     strParams = strParams & " /HOSTIP=" & HostIP
                     strParams = strParams & " /PORT=" & Trim$(Str$(PortVal))
                     strParams = strParams & " /ENV=" & EnvName
                     strParams = strParams & " /INI=""" & IniFilename & """"
                     strParams = strParams & " /MEM=""" & strMemFilename & """"
                     strCommand = strCommand & " " & strParams
                     ' --- disconnect existing socket ---
                     DebugMessage "*** wsToServer.Close in APPLICATION:SWITCHREADY"
                     rtFormMain.wsToServer.Close
                     DoEvents
                     ' --- show message in debug window ---
                     DebugMessage "SWITCH TO RUNTIME > " & Replace(strCommand, ".EXE""", ".EXE""" & vbCrLf & Space(21))
                     ' --- create a runtime executable ---
                     strShellFlag = "TRUE"
                     If DebugFlag Then
                        strShellFlag = UCase$(GetINIString(APP_NAME & EnvName, "ShellFlag", "True", IniFilename))
                        If strShellFlag = "ASK" Then
                           Clipboard.Clear
                           Clipboard.SetText strParams
                           lngResult = MsgBox(Replace(strCommand, " /", vbCrLf & "/"), vbYesNo, "Shell to this file?")
                           If lngResult = vbYes Then strShellFlag = "TRUE"
                        End If
                     End If
                     If strShellFlag = "TRUE" Then
                        On Error Resume Next
                        strLibName = MakeLibExe(oStackEntry.ToString)
                        lngResult = ShellExecute(0, "", strLibName, strParams, LibraryPath, 0)
                        On Error GoTo 0
                        If lngResult <= 32 Then
                           ThrowError "APPLICATION:SWITCHREADY", "Cannot start new library: " & SpawnTarget & _
                                      vbCrLf & "Result = " & Trim$(Str$(lngResult))
                           Exit Sub
                        End If
                     Else
                        DebugMessage "Please start this library manually..."
                     End If
                     ' --- done ---
                     SWITCHING = False
                     Exit Sub
                  Case "SERVEROK" ' --- bounceback SERVEROK is not valid, so end ---
                     MUSTEXIT = True
                     EXITING = True
                     DebugMessage "Exiting from 'Application ServerOK'"
                     SWITCHING = False ' allows runtime to exit
                     WAITTOEXIT = False ' allows runtime to exit
                     ' --- force an escape ---
                     LET_ESCVAL 0 ' perform normal escape
                     Exit Sub
                  Case Else
                     InvalidCommand Tokens
                     Exit Sub
               End Select
            
            Case "REQUEST"
               If UBound(Tokens) < 1 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               ' --- if last token is "SERVER" then send back to server, otherwise send to client ---
               strTarget = ""
               If UBound(Tokens) > 1 Then
                  If UCase$(Tokens(UBound(Tokens))) = "SERVER" Then
                     strTarget = "SERVER" & vbTab
                  End If
               End If
               Select Case UCase$(Tokens(1))
                  Case "USERNUM"
                     ' --- for a background job, UserNum is in the memory file ---
                     SendToServer strTarget & "REPLY" & vbTab & "USERNUM" & vbTab & Trim$(Str$(USER))
                  Case "ORIGNUM"
                     ' --- for a background job, OrigNum is in the memory file ---
                     SendToServer strTarget & "REPLY" & vbTab & "ORIGNUM" & vbTab & Trim$(Str$(ORIG))
                  Case "LOGINID"
                     ' --- for a background job, LoginID is in the memory file ---
                     SendToServer strTarget & "REPLY" & vbTab & "LOGINID" & vbTab & UCase$(LoginID)
                  Case "MACHINENAME"
                     ' --- for a background job, MachineName is in the memory file ---
                     SendToServer strTarget & "REPLY" & vbTab & "MACHINENAME" & vbTab & UCase$(MachineName)
                  Case "CURRPROG"
                     ' --- contains DEV:VOL:LIBRARY:PROG:JUMPPOINT ---
                     strTemp = Trim$(Str$(CurrDevNum)) & ":" & CurrVolName & ":" & _
                               CurrLibName & ":" & Trim$(Str$(PROG)) & ":" & Trim$(Str$(CurrJumpPoint))
                     SendToServer strTarget & "REPLY" & vbTab & "CURRPROG" & vbTab & strTemp
                  Case "CURRINFO"
                     ' --- contains USER:ORIG:LOGINID:MACHINENAME:DEV:VOL:LIBRARY:PROG:JUMPPOINT ---
                     strTemp = Trim$(Str$(USER)) & ":" & Trim$(Str$(ORIG)) & ":" & _
                               Replace(UCase$(LoginID), ":", ";") & ":" & _
                               Replace(UCase$(MachineName), ":", ";") & ":" & _
                               Trim$(Str$(CurrDevNum)) & ":" & Replace(CurrVolName, ":", ";") & ":" & _
                               Replace(CurrLibName, ":", ";") & ":" & Trim$(Str$(PROG)) & _
                               ":" & Trim$(Str$(CurrJumpPoint))
                     SendToServer strTarget & "REPLY" & vbTab & "CURRINFO" & vbTab & strTemp
                  Case Else
                     InvalidCommand Tokens
                     Exit Sub
               End Select
            
            Case "VALUE"
               If UBound(Tokens) <> 2 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               Select Case UCase$(Tokens(1))
                  Case "USERINFO"
                     LoginID = LCase$(Tokens(2))
                     MachineName = LCase$(Tokens(3))
                  Case "LOGINID", "LOGIN"
                     LoginID = LCase$(Tokens(2))
                  Case "MACHINENAME", "MACHINE"
                     MachineName = LCase$(Tokens(2))
                  Case Else
                     InvalidCommand Tokens
                     Exit Sub
               End Select
         
            Case "KBD"
               If UBound(Tokens) <> 1 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               ' --- process all chars of message ---
               strTemp = ""
               lngValue = 255
               For lngLoop = 1 To Len(Tokens(1)) - 1 Step 2
                  lngLastVal = lngValue
                  lngValue = Val("&H" & Mid$(Tokens(1), lngLoop, 2))
                  ' --- check if doing local edit ---
                  If MEM(MemPos_LocalEdit) = 1 Then
                     ' --- handle special chars, even if ignoring other keystrokes ---
                     If lngLastVal = 0 Then
                        ' --- handle special chars ---
                        HandleSpecialChar lngValue
                        ' --- must acknowledge and use any Local Edit changes ---
                        If lngValue = SpecChar_LocalEditOn Or lngValue = SpecChar_LocalEditOn Then
                           KBuff.Add 0
                           KBuff.Add lngValue
                           strTemp = strTemp & HexChar(0) & HexChar(lngValue)
                        End If
                     End If
                  Else
                     ' --- add chars to keyboard buffer ---
                     KBuff.Add lngValue ' normal char
                     strTemp = strTemp & HexChar(lngValue)
                  End If
               Next lngLoop
               ' --- send back list of chars received ---
               If strTemp <> "" Then
                  SendToServer "KEYBOARD" & vbTab & "RCVD" & vbTab & strTemp
               End If
            
            Case "CUSTOMWINDOW"
               If UBound(Tokens) < 1 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               Select Case UCase$(Tokens(1))
                  Case "WINRESULT"
                     If UBound(Tokens) <> 2 Then
                        InvalidCommand Tokens
                        Exit Sub
                     End If
                     CustomWindowResult = Tokens(2)
                  Case "WINPROCESSDONE"
                     CustomWindowProcessDone = True
                  Case Else
                     InvalidCommand Tokens
                     Exit Sub
               End Select
            
            Case "DEBUG"
               If UBound(Tokens) < 1 Then
                  InvalidCommand Tokens
                  Exit Sub
               End If
               ' --- note: the "If InBreakMode..." structure is needed for synchronization ---
               Select Case UCase$(Tokens(1))
                  Case "GO"
                     InBreakMode = False
                  Case "ONESTEP", "BREAK", "STEPINTO"
                     If InBreakMode Then
                        DebugOneStep = True
                        InBreakMode = False
                     Else
                        DebugOneStep = True
                     End If
                  Case "STEPOVER"
                     If InBreakMode Then
                        DebugStepOver = True
                        InBreakMode = False
                     Else
                        DebugStepOver = True
                     End If
                  Case "STEPOUT"
                     blnAdded = False
                     For lngLoop = GosubStack.Count To 1 Step -1
                        Set oStackEntry = GosubStack.Item(lngLoop)
                        With oStackEntry
                           If .ItemType = GOSUB_TYPEVAL Then
                              ' --- ok to add the breakpoint before turning off InBreakMode ---
                              If .DevNum = CurrDevNum And .VolName = CurrVolName And .LibName = CurrLibName Then
                                 AddBreakPoint GOSUB_TYPEVAL, .ProgNum, .JumpNum
                                 blnAdded = True
                              Else
                                 ' --- can't break across libraries (yet) ---
                                 AddBreakPoint GOSUB_TYPEVAL, 0, 0
                                 blnAdded = True
                              End If
                              Exit For
                           End If
                        End With
                     Next lngLoop
                     If Not blnAdded Then
                        ' --- nothing on the gosub stack ---
                        AddBreakPoint GOSUB_TYPEVAL, 0, 0
                     End If
                     If InBreakMode Then
                        InBreakMode = False
                     End If
                  Case "LOOP"
                     AddBreakPoint GOSUB_TYPEVAL, PROG, CurrJumpPoint
                     If InBreakMode Then
                        InBreakMode = False
                     End If
                  Case "BREAKPOINT"
                     If UBound(Tokens) < 2 Then
                        InvalidCommand Tokens
                        Exit Sub
                     End If
                     Set oStackEntry = New rtStackEntry
                     oStackEntry.FromString Tokens(2)
                     Breakpoints.Add oStackEntry
                     Set oStackEntry = Nothing
                     ' --- send debug data to client program ---
                     SendDebugData
                     ' --- BREAKPOINT does not turn off InBreakMode ---
                  Case Else
                     InvalidCommand Tokens
                     Exit Sub
               End Select
               
            Case Else
               InvalidCommand Tokens
               Exit Sub

         End Select

      End If
   
   Loop

End Sub

Public Sub InvalidCommand(ByRef Tokens() As String)
   Dim lngLoop As Long
   Dim strLine As String
   ' -------------------
   strLine = ""
   For lngLoop = LBound(Tokens) To UBound(Tokens)
      strLine = Trim$(strLine & " " & Tokens(lngLoop))
   Next lngLoop
   ThrowError "DataArrival", "Invalid Command: " & strLine
End Sub

Public Function InsideIDE() As Boolean
   Static Checked As Boolean
   Static Result As Boolean
   ' -----------------------
   If Checked Then GoTo Done
   ' --- check if running this program inside the VB6 IDE ---
   On Error GoTo AreInsideIDE
   Debug.Print 1 / 0 ' only throws an error inside the VB6 IDE
   Result = False
   GoTo DoneIDECheck
AreInsideIDE:
   Result = True
DoneIDECheck:
   On Error GoTo 0
   Checked = True
Done:
   InsideIDE = Result
End Function

Public Function GetUserNum() As Long
   Dim TempUserNum As Long
   Dim rsUser As ADODB.Recordset
   ' ---------------------------
TryAgain:
   On Error GoTo ErrorFound
   If cnSQL Is Nothing Then GoTo ConnError
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   Set rsUser = cnSQL.Execute("GETUSERNUM", , adCmdStoredProc)
   TempUserNum = rsUser.Fields(0)
   If TempUserNum < 1 Or TempUserNum >= 32767 Then GoTo ErrorFound
   On Error GoTo AlreadyInUse
   cnSQL.Execute "INSERT INTO [%USER] ([USERNUM]) VALUES (" & TempUserNum & ")", , adCmdText
   On Error GoTo 0
   GetUserNum = TempUserNum
   Exit Function
AlreadyInUse:
   cnSQL.Errors.Clear
   Resume TryAgain
ConnError:
   ThrowError "GetUserNum", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   On Error GoTo 0
   GetUserNum = -1 ' error
End Function

Public Function UpdateUserInfo() As Boolean
   Dim strSQL As String
   Dim lngNumChanged As Long
   ' -----------------------
   UpdateUserInfo = False
   strSQL = "IF NOT EXISTS (SELECT USERNUM FROM [%USER] " & _
            "WHERE USERNUM = " & Trim$(Str$(USER)) & ") " & _
            "INSERT INTO [%USER] ([USERNUM],[LOGINID]) " & _
            "VALUES (" & Trim$(Str$(USER)) & ",'" & UCase$(LoginID) & "') "
   On Error GoTo ErrorFound
   If cnSQL Is Nothing Then GoTo ConnError
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   cnSQL.Execute strSQL, lngNumChanged, adCmdText + adExecuteNoRecords
   On Error GoTo 0
   strSQL = "UPDATE [%USER] "
   strSQL = strSQL & "SET [ORIGNUM] = " & Trim$(Str$(ORIG)) & " , "
   strSQL = strSQL & "[LOGINID] = '" & UCase$(LoginID) & "' , "
   strSQL = strSQL & "[MACHINENAME] = '" & UCase$(MachineName) & "' , "
   strSQL = strSQL & "[DEVICE] = " & Trim$(Str$(CurrDevNum)) & " , "
   strSQL = strSQL & "[VOLUME] = '" & UCase$(CurrVolName) & "' , "
   strSQL = strSQL & "[LIBRARY] = '" & UCase$(CurrLibName) & "' "
   strSQL = strSQL & "WHERE [USERNUM] = " & Trim$(Str$(USER)) & " "
   strSQL = strSQL & "AND ([LOGINID] IS NULL OR [LOGINID] = '" & UCase$(LoginID) & "') "
   On Error GoTo ErrorFound
   If cnSQL Is Nothing Then GoTo ConnError
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   cnSQL.Execute strSQL, lngNumChanged, adCmdText + adExecuteNoRecords
   On Error GoTo 0
   If lngNumChanged = 1 Then
      UpdateUserInfo = True
   End If
   Exit Function
ConnError:
   ThrowError "UpdateUserInfo", "SQL Connection Error:"
   GoTo ErrorDone
ErrorFound:
   Resume ErrorDone
ErrorDone:
End Function
 
Public Function GetPRTNUM() As Long
   ' --- this will find the default printer number for the current login id ---
   Dim strSQL As String
   Dim TempPrtNum As Long
   Dim rsLogin As ADODB.Recordset
   ' ----------------------------
   On Error GoTo Done
   TempPrtNum = 0 ' default to slave printer
   Set rsLogin = New ADODB.Recordset
   strSQL = "SELECT PRTNUM FROM [%LOGIN] "
   strSQL = strSQL & "WHERE LOGINID = '" & LoginID & "' "
   With rsLogin
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
      If Not .EOF Then
         TempPrtNum = .Fields("PRTNUM").Value
      End If
      .Close
   End With
   Set rsLogin = Nothing
Done:
   On Error GoTo 0
   GetPRTNUM = TempPrtNum
   Exit Function
ConnError:
   ThrowError "GetPRTNUM", "SQL Connection Error:"
   GetPRTNUM = -1
End Function

Public Function GetGReg(ByVal RegName As String) As Currency
   Dim strSQL As String
   ' ------------------
   strSQL = "SELECT * FROM [%GREGS] "
   strSQL = strSQL & "WHERE [NAME] = '" & RegName & "' "
   With rsGRegs
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
      GetGReg = .Fields("VALUE")
      .Close
   End With
   Exit Function
ConnError:
   ThrowError "GetGReg", "SQL Connection Error:"
   GetGReg = 0
End Function

Public Sub LetGReg(ByVal RegName As String, ByVal Value As Currency)
   Dim strSQL As String
   ' ------------------
   strSQL = "SELECT * FROM [%GREGS] "
   strSQL = strSQL & "WHERE [NAME] = '" & RegName & "' "
   With rsGRegs
      ' --- must be adUseServer for concurrency issues ---
      .CursorLocation = adUseServer
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      ' --- shift value into 0-255 range ---
      If Value < 0 Or Value > 255 Then
         Value = Value - (Int(Value / 256) * 256)
      End If
      ' --- set field value ---
      .Fields("VALUE") = Value
      .Update
      .Close
   End With
   Exit Sub
ConnError:
   ThrowError "LetGReg", "SQL Connection Error:"
   Exit Sub
End Sub

Public Sub CheckDoEvents()
   ' --- make sure DoEvents fires sometimes ---
   Dim sngTimer As Single
   ' --------------------
   ' --- only check every 100 times ---
   CheckDoEventsCount = CheckDoEventsCount + 1
   If CheckDoEventsCount < 100 Then Exit Sub
   CheckDoEventsCount = 0
   ' --- only run DoEvents once per second ---
   sngTimer = Timer
   If sngTimer < Last_DoEvents Or sngTimer >= Last_DoEvents + 1 Then
      Last_DoEvents = sngTimer
      DoEvents
   End If
End Sub

Public Function MakeLibExe(ByVal Value As String) As String
   Dim Tokens() As String
   Dim strResult As String
   Dim strResult2 As String
   ' ----------------------
   Tokens = Split(Value, ":")
   strResult = "DEVICE" & Format(Val(Tokens(0)), "00") & "\" & _
               AdjustFilenameWindows(Tokens(1)) & "\" & _
               "LIB_" & AdjustFilenameWindows(Tokens(2)) & ".EXE"
   strResult2 = "DEVICE" & Format(Val(Tokens(0)), "00") & "\" & _
                AdjustFilenameWindows(Tokens(1)) & "\" & _
                "UPDATES\" & _
                "LIB_" & AdjustFilenameWindows(Tokens(2)) & ".EXE"
   ' --- check for latest library ---
   On Error GoTo Done
   If Dir$(LibraryPath & "\" & strResult) = "" Then GoTo Done
   If Dir$(LibraryPath & "\" & strResult2) = "" Then GoTo Done
   If FileDateTime(LibraryPath & "\" & strResult2) > FileDateTime(LibraryPath & "\" & strResult) Then
      strResult = strResult2
   End If
Done:
   On Error GoTo 0
   MakeLibExe = strResult
End Function

Public Sub DebugMessage(ByVal Msg As String)
   If Not DebugFlag Then Exit Sub
   If DebugLogFile > 0 Then
      Print #DebugLogFile, Msg
   Else
      rtDebugLog.AddMessage Msg
      Debug.Print Msg
   End If
End Sub

Public Function GetAppLock(ByVal Resource As String) As Boolean
   ' ----------------------------------------------------------
   ' --- The SQL 2000 Application Lock will prevent another
   ' --- Runtime from locking this record specified in the
   ' --- Resource string. In Cadol, the Resource specifies the
   ' --- "Dev:Vol:File:Rec#" which is to be locked.
   ' ----------------------------------------------------------
   ' --- If LOCKVAL is zero, this routine will wait forever
   ' --- (checking for EXITING every few seconds). If LOCKVAL
   ' --- is one, it will fail immediately, as per Cadol logic.
   ' ----------------------------------------------------------
   ' --- Only one Application Lock can be held at a time using
   ' --- this routine, due to its use of rsLockedResult.
   ' ----------------------------------------------------------
   ' --- Note: Don't close the rsLockedResult recordset until
   ' --- ReleaseAppLock (see notes in ReleaseAppLock for info).
   ' ----------------------------------------------------------
   Dim strSQL As String
   Dim lngWaitTime As Long
   ' ---------------------
   lngWaitTime = 0 ' immediate return to start
   If Resource = "" Then GoTo ErrorFound
   On Error GoTo ErrorFound
TryAgain:
   strSQL = "DECLARE @Result INT " & vbCrLf
   strSQL = strSQL & "EXEC @Result = sp_getapplock "
   strSQL = strSQL & "@Resource = '" & Resource & "' "
   strSQL = strSQL & ", @LockMode = 'Exclusive' "
   strSQL = strSQL & ", @LockOwner = 'Session' "
   If LOCKVAL = 0 Then
      strSQL = strSQL & ", @LockTimeout = '" & Trim$(Str$(lngWaitTime)) & "' "
   Else
      strSQL = strSQL & ", @LockTimeout = '0' " ' return immediately with error
   End If
   strSQL = strSQL & vbCrLf & "SELECT @Result "
   If cnSQL Is Nothing Then GoTo ConnError
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   Set rsLockedResult = cnSQL.Execute(strSQL, , adCmdText)
   If rsLockedResult.Fields(0).Value < 0 Then
      DoEvents
      If EXITING Then GoTo ErrorFound
      ' --- check for timeout ---
      If LOCKVAL = 0 And rsLockedResult.Fields(0).Value = -1 Then
         If lngWaitTime < 30000 Then
            lngWaitTime = lngWaitTime + 1000 ' add one second up to 30 seconds
         End If
         GoTo TryAgain
      End If
      GoTo ErrorFound
   End If
   DebugMessage "GETAPPLOCK: " & Resource
   GetAppLock = True
   Exit Function
ConnError:
   ThrowError "GetAppLock", "SQL Connection Error:"
   GoTo ErrorDone
ErrorFound:
   Resume ErrorDone ' clear error
ErrorDone:
   On Error Resume Next
   rsLockedResult.Close
   Set rsLockedResult = Nothing
   On Error GoTo 0
   GetAppLock = False
End Function

Public Function ReleaseAppLock(ByVal Resource As String) As Boolean
   ' -----------------------------------------------------------------
   ' --- The SQL 2000 Application Lock will sometimes be released when
   ' --- the rsLockedResult recordset is closed, and sometimes it has
   ' --- to be released explicitly by a call to "sp_releaseapplock".
   ' --- There doesn't seem to be any way to tell which method will
   ' --- work in any given situation.
   ' -----------------------------------------------------------------
   ' --- Note: 08/18/2009, ReleaseAppLock started getting stuck in an
   ' --- infinite loop. Added code to exit immediately when it knows
   ' --- it can't process properly.
   ' -----------------------------------------------------------------
   ' --- Don't close the rsLockedResult recordset anywhere other than
   ' --- here! Also, all errors get ignored, as either way could work.
   ' -----------------------------------------------------------------
   Dim strSQL As String
   ' ------------------
   ' --- Don't execute code if there are errors ---
   ReleaseAppLock = True
   If cnSQL Is Nothing Then Exit Function
   If cnSQL.Errors.Count > 0 Then Exit Function
   If rsLockedResult Is Nothing Then Exit Function
   ' --- Check for invalid parameters ---
   If Resource = "" Then GoTo ErrorFound
   ' --- try to release by closing recordset ---
   On Error Resume Next
   rsLockedResult.Close
   Set rsLockedResult = Nothing
   ' --- try to release using sp_releaseapplock ---
   On Error GoTo ErrorFound
   strSQL = "DECLARE @Result INT " & vbCrLf
   strSQL = strSQL & "EXEC @Result = sp_releaseapplock "
   strSQL = strSQL & "@Resource = '" & Resource & "' "
   strSQL = strSQL & ", @LockOwner = 'Session' "
   strSQL = strSQL & vbCrLf & "SELECT @Result "
   Set rsLockedResult = cnSQL.Execute(strSQL, , adCmdText)
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   GoTo Done
ConnError:
   ThrowError "ReleaseAppLock", "SQL Connection Error:"
   Exit Function
ErrorFound:
   Resume Done ' clear error
Done:
   ' --- ignore all errors, assume everything worked ---
   DebugMessage "RELEASEAPPLOCK: " & Resource
   cnSQL.Errors.Clear
   If Not (rsLockedResult Is Nothing) Then
      On Error Resume Next
      rsLockedResult.Close
      Set rsLockedResult = Nothing
   End If
   On Error GoTo 0
   ReleaseAppLock = True
End Function

Private Function WildcardMatch(ByVal MatchTo As String, ByVal Value As String) As Boolean
   ' --- @ - can match any character
   ' --- # - can match any number
   Dim lngLoop As Long
   Dim strTemp1 As String
   Dim strTemp2 As String
   ' --------------------
   ' --- check lengths ---
   If Len(MatchTo) <> Len(Value) Then GoTo NoMatch
   ' --- see if wildcards match ---
   For lngLoop = 1 To Len(MatchTo)
      strTemp1 = Mid$(MatchTo, lngLoop, 1)
      strTemp2 = Mid$(Value, lngLoop, 1)
      If strTemp1 = strTemp2 Then GoTo NextChar
      If strTemp1 = "#" Then
         If strTemp2 >= "0" And strTemp2 <= "9" Then GoTo NextChar
      End If
      If strTemp1 = "@" Then GoTo NextChar
      GoTo NoMatch
NextChar:
   Next lngLoop
   ' --- done ---
   WildcardMatch = True
   Exit Function
NoMatch:
   WildcardMatch = False
End Function

Public Function GetLibNameFromApp() As String
   Dim lngTemp As Long
   Dim strTemp As String
   Dim strTemp2 As String
   Dim strTemp3 As String
   Dim strStartProg As String
   ' ------------------------
   strStartProg = ""
   ' --- build library name from path ---
   strTemp = UCase$(App.Path)
   If InStr(strTemp, "\DEVICE") > 0 Then
      strTemp = Mid$(strTemp, InStr(strTemp, "\DEVICE") + 7)
      lngTemp = InStr(strTemp, "\")
      If lngTemp > 0 Then
         strTemp2 = Trim$(Str$(Val(Left$(strTemp, lngTemp - 1)))) ' device
         strTemp = Mid$(strTemp, lngTemp + 1) ' volume
         If InStr(strTemp, "\") = 0 Then ' not other stuff
            strTemp = Replace(strTemp, "_", "/") ' fix volume name
            strTemp3 = UCase$(App.EXEName)
            If Left$(strTemp3, 4) = "LIB_" Then strTemp3 = Mid$(strTemp3, 5)
            strTemp3 = Replace(strTemp3, "_", "/")
            strStartProg = strTemp2 & ":" & strTemp & ":" & strTemp3
         End If
      End If
   End If
   GetLibNameFromApp = strStartProg
End Function

Public Sub AssignBackgroundPrintFile()
   PrinterFileName = GetTempFile(TempPath, "BGP")
   Kill PrinterFileName
   PrinterFileName = Left$(PrinterFileName, Len(PrinterFileName) - 4) & ".txt"
   PrinterType = "f" ' file
   PrinterParameters = ""
   PrinterFileNum = FreeFile
   Open PrinterFileName For Append As #PrinterFileNum
   LET_MEMTF MemPos_FFPending, False
   LET_MEMTF MemPos_PageHasData, False
   LET_MEMTF MemPos_LineHasData, False
   MEM(MemPos_PrintDev) = 127 ' printer handled programmatically
   LET_MEMTF MemPos_PrintOn, True
   MEM(MemPos_Status) = 0 ' ok
End Sub

Public Function GetKeyboardChar() As Byte
   Dim bChar As Byte
   Dim bLastChar As Byte
   ' -------------------
   bChar = 255
   bLastChar = 255
   Do
      If EXITING Then GoTo DoExit
      ' --- wait until a character is available ---
      Do While KBuff.Count = 0
         Sleep 1
         DoEvents
         If EXITING Then GoTo DoExit
         If DebugOneStep Then
            bChar = 255 ' will never happen normally
            GoTo Done
         End If
      Loop
      ' --- save last character ---
      bLastChar = bChar
      ' --- get next char ---
      bChar = KBuff.GetChar
      ' --- acknowledge characters used ---
      SendToServer "KEYBOARD" & vbTab & "USED" & vbTab & HexChar(bChar)
      ' --- handle special characters ---
      If bLastChar = 0 Then
         HandleSpecialChar bChar
      End If
   Loop Until bChar <> 0 And bLastChar <> 0
   ' --- return character ---
Done:
   GetKeyboardChar = bChar
   Exit Function
DoExit:
   GetKeyboardChar = 27 ' return escapes when exiting
End Function

Public Sub HandleSpecialChar(ByVal bChar As Byte)
   ' --- Should only be called when previous character was a Zero (=0). ---
   ' --- Also has to be called when using the chars, not when chars     ---
   ' --- are first received. Otherwise things will happen out of order. ---
   Select Case bChar
      ' --- local edit handling ---
      Case SpecChar_LocalEditOn
         MEM(MemPos_LocalEdit) = 2 ' have switched into local edit mode on client's end
      Case SpecChar_LocalEditOff
         MEM(MemPos_LocalEdit) = 0 ' done with local edit mode
      ' --- run scripts handling ---
      Case SpecChar_ScriptRunOn
         LET_MEMTF MemPos_ScriptRunFlag, True ' running a keyboard script
      Case SpecChar_ScriptRunOff
         LET_MEMTF MemPos_ScriptRunFlag, False ' done with keyboard script
         ' --- handshaking to finish script processing ---
         SendToServer "KEYBOARD" & vbTab & "SCRIPT" & vbTab & "DONE"
      ' --- write scripts handling ---
      Case SpecChar_ScriptWriteOn
         LET_MEMTF MemPos_ScriptWriteFlag, True ' writing a keyboard script
      Case SpecChar_ScriptWriteOff
         LET_MEMTF MemPos_ScriptWriteFlag, False ' done with keyboard script
   End Select
End Sub

Public Sub CustomWindowProcessing()
   Do While (Not CustomWindowProcessDone) And (Not EXITING)
      Sleep 10 ' prevents too much cpu usage
      DoEvents
   Loop
End Sub

Public Sub AddBreakPoint(ByVal ItemType As Long, ByVal ProgNum As Long, ByVal JumpNum As Long)
   Dim objItem As rtStackEntry
   ' -------------------------
   ' --- add item to gosub stack ---
   Set objItem = New rtStackEntry
   With objItem
      .ItemType = ItemType
      ' --- check for a call to /USERLIB - it only exists in 0:/SYSVOL ---
      If CurrLibName = "/USERLIB" Then
         .DevNum = 0
         .VolName = "/SYSVOL"
      Else
         .DevNum = CurrDevNum
         .VolName = CurrVolName
      End If
      .LibName = CurrLibName
      .ProgNum = ProgNum
      .JumpNum = JumpNum
   End With
   Breakpoints.Add objItem
End Sub

Public Function UserAuthorized() As Boolean
   Dim rsCheck As ADODB.Recordset
   ' ----------------------------
   On Error GoTo ErrorFound
   UserAuthorized = False
   Set rsCheck = cnSQL.Execute("select * from [%EOMDates] " & _
                               "where getdate() >= StartDate " & _
                               "and getdate() < EndDate")
   If rsCheck.BOF And rsCheck.EOF Then
      ' --- not during an EOM period ---
      UserAuthorized = True
   Else
      Set rsCheck = cnSQL.Execute("select * from [%EOMUsers] " & _
                                  "where LoginID = '" & LCase$(LoginID) & "' " & _
                                  "and ((BeginDate is null) or (getdate() >= BeginDate)) " & _
                                  "and ((TermDate is null) or (getdate() < TermDate))")
      If (Not rsCheck.BOF) Or (Not rsCheck.EOF) Then
         ' --- user is authorized during EOM ---
         UserAuthorized = True
      End If
   End If
   Set rsCheck = Nothing
   On Error GoTo 0
   Exit Function
ErrorFound:
   On Error GoTo 0
   UserAuthorized = True
End Function

Public Sub PrepareClientLists()
   Dim LoopNum As Integer
   Dim Clients() As String
   ' ---------------------
   ' --- remove any single quote characters ---
   ClientList = Replace(ClientList, "'", "")
   ' --- check for a comma-separated list of clients ---
   If InStr(ClientList, ",") > 0 Then
      Clients = Split(ClientList, ",")
      ShortClientCompare = " IN ('" & Right$("0000" & Trim$(Clients(LBound(Clients, 1))), 3)
      LongClientCompare = " IN ('" & Right$("0000" & Trim$(Clients(LBound(Clients, 1))), 4)
      For LoopNum = LBound(Clients, 1) + 1 To UBound(Clients, 1)
         ShortClientCompare = ShortClientCompare & "','" & Right$("0000" & Trim$(Clients(LoopNum)), 3)
         LongClientCompare = LongClientCompare & "','" & Right$("0000" & Trim$(Clients(LoopNum)), 4)
      Next LoopNum
      ShortClientCompare = ShortClientCompare & "') "
      LongClientCompare = LongClientCompare & "') "
   Else ' --- single client ---
      ShortClientCompare = " = '" & Right$("0000" & Trim$(ClientList), 3) & "' "
      LongClientCompare = " = '" & Right$("0000" & Trim$(ClientList), 4) & "' "
   End If
   DebugMessage "ShortClientCompare = """ & ShortClientCompare & """"
   DebugMessage "LongClientCompare = """ & LongClientCompare & """"
End Sub

Public Function GetClientWhere(ByVal SQLTableName As String) As String
   Dim Result As String
   Dim strSQL As String
   Dim rsClientScript As ADODB.Recordset
   ' -----------------------------------
   ' --- default value returns no records ---
   Result = " REC < 0 "
   ' --- find proper script for reading tables for specified clients ---
   On Error GoTo ErrorFound
   Set rsClientScript = New ADODB.Recordset
   strSQL = "SELECT * FROM [%CLIENT_TABLE_SCRIPT] "
   strSQL = strSQL & "WHERE SQLTableName = '" & FixSqlStr(SQLTableName) & "' "
   With rsClientScript
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
      If Not .EOF Then
         Result = rsClientScript.Fields("SQLScript")
         If Result <> "" Then
            Result = Replace(Result, "$S$", ShortClientCompare)
            Result = Replace(Result, "$L$", LongClientCompare)
            Result = " " & Trim$(Result) & " " ' need leading and trailing spaces, just in case
         End If
      End If
   End With
   ' --- done with recordset ---
   rsClientScript.Close
   Set rsClientScript = Nothing
   On Error GoTo 0
   ' --- done ---
   DebugMessage "GetClientWhere = """ & Result & """"
   GetClientWhere = Result
   Exit Function
ConnError:
   ThrowError "GetClientWhere", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   GetClientWhere = Result
End Function
