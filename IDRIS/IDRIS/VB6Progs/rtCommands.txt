Attribute VB_Name = "rtCommands"
' -------------------------------
' --- rtCommands - 02/09/2010 ---
' -------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 02/09/2010 - SBakker - URD 11076
'            - Changed writing records to be done with adUseServer. Adding the
'              handling of History files throws errors using adUseClient.
'            - Make sure .CursorLocation is always the first property set.
' 06/18/2009 - SBakker - URD 11076
'            - Update "ChangedBy" field with the current LoginID when writing
'              records. This will be used in single records as a memo, but later
'              used in history records to track who made changes.
' 10/13/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
'            - Made changes recommended by CodeAdvisor.
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
' 12/06/2007 - SBAKKER - URD 11049
'            - Added error information to RELEASEDEVICE indicating which routine
'              causes an error to be thrown and what the error number was.
' 01/22/2007 - SBAKKER - URD 10883
'            - Added new functions ENTERBYTE and EDITBYTE. Will be used when the
'              target numeric register is a byte register, and will prevent any
'              numeric overflows from "ENTER (3) TERM" or such. The IDRIS
'              compilers will be changed to know when to use these functions.
' 01/22/2007 - SBAKKER - URD 9739
'            - Turn off rtFormMain.SQLTimer before running command with EXECSQL.
'              It may run for longer than the normal timeout period, and thus
'              would cause a timeout error before it could finish.
'            - Added error checking on the cnSQL object before using it, to
'              prevent hanging runtime processes.
' 12/12/2006 - Removed unused local variables.
' 05/23/2006 - Added new numeric function GETCOMPANYNAMELEN.
' 05/03/2006 - Added new alpha functions GETCOMPANYNAME and GETCOMPANYINITIALS.
' 02/08/2006 - Split out SendDebugData into separate routine so that it can be
'              called directly from the "DEBUG | BREAKPOINT" command.
' 02/06/2006 - Allow DEBUG BREAK signal to interrupt an EDITALPHA command (and
'              all associated variants). This means that sending the DEBUG BREAK
'              signal will get an immediate response, instead of waiting for the
'              user to press ENTER/ESC/CAN.
' 02/03/2006 - Update CurrJumpPoint from each ILCode line in DBUG.
' 01/30/2006 - Removed "DebugMessage '*** SWITCHING...'" messages. The problem
'              these were tracking has been corrected and is not needed anymore.
'            - Changed specific values to use MaxFile or MaxDevice constants.
' 01/27/2006 - Added in new command "DBUG". It is sent the IL code for the line
'              as a string (for display purposes), and will be used to enable
'              stepping, breaking, and other debugging tasks.
' 01/18/2006 - Use AltCommonFilePath if CommonFilePath doesn't exist. Otherwise
'              try to MKDIR CommonFilePath. Only throw an error if everything
'              that is tried fails.
' 12/30/2005 - Added BREAK command.
' 11/23/2005 - Set REC=0 at the start of READFILE, for proper REC value when an
'              error is encountered. So far only found necessary in CLM97STD.
' 11/22/2005 - Make sure that UpdateRecField gets cleared properly in WRITEFILE.
' 10/28/2005 - VOL=255 allows the file to be opened across all volumes. Added
'              support for parts of SQL WHERE clause not existing (such as
'              DEVICE and VOLUME).
' 10/07/2005 - Added extra checking for missing CommonFilePath before use.
' ------------------------------------------------------------------------------

' ------------------
' --- Assignment ---
' ------------------

Public Sub ZERO()
   N = 0
   N1 = 0
   N2 = 0
   N3 = 0
   N4 = 0
   N5 = 0
   N6 = 0
   N7 = 0
   N8 = 0
   N9 = 0
   N10 = 0
   N11 = 0
   N12 = 0
   N13 = 0
   N14 = 0
   N15 = 0
   N16 = 0
   N17 = 0
   N18 = 0
   N19 = 0
   N20 = 0
End Sub

' ---------------------------
' --- Transfer of control ---
' ---------------------------

Public Sub GOS(ByVal JumpNum As Long, ByVal ReturnPoint As Long)
   ' --- code must do an EXIT SUB after this command ---
   AddGosubStack GOSUB_TYPEVAL, PROG, ReturnPoint
   AddGosubStack GOSUB_TYPEVAL, PROG, JumpNum
   If DebugStepOver Then
      AddBreakPoint GOSUB_TYPEVAL, PROG, ReturnPoint
      DebugStepOver = False
   End If
End Sub

Public Sub GOSUBPROG(ByVal ProgNum As Long, ByVal ReturnPoint As Long)
   ' --- code must do an EXIT SUB after this command ---
   AddGosubStack GOSUB_TYPEVAL, PROG, ReturnPoint
   AddGosubStack GOSUB_TYPEVAL, ProgNum, 0
   If DebugStepOver Then
      AddBreakPoint GOSUB_TYPEVAL, PROG, ReturnPoint
      DebugStepOver = False
   End If
End Sub

Public Sub LOADPROG(ByVal ProgNum As Long)
   ' --- code must do an EXIT SUB after this command ---
   AddGosubStack GOSUB_TYPEVAL, ProgNum, 0
End Sub

Public Sub RETURNPROG()
   ' --- code must do an EXIT SUB after this command ---
   ' --- doesn't need to do anything except EXIT SUB ---
End Sub

Public Sub Cancel()
   MUSTEXIT = True
   Select Case CANVAL
      Case 0
         AddGosubStack GOSUB_TYPEVAL, PROG, 0
         MEM(MemPos_Status) = 1
      Case 1
         MEM(MemPos_Status) = 1
         MUSTEXIT = False
      Case 2
         BELL 1
         MUSTEXIT = False
      Case 3
         REJECT
         MUSTEXIT = False
      Case 4
         If Not HasGosubStackItem(WHEN_CANCEL_TYPEVAL) Then
            SYSERROR "NO CANCEL VECTOR" ' trappable
            Exit Sub
         End If
         ClearGosubStackTo WHEN_CANCEL_TYPEVAL
         LET_CANVAL 0
      Case Else
         FATALERROR "UNKNOWN CANCEL VALUE: " & Trim$(Str$(CANVAL))
   End Select
End Sub

Public Sub ESC()
   MUSTEXIT = True
   If MEMTF(MemPos_Background) Then
      LET_ESCVAL 0 ' must escape to clean up background process
   End If
   Select Case ESCVAL
      Case 0
         ExecuteEscape
      Case 1
         MEM(MemPos_Status) = 2
         MUSTEXIT = False
      Case 2
         BELL 1
         MUSTEXIT = False
      Case 3
         REJECT
         MUSTEXIT = False
      Case 4
         If Not HasGosubStackItem(WHEN_ESCAPE_TYPEVAL) Then
            SYSERROR "NO ESCAPE VECTOR" ' trappable
            Exit Sub
         End If
         ClearGosubStackTo WHEN_ESCAPE_TYPEVAL
         LET_ESCVAL 0
      Case Else
         FATALERROR "UNKNOWN ESCAPE VALUE: " & Trim$(Str$(ESCVAL))
   End Select
End Sub

Public Sub WHENCANCEL(ByVal ProgNum As Long, ByVal JumpNum As Long)
   AddGosubStack WHEN_CANCEL_TYPEVAL, ProgNum, JumpNum
   LET_CANVAL 4
End Sub

Public Sub WHENESCAPE(ByVal ProgNum As Long, ByVal JumpNum As Long)
   AddGosubStack WHEN_ESCAPE_TYPEVAL, ProgNum, JumpNum
   LET_ESCVAL 4
End Sub

Public Sub WHENERROR(ByVal ProgNum As Long, ByVal JumpNum As Long)
   ' --- if JumpNum = -1, this will clear the error handler ---
   AddGosubStack WHEN_ERROR_TYPEVAL, ProgNum, JumpNum
End Sub

' -------------------------------------
' --- Device allocation and control ---
' -------------------------------------

Public Function ASSIGNDEVICE(ByVal Value As Long) As Boolean
   Dim strSQL As String
   Dim lngResult As Long
   Dim rsPrinters As ADODB.Recordset
   ' -------------------------------
   If MEM(MemPos_PrintDev) <> 255 Then
      MEM(MemPos_Status) = 1 ' already assigned
      GoTo ErrorFound
   End If
   ' --- adjust device number if 1,2,3 by adding PRTNUM-1 ---
   If Value >= 1 And Value <= 3 Then
      If PRTNUM = 0 Then ' default printer is slave
         Value = 0 ' slave printer
      Else
         Value = Value + PRTNUM - 1
      End If
   End If
   ' --- check for slave printer ---
   If Value = 0 Then
      ' --- can't use a slave printer in the background ---
      If MEMTF(MemPos_Background) Then GoTo ErrorFound
      PrinterType = "s" ' slave
      PrinterDeviceName = ""
      PrinterParameters = ""
      GoTo HavePrinterInfo
   End If
   ' --- check for invalid devices ---
   If Value < 1 Or Value > 99 Then
      MEM(MemPos_Status) = 0 ' device not configured
      GoTo ErrorFound
   End If
   ' --- make sure device is configured ---
   On Error GoTo ErrorFound
   Set rsPrinters = New ADODB.Recordset
   strSQL = "SELECT * FROM [%PRINTERS] "
   strSQL = strSQL & "WHERE PRINTERNUM = " & Trim$(Str$(Value)) & " "
   strSQL = strSQL & "ORDER BY PRINTERNUM "
   With rsPrinters
      ' --- this is static data. adUseClient is fine. ---
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .ActiveConnection = Nothing
      If .BOF And .EOF Then
         .Close
         MEM(MemPos_Status) = 0 ' device not configured
         GoTo ErrorFound
      End If
   End With
   ' --- save printer information ---
   PrinterType = rsPrinters.Fields("PrinterType")
   PrinterDeviceName = rsPrinters.Fields("DeviceName")
   PrinterParameters = rsPrinters.Fields("Parameters")
   ' --- done with recordset ---
   rsPrinters.Close
   Set rsPrinters = Nothing
   On Error GoTo 0
   ' --- have printer info ---
HavePrinterInfo:
   ' --- create temporary print file ---
   If Value = 0 Then ' Printer 0
      ' --- check if common file path doesn't exist ---
      On Error Resume Next
      If Dir$(CommonFilePath, vbDirectory) = "" Then
         MkDir CommonFilePath
         Err.Clear
         On Error Resume Next
      End If
      If Dir$(CommonFilePath, vbDirectory) = "" And AltCommonFilePath <> "" Then
         CommonFilePath = AltCommonFilePath
         If Dir$(CommonFilePath, vbDirectory) = "" Then
            MkDir CommonFilePath
            Err.Clear
            On Error Resume Next
         End If
      End If
      If Dir$(CommonFilePath, vbDirectory) = "" Then
         ThrowError "AssignDevice", "CommonFilePath not found: " & CommonFilePath
         GoTo ErrorFound
      End If
      On Error GoTo ErrorFound
      ' --- slave printing ---
      PrinterFileName = GetTempFile(CommonFilePath, "PRT")
      Kill PrinterFileName
      PrinterFileName = Left$(PrinterFileName, Len(PrinterFileName) - 4) & ".txt"
      If Not MEMTF(MemPos_Background) Then
         SendToServer "APPLICATION" & vbTab & "STATUSLINE" & vbTab & "Printing to file " & PrinterFileName
      End If
   ElseIf LCase$(PrinterType) = "f" Then
      ' --- build printer filename ---
      PrinterFileName = FileServerPath
      If Left$(PrinterDeviceName, 1) = "\" Then
         PrinterFileName = PrinterFileName & Mid$(PrinterDeviceName, 2) ' remove leading "\"
      Else
         PrinterFileName = PrinterFileName & PrinterDeviceName
      End If
      PrinterFileName = Replace(PrinterFileName, "*", FormatNum("z6", Int(Rnd * 1000000)))
   Else
      ' --- check if printer is accessible ---
      If LCase$(PrinterType) = "d" Then
         lngResult = OpenPrinter(PrinterDeviceName, PrinterHandle, 0)
         If lngResult = 0 Then GoTo ErrorFound
         lngResult = ClosePrinter(PrinterHandle)
      End If
      ' --- build printer filename ---
      PrinterFileName = GetTempFile(TempPath, "PRT")
   End If
   ' --- create output file ---
   PrinterFileNum = FreeFile
   Open PrinterFileName For Output As #PrinterFileNum
   ' --- check for control characters needed ---
   If LCase$(PrinterType) = "d" Then
      Select Case LCase$(PrinterParameters)
         Case "-c"
            Print #PrinterFileNum, Chr$(27) & "&l2a0O" & Chr$(27) & "(s0p16.67h8.5v0T";
         Case "-8"
            Print #PrinterFileNum, Chr$(27) & "&l2a0o8D" & Chr$(27) & "(s0p16.67h0T";
         Case "-l"
            Print #PrinterFileNum, Chr$(27) & "&l2a1o5.45C" & Chr$(27) & "(s0p12h3T";
         Case "-w"
            Print #PrinterFileNum, Chr$(27) & "&l2a1o5.45C" & Chr$(27) & "(s0p16.67h3T";
         Case "-o"
            Print #PrinterFileNum, Chr$(27) & "&k2S";
      End Select
   End If
   ' --- not waiting for a formfeed to print ---
   LET_MEMTF MemPos_FFPending, False
   LET_MEMTF MemPos_PageHasData, False
   LET_MEMTF MemPos_LineHasData, False
   MEM(MemPos_PrintDev) = Value
   MEM(MemPos_Status) = 0 ' ok
   ASSIGNDEVICE = True
   Exit Function
ConnError:
   ThrowError "ASSIGNDEVICE", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   Set rsPrinters = Nothing
   ASSIGNDEVICE = False
End Function

Public Sub RELEASEDEVICE()
   Dim lngPID As Long
   Dim lngResult As Long
   Dim lngWritten As Long
   Dim strTemp As String
   Dim MyDocInfo As DOCINFO
   Dim strCommand As String
   Dim strError As String
   ' ----------------------
   strError = "Unknown Error"
   If MEM(MemPos_PrintDev) = 255 Then Exit Sub ' not assigned
   ' --- check if missing filename ---
   If PrinterFileName = "" Then GoTo Done
   ' --- finish last line ---
   If MEMTF(MemPos_LineHasData) Then
      Print #PrinterFileNum,
      LET_MEMTF MemPos_LineHasData, False
   End If
   ' --- don't print final formfeed ---
   LET_MEMTF MemPos_FFPending, False
   LET_MEMTF MemPos_PageHasData, False
   ' --- close printer file ---
   Close #PrinterFileNum
   DoEvents
   PrinterFileNum = 0
   ' --- let environment settle down some ---
   DoEvents
   ' --- check if print file is empty ---
   If FileLen(PrinterFileName) = 0 Then
      On Error Resume Next
      Kill PrinterFileName
      On Error GoTo 0
      GoTo Done
   End If
   ' --- handle slave printer ---
   If MEM(MemPos_PrintDev) = 0 Then
      If Not MEMTF(MemPos_Background) Then
         SendToServer "PRINTER" & vbTab & "SLAVE" & vbTab & PrinterFileName
         SendToServer "KEYBOARD" & vbTab & "LOCK" & vbTab & "OFF"
         SendToServer "APPLICATION" & vbTab & "STATUSLINE" & vbTab & ""
      End If
      GoTo Done
   End If
   ' --- send to printer and delete when done ---
   If LCase$(PrinterType) = "d" Then
      ' --- re-open printer file for input ---
      If Dir$(PrinterFileName) = "" Then
         ThrowError "RELEASEDEVICE", "File not found: " & PrinterFileName
         Exit Sub
      End If
      PrinterFileNum = FreeFile
      Open PrinterFileName For Input As #PrinterFileNum
      ' --- build DOCINFO record ---
      strTemp = Right$(PrinterFileName, Len(PrinterFileName) - InStrRev(PrinterFileName, "\"))
      strTemp = Left$(strTemp, InStr(strTemp, ".") - 1) ' remove extension
      MyDocInfo.pDocName = UCase$(LoginID) & "_" & UCase$(MachineName) & "_" & strTemp
      MyDocInfo.pOutputFile = vbNullString
      MyDocInfo.pDatatype = vbNullString
      ' --- open handle to printer for spooling ---
      DoEvents
      strError = "Error Adding Printer Connection"
      lngResult = AddPrinterConnection(PrinterDeviceName)
      If lngResult = 0 Then GoTo ErrorFound
      DoEvents
      strError = "Error Opening Printer"
      lngResult = OpenPrinter(PrinterDeviceName, PrinterHandle, 0)
      If lngResult = 0 Then GoTo ErrorFound
      DoEvents
      strError = "Error Starting Document Printer"
      PrinterJobNum = StartDocPrinter(PrinterHandle, 1, MyDocInfo)
      If PrinterJobNum = 0 Then GoTo ErrorFound
      DoEvents
      strError = "Error Starting Page Print"
      lngResult = StartPagePrinter(PrinterHandle)
      If lngResult = 0 Then GoTo ErrorFound
      ' --- write out all lines from print file ---
      Do While Not EOF(PrinterFileNum)
         Line Input #PrinterFileNum, strTemp
         strTemp = strTemp & vbCrLf
         lngResult = WritePrinter(PrinterHandle, ByVal strTemp, Len(strTemp), lngWritten)
      Loop
      ' --- add trailing formfeed (for Printer 6 especially) ---
      strTemp = Chr$(12) ' formfeed
      lngResult = WritePrinter(PrinterHandle, ByVal strTemp, Len(strTemp), lngWritten)
      ' --- close all files and handles ---
      Close #PrinterFileNum
      DoEvents
      strError = "Error Ending Page Printer"
      lngResult = EndPagePrinter(PrinterHandle)
      If lngResult = 0 Then GoTo ErrorFound
      DoEvents
      strError = "Error Ending Doc Printer"
      lngResult = EndDocPrinter(PrinterHandle)
      If lngResult = 0 Then GoTo ErrorFound
      DoEvents
      strError = "Error Closing Printer"
      lngResult = ClosePrinter(PrinterHandle)
      If lngResult = 0 Then GoTo ErrorFound
      ' --- delete print file ---
      If Not InsideIDE Then
         On Error Resume Next
         Kill PrinterFileName
         On Error GoTo 0
      End If
      GoTo Done
   End If
   If PrinterParameters <> "" Then
      ' --- replace * with actual filename ---
      strCommand = Replace(PrinterParameters, "*", PrinterFileName)
      DebugMessage "CMD: " & strCommand
      lngPID = Shell(strCommand, vbHide)
      ' --- don't wait for terminate, don't delete file ---
      GoTo Done
   End If
Done:
   MEM(MemPos_PrintDev) = 255 ' not assigned
   LET_MEMTF MemPos_PrintOn, False ' implied PRINTOFF
   PrinterFileName = ""
   ' --- done with background when release device ---
   If MEMTF(MemPos_Background) Then
      ' --- can't just call "ExecuteEscape", as it calls "ReleaseDevice" again ---
      MEM(MemPos_EscVal) = 0 ' force normal escape
      UNLOCKREC
      CloseAllChannels
      CloseSortFile
      ClearGosubStack
      LET_MEMTF MemPos_TBAlloc, False
      MEM(MemPos_EscVal) = 0
      MEM(MemPos_CanVal) = 0
      MEM(MemPos_Status) = 0
      EXITRUNTIME
   End If
   Exit Sub
ErrorFound:
   ThrowError "ReleaseDevice", "*** Unable to print file """ & PrinterFileName & _
                               """ to printer """ & PrinterDeviceName & """ ***" & vbCrLf & _
                               "*** " & strError & " - Error# = " & Trim$(Str$(Err.LastDllError)) & " ***"
End Sub

Public Sub RELEASETERMINAL(ByVal JUMPPOINT As Long, ByVal ErrorPoint As Long)
   Dim strTemp As String
   Dim strCommand As String
   Dim strLibName As String
   Dim strParams As String
   Dim strShellFlag As String
   Dim lngResult As Long
   Dim strMemFilename As String
   Dim lngSaveUser As Long
   Dim oStackEntry As rtStackEntry
   ' -----------------------------
   On Error GoTo ErrorFound
   ' --- cannot carry locked record across spawn ---
   UNLOCKREC
   ' --- don't actually release terminal if using a slave printer ---
   If MEM(MemPos_PrintDev) = 0 Then ' slave printing
      MEM(MemPos_Status) = 0 ' indicates current partition has printer
      Exit Sub
   End If
   ' --- assign default background print file if needed ---
   If MEM(MemPos_PrintDev) = 255 Then ' not printing
      AssignBackgroundPrintFile
   End If
   ' --- set status for background partition ---
   MEM(MemPos_Status) = 0
   ' --- must get a new user number for background partition ---
   lngSaveUser = USER
   LetNumeric MemPos_User, 2, GetUserNum
   ' --- store memory into temp file ---
   strMemFilename = GetTempFile(TempPath, "MEM")
   DebugMessage "Saving to Memory File """ & strMemFilename & """"
   SaveMemory strMemFilename
   ' --- once the memory file is saved, must clean up everything ---
   LetNumeric MemPos_User, 2, lngSaveUser
   ' --- close printer if assigned, but don't use ReleaseDevice ---
   LET_MEMTF MemPos_PrintOn, False
   If MEM(MemPos_PrintDev) <> 255 Then ' assigned
      If PrinterFileNum <> 0 Then
         Close #PrinterFileNum
         DoEvents
         PrinterFileNum = 0
      End If
      MEM(MemPos_PrintDev) = 255 ' not assigned
      PrinterFileName = ""
   End If
   ' --- close open sort file, but don't use CloseSortFile ---
   If MEM(MemPos_SortState) <> 0 And SortFileNum <> 0 Then
      Close #SortFileNum
      DoEvents
      SortFileName = ""
      SortFileNum = 0
      SortTagSize = 0
      SortLineCount = 0
      FetchLineCount = 0
      MEM(MemPos_SortState) = 0 ' not sorting
   End If
   ' --- clean up all resources before spawning ---
   If ESCVAL = 3 Then
      ' --- close all channels if transferred to background ---
      If TCHAN = 0 Then
         CloseAllChannels
      End If
      ' --- add line number of next command ---
      AddGosubStack GOSUB_TYPEVAL, PROG, JUMPPOINT
   Else
      ' --- execute an escape in this partition ---
      ExecuteEscape
   End If
   ' --- let environment settle down some ---
   DoEvents
   ' --- get target library info ---
   Set oStackEntry = New rtStackEntry
   oStackEntry.DevNum = CurrDevNum
   oStackEntry.VolName = CurrVolName
   oStackEntry.LibName = CurrLibName
   oStackEntry.ProgNum = PROG
   oStackEntry.JumpNum = JUMPPOINT
   ' --- check if target library exists ---
   If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundSpawnLibrary
   ' --- check for library in PROG_VOL next ---
   If oStackEntry.VolName <> "/SYSVOL" And oStackEntry.VolName <> "PROG_VOL" Then
      oStackEntry.DevNum = 0
      oStackEntry.VolName = "PROG_VOL"
      If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundSpawnLibrary
   End If
   ' --- finally check in /SYSVOL ---
   If oStackEntry.VolName <> "/SYSVOL" Then
      oStackEntry.DevNum = 0
      oStackEntry.VolName = "/SYSVOL"
      If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundSpawnLibrary
   End If
   ' --- change /SYSLIB to /USERLIB. /SYSLIB doesn't exist ---
   If oStackEntry.LibName = "/SYSLIB" Then
      oStackEntry.LibName = "/USERLIB"
      If Dir$(LibraryPath & MakeLibExe(oStackEntry.ToString)) <> "" Then GoTo FoundSpawnLibrary
   End If
   ' --- library not found ---
   ThrowError "ReleaseTerminal", "Library not found: " & _
              Trim$(Str$(CurrDevNum)) & ":" & CurrVolName & ":" & CurrLibName
   Exit Sub
   ' --- switch to another runtime executable ---
FoundSpawnLibrary:
   strCommand = """" & LibraryPath & MakeLibExe(oStackEntry.ToString) & """"
   strParams = "/BG"
   strParams = strParams & " /START=" & oStackEntry.ToString
   strParams = strParams & " /HOSTIP=" & BackgroundIP
   strParams = strParams & " /PORT=" & Trim$(Str$(BackgroundPort))
   strParams = strParams & " /ENV=" & EnvName
   strParams = strParams & " /INI=""" & IniFilename & """"
   strParams = strParams & " /MEM=""" & strMemFilename & """"
   strCommand = strCommand & " " & strParams
   ' --- show message in debug window ---
   DebugMessage "SPAWN RUNTIME > " & Replace(strCommand, ".EXE""", ".EXE""" & vbCrLf & Space(21))
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
      strLibName = MakeLibExe(oStackEntry.ToString)
      lngResult = ShellExecute(0, "", strLibName, strParams, LibraryPath, 0)
      If lngResult <= 32 Then
         ThrowError "RELEASETERMINAL", "Cannot start new library: " & strLibName & _
                    vbCrLf & "Result = " & Trim$(Str$(lngResult))
         Exit Sub
      End If
   Else
      DebugMessage "Please start this library manually..."
   End If
   ' --- continue execution if Esc=3 ---
   If ESCVAL = 3 Then
      ' --- set status for terminal partition ---
      MEM(MemPos_Status) = 1
   End If
   On Error GoTo 0
   Exit Sub
ErrorFound:
   On Error GoTo 0
   AddGosubStack GOSUB_TYPEVAL, PROG, ErrorPoint
End Sub

Public Sub PRINTON()
   If MEM(MemPos_PrintDev) = 255 Then ' no device assigned
      If Not ASSIGNDEVICE(0) Then
         ThrowError "PRINTON", "Cannot assign temporary file for slave printing"
         Exit Sub
      End If
   End If
   ' --- can't send messages if background partition ---
   If Not MEMTF(MemPos_Background) Then
      If MEM(MemPos_PrintDev) = 0 Then ' slave printing
         SendToServer "KEYBOARD" & vbTab & "LOCK" & vbTab & "ON"
      End If
   End If
   LET_MEMTF MemPos_PrintOn, True
End Sub

Public Sub PRINTOFF()
   ' --- can't turn off printing for background partition ---
   If Not MEMTF(MemPos_Background) Then
      If MEM(MemPos_PrintDev) = 0 Then ' slave printing
         SendToServer "KEYBOARD" & vbTab & "LOCK" & vbTab & "OFF"
      End If
      LET_MEMTF MemPos_PrintOn, False
   End If
End Sub

' ----------------------
' --- Terminal Input ---
' ----------------------

Public Function ENTERNUM(ByVal NumericFmt As String) As Boolean
   Dim blnOK As Boolean
   Dim lngLen As Long
   Dim strResult As String
   Dim curResult As Currency
   ' -----------------------
   ' --- can't enter during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "ENTERNUM", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
      MEM(MemPos_Char) = 255
   End If
   ' --- format value for editing ---
   strResult = ""
   lngLen = LenFormat(NumericFmt)
   ' --- loop until valid numeric ---
   Do
      If strResult <> "" Then
         SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(Len(strResult))
         SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(Len(strResult)))
         strResult = Trim$(strResult)
      End If
      If Not EDITALPHA(lngLen, strResult) Then GoTo HadEvent
      FREEZE_LENGTH = False ' unfreeze length for numeric edits
      strResult = ALPHA_RESULT
      blnOK = ConvertNum(NumericFmt, strResult, curResult)
      If Not blnOK Then
         REJECT
         strResult = ""
      End If
   Loop Until blnOK
   ' --- redisplay result ---
   If strResult = "" Then
      LET_CHAR 0 ' no chars entered
   Else
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(CHAR))
      strResult = FormatNum(NumericFmt, curResult)
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & strResult
      LET_CHAR Len(strResult) ' number redisplayed
   End If
   ' --- return result ---
   NUMERIC_RESULT = curResult
   ENTERNUM = True
   Exit Function
HadEvent:
   ENTERNUM = False
End Function

Public Function ENTERBYTE(ByVal NumericFmt As String) As Boolean
   Dim blnOK As Boolean
   Dim lngLen As Long
   Dim strResult As String
   Dim curResult As Currency
   ' -----------------------
   ' --- can't enter during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "ENTERBYTE", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
      MEM(MemPos_Char) = 255
   End If
   ' --- format value for editing ---
   strResult = ""
   lngLen = LenFormat(NumericFmt)
   ' --- loop until valid numeric ---
   Do
      If strResult <> "" Then
         SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(Len(strResult))
         SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(Len(strResult)))
         strResult = Trim$(strResult)
      End If
      If Not EDITALPHA(lngLen, strResult) Then GoTo HadEvent
      FREEZE_LENGTH = False ' unfreeze length for numeric edits
      strResult = ALPHA_RESULT
      blnOK = ConvertNum(NumericFmt, strResult, curResult)
      ' --- must be a valid byte value ---
      If blnOK Then
         If curResult < 0 Or curResult > 255 Then blnOK = False
      End If
      If Not blnOK Then
         REJECT
         strResult = ""
      End If
   Loop Until blnOK
   ' --- redisplay result ---
   If strResult = "" Then
      LET_CHAR 0 ' no chars entered
   Else
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(CHAR))
      strResult = FormatNum(NumericFmt, curResult)
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & strResult
      LET_CHAR Len(strResult) ' number redisplayed
   End If
   ' --- return result ---
   NUMERIC_RESULT = curResult
   ENTERBYTE = True
   Exit Function
HadEvent:
   ENTERBYTE = False
End Function

Public Function ENTERALPHA(ByVal Size As Long) As Boolean
   ' --- can't enter during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "ENTERALPHA", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check the size specified ---
   If Size < 0 Or Size > 256 Then
      ThrowError "ENTERALPHA", "Invalid size: " & Trim$(Str$(Size))
      Exit Function
   End If
   ' --- use EDITALPHA to do all the work ---
   ENTERALPHA = EDITALPHA(Size, "")
   Exit Function
HadEvent:
   ENTERALPHA = False
End Function

Public Function EDITNUM(ByVal NumericFmt As String, ByVal Value As Currency) As Boolean
   Dim blnOK As Boolean
   Dim lngLen As Long
   Dim strResult As String
   Dim curResult As Currency
   ' -----------------------
   ' --- can't edit during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "EDITNUM", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
      MEM(MemPos_Char) = 255
   End If
   ' --- format value for editing ---
   strResult = FormatNum(NumericFmt, Value)
   lngLen = LenFormat(NumericFmt)
   ' --- loop until valid numeric ---
   Do
      If strResult <> "" Then
         SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(Len(strResult))
         SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(Len(strResult)))
         strResult = Trim$(strResult)
      End If
      ' --- use EDITALPHA to do all the work ---
      If Not EDITALPHA(lngLen, strResult) Then GoTo HadEvent
      FREEZE_LENGTH = False ' unfreeze length for numeric edits
      strResult = ALPHA_RESULT
      blnOK = ConvertNum(NumericFmt, strResult, curResult)
      If Not blnOK Then
         REJECT
         strResult = ""
      End If
   Loop Until blnOK
   ' --- redisplay result ---
   If strResult = "" Then
      LET_CHAR 0 ' no chars entered
   Else
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(CHAR))
      strResult = FormatNum(NumericFmt, curResult)
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & strResult
      LET_CHAR Len(strResult) ' number redisplayed
   End If
   ' --- return result ---
   NUMERIC_RESULT = curResult
   EDITNUM = True
   Exit Function
HadEvent:
   EDITNUM = False
End Function

Public Function EDITBYTE(ByVal NumericFmt As String, ByVal Value As Currency) As Boolean
   Dim blnOK As Boolean
   Dim lngLen As Long
   Dim strResult As String
   Dim curResult As Currency
   ' -----------------------
   ' --- can't edit during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "EDITBYTE", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
      MEM(MemPos_Char) = 255
   End If
   ' --- format value for editing ---
   strResult = FormatNum(NumericFmt, Value)
   lngLen = LenFormat(NumericFmt)
   ' --- loop until valid numeric ---
   Do
      If strResult <> "" Then
         SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(Len(strResult))
         SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(Len(strResult)))
         strResult = Trim$(strResult)
      End If
      ' --- use EDITALPHA to do all the work ---
      If Not EDITALPHA(lngLen, strResult) Then GoTo HadEvent
      FREEZE_LENGTH = False ' unfreeze length for numeric edits
      strResult = ALPHA_RESULT
      blnOK = ConvertNum(NumericFmt, strResult, curResult)
      ' --- must be a valid byte value ---
      If blnOK Then
         If curResult < 0 Or curResult > 255 Then blnOK = False
      End If
      If Not blnOK Then
         REJECT
         strResult = ""
      End If
   Loop Until blnOK
   ' --- redisplay result ---
   If strResult = "" Then
      LET_CHAR 0 ' no chars entered
   Else
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(CHAR))
      strResult = FormatNum(NumericFmt, curResult)
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & strResult
      LET_CHAR Len(strResult) ' number redisplayed
   End If
   ' --- return result ---
   NUMERIC_RESULT = curResult
   EDITBYTE = True
   Exit Function
HadEvent:
   EDITBYTE = False
End Function

Public Function EDITALPHA(ByVal Size As Long, ByVal Value As String) As Boolean
   Dim bChar As Byte
   Dim strTemp As String
   Dim strResult As String
   Dim blnDoingLocalEdit As Boolean
   Dim blnSaveDebugOneStep As Boolean
   ' ------------------------------
   blnDoingLocalEdit = False
   ' --- save DebugOneStep for interrupt checking and restoring ---
   blnSaveDebugOneStep = DebugOneStep
   DebugOneStep = False
   ' --- can't edit during background processing ---
   If MEMTF(MemPos_Background) Then
      ThrowError "EDITALPHA", "Enter/Edit during Background Processing"
      GoTo HadEvent
   End If
   ' --- check the size specified ---
   If Size < 0 Or Size > 256 Then
      ThrowError "EDITALPHA", "Invalid size: " & Trim$(Str$(Size))
      GoTo HadEvent
   End If
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
      MEM(MemPos_Char) = 255
   End If
   ' --- get starting value ---
   strResult = Value ' starting value
   LET_CHAR Len(strResult)
   ' --- show value on screen ---
   If strResult <> "" Then
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & strResult
   End If
NextChar:
   If EXITING Then GoTo HadEvent
   ' --- if keyboard buffer is empty, switch to local edit ---
   If AllowLocalEdit And (Not blnDoingLocalEdit) Then
      ' --- give messages a chance to settle ---
      If KBC = 0 Then
         Sleep 1
         DoEvents
         If EXITING Then GoTo HadEvent
      End If
      ' --- only start local edit if no chars in keyboard buffer ---
      If KBC = 0 Then
         ' --- turn on local edit flag. all further chars are discarded ---
         ' --- until KBD SpecChar_LocalEditOn/Off message is received.  ---
         MEM(MemPos_LocalEdit) = 1 ' starting local edit process
         ' --- set internal flag to prevent re-display ---
         blnDoingLocalEdit = True
         ' --- send local edit message ---
         SendToServer "KEYBOARD" & vbTab & "EDIT" & vbTab & _
                      Trim$(Str$(Size)) & vbTab & HexString(strResult)
      End If
   End If
   ' --- get next char ---
   bChar = GetKeyboardChar
   If EXITING Then GoTo HadEvent
   ' --- check if a DEBUG BREAK was sent ---
   If DebugOneStep Then
      DebugOneStep = False
      BREAK
      blnSaveDebugOneStep = DebugOneStep
      DebugOneStep = False
      If bChar = 0 Or bChar = 255 Then GoTo NextChar
   End If
   ' --- process character ---
   Select Case bChar
      Case 32 To 126 ' normal character
         If Len(strResult) < Size Then
            If Not blnDoingLocalEdit Then ' not in local edit process
               SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Chr$(bChar)
            End If
            strResult = strResult & Chr$(bChar)
            LET_CHAR Len(strResult)
         End If
      Case 8 ' backspace
         If Len(strResult) > 0 Then
            If Not blnDoingLocalEdit Then ' not in local edit process
               SendToServer "CURSOR" & vbTab & "LEFT"
               SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & " "
               SendToServer "CURSOR" & vbTab & "LEFT"
            End If
            strResult = Left$(strResult, Len(strResult) - 1)
            LET_CHAR Len(strResult)
         End If
      Case 127 ' delete
         If strResult <> "" Then
            strTemp = Trim$(Str$(Len(strResult)))
            If Not blnDoingLocalEdit Then ' not in local edit process
               SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & strTemp
               SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(Len(strResult))
               SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & strTemp
            End If
         End If
         strResult = ""
         LET_CHAR 0
      Case 13 ' enter
         MEM(MemPos_Status) = 0
         GoTo Done
      Case 24 ' cancel
         Select Case CANVAL
            Case 0
               AddGosubStack GOSUB_TYPEVAL, PROG, 0
               MEM(MemPos_Status) = 1
               GoTo HadEvent
            Case 1
               MEM(MemPos_Status) = 1
               GoTo Done
            Case 2
               BELL 1
            Case 3
               REJECT
               strResult = ""
            Case 4
               If Not HasGosubStackItem(WHEN_CANCEL_TYPEVAL) Then
                  SYSERROR "NO CAN VECTOR" ' trappable
                  GoTo HadEvent
               End If
               ClearGosubStackTo WHEN_CANCEL_TYPEVAL
               LET_CANVAL 0
               GoTo HadEvent
            Case Else
               FATALERROR "UNKNOWN CANCEL VALUE: " & Trim$(Str$(CANVAL))
               GoTo HadEvent
         End Select
      Case 27 ' escape
         Select Case ESCVAL
            Case 0
               ExecuteEscape
               GoTo HadEvent
            Case 1
               MEM(MemPos_Status) = 2
               GoTo Done
            Case 2
               BELL 1
            Case 3
               REJECT
               strResult = ""
            Case 4
               If Not HasGosubStackItem(WHEN_ESCAPE_TYPEVAL) Then
                  SYSERROR "NO ESC VECTOR" ' trappable
                  GoTo HadEvent
               End If
               ClearGosubStackTo WHEN_ESCAPE_TYPEVAL
               LET_ESCVAL 0
               GoTo HadEvent
            Case Else
               FATALERROR "UNKNOWN ESCAPE VALUE: " & Trim$(Str$(ESCVAL))
               GoTo HadEvent
         End Select
   End Select
   GoTo NextChar
Done:
   ' --- update Length ---
   LET_LENGTH Len(strResult)  ' null string sets length to 0
   ' --- freeze the value of LENGTH for one string assignment. ---
   ' --- this allows a null string to have a length of zero.   ---
   FREEZE_LENGTH = True
   ' --- return result ---
   ALPHA_RESULT = strResult
   ' --- done ---
   EDITALPHA = True
   DebugOneStep = blnSaveDebugOneStep
   Exit Function
HadEvent:
   ' --- done ---
   FREEZE_LENGTH = False
   EDITALPHA = False
   DebugOneStep = blnSaveDebugOneStep
End Function

' ---------------------
' --- Device Output ---
' ---------------------

Public Sub DISPLAYSTRING(ByVal Value As String)
   ' --- this handles literal strings with no other processing ---
   If Value <> "" Then
      ' --- handle display ---
      If MEMTF(MemPos_PrintOn) Then
         ' --- print to printer ---
         CheckFormFeed
         Print #PrinterFileNum, Value;
         LET_MEMTF MemPos_PageHasData, True
         LET_MEMTF MemPos_LineHasData, True
      Else
         ' --- display to screen ---
         SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Value
      End If
   End If
   ' --- save number of chars displayed ---
   MEM(MemPos_Char) = Len(Value)
End Sub

Public Sub DISPLAY(ByVal Value As String)
   ' --- check if TAB needed ---
   If MEM(MemPos_Char) <> 255 Then
      If Not MEMTF(MemPos_PrintOn) Then
         SendToServer "CURSOR" & vbTab & "TAB"
      End If
   End If
   DISPLAYSTRING Value
   MEM(MemPos_Status) = 0 ' ok
End Sub

Public Sub DISPLAYNUM(ByVal NumericFmt As String, ByVal Value As Currency)
   ' --- get formatted numeric string ---
   DISPLAY FormatNum(NumericFmt, Value)
End Sub

Public Sub DCH(ByVal Value As String, ByVal NumTimes As Long)
   If NumTimes <= 0 Then Exit Sub
   ' --- handle display ---
   If MEMTF(MemPos_PrintOn) Then
      ' --- print to printer ---
      CheckFormFeed
      Print #PrinterFileNum, String$(NumTimes, Value);
      LET_MEMTF MemPos_PageHasData, True
      LET_MEMTF MemPos_LineHasData, True
   Else
      ' --- display to screen ---
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & String$(NumTimes, Value)
   End If
End Sub

Public Sub DCH_HEX(ByVal Value As String, ByVal NumTimes As Long)
   Dim bChar As Byte
   Dim lngLoop As Long
   Dim strResult As String
   ' ---------------------
   If NumTimes <= 0 Then Exit Sub
   ' --- get character specified ---
   bChar = Val("&H" & Value)
   ' --- build string to be sent ---
   strResult = ""
   For lngLoop = 1 To NumTimes
      If bChar < 32 Then
         strResult = strResult & Chr$(bChar) ' char only
      ElseIf bChar >= 32 And bChar < 128 Then
         strResult = strResult & Chr$(27) & Chr$(bChar) ' ESC + char
      Else
         strResult = strResult & Chr$(bChar - 128) ' char only
      End If
   Next lngLoop
   ' --- handle display ---
   If MEMTF(MemPos_PrintOn) Then
      ' --- print directly to printer file ---
      CheckFormFeed
      Print #PrinterFileNum, strResult;
      LET_MEMTF MemPos_PageHasData, True
      LET_MEMTF MemPos_LineHasData, True
   Else
      ' --- display to screen ---
      SendToServer "SCREEN" & vbTab & "DCHHEX" & vbTab & HexString(strResult)
   End If
End Sub

' ---------------------------------------
' --- Input/Output Control Statements ---
' ---------------------------------------

Public Sub STAY()
   MEM(MemPos_Char) = 255 ' stay
End Sub

Public Sub BACK()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   If MEM(MemPos_Char) > 0 And MEM(MemPos_Char) <> 255 Then
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(MEM(MemPos_Char)))
   End If
   MEM(MemPos_Char) = 255 ' implicit stay
End Sub

Public Sub REJECT()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   If MEM(MemPos_Char) > 0 And MEM(MemPos_Char) <> 255 Then
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(MEM(MemPos_Char)))
      SendToServer "SCREEN" & vbTab & "DISPLAY" & vbTab & Space(MEM(MemPos_Char))
      SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(MEM(MemPos_Char)))
   End If
   SendToServer "SOUND" & vbTab & "BELL"
   MEM(MemPos_Char) = 255 ' implicit stay
End Sub

Public Sub PAD(ByVal Value As Long)
   If Value - MEM(MemPos_Char) > 0 Then
      DISPLAYSTRING Space(Value - MEM(MemPos_Char))
   End If
   MEM(MemPos_Char) = 255 ' implicit stay
End Sub

' -----------------------------
' --- Open/Close Statements ---
' -----------------------------

Public Function OPENDEVICE(ByVal Value As Long) As Boolean
   Dim strPath As String
   ' -------------------
   If Value < 0 Or Value > MaxDevice Then
      MEM(MemPos_Status) = 1 ' invalid device number
      GoTo ErrorFound
   End If
   If MEMTF(MemPos_DevTable + (Value * DevEntrySize)) Then
      MEM(MemPos_Status) = 2 ' device already open
      GoTo ErrorFound
   End If
   ' --- see if device already exists ---
   strPath = FileServerPath & "DEVICE" & FormatNum("z2", Value)
   If Dir$(strPath, vbDirectory) <> "" Then GoTo CanOpenDevice
   ' --- build the device - all devices are valid ---
   On Error GoTo DeviceError
   MkDir strPath ' build the device directory on file server
   On Error GoTo 0
CanOpenDevice:
   LET_MEMTF MemPos_DevTable + (Value * DevEntrySize), True
   OPENDEVICE = True
   Exit Function
DeviceError:
   On Error GoTo 0
   MEM(MemPos_Status) = 3 ' device not configured
ErrorFound:
   OPENDEVICE = False
End Function

Public Sub CLOSEDEVICE(ByVal Value As Long)
   Dim lngLoop As Long
   ' -----------------
   If Value < 0 Or Value > MaxDevice Then
      MEM(MemPos_Status) = 1 ' invalid device number
      Exit Sub
   End If
   If Not MEMTF(MemPos_DevTable + (Value * DevEntrySize)) Then
      MEM(MemPos_Status) = 0 ' device already closed, no error
      Exit Sub
   End If
   ' --- check if any volumes are using this device ---
   For lngLoop = 0 To MaxVolume
      If MEMTF(MemPos_VolTable + (lngLoop * VolEntrySize)) Then ' volume is open
         If MEM(MemPos_VolTable + (lngLoop * VolEntrySize) + 1) = Value Then
            MEM(MemPos_Status) = 10 ' resource in use
            Exit Sub
         End If
      End If
   Next
   ' --- close the device ---
   LET_MEMTF MemPos_DevTable + (Value * DevEntrySize), False
End Sub

Public Function OPENVOLUME() As Boolean
   Dim lngLoop As Long
   Dim strPath As String
   ' -------------------
   ' --- check for bad parameters ---
   If (VOL < 0) Or (VOL > MaxDevice And VOL <> 255) Or _
      (KEY = "") Or (Len(KEY) > 8) Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check all devices ---
   If VOL = 255 Then
      For lngLoop = 0 To MaxDevice
         If MEMTF(MemPos_DevTable + (lngLoop * DevEntrySize)) Then ' device is open
            strPath = FileServerPath & "Device" & FormatNum("z2", lngLoop) & "\" & AdjustFilenameWindows(KEY)
            If Dir$(strPath, vbDirectory) <> "" Then
               MEM(MemPos_Vol) = lngLoop ' device number
               GoTo CanOpenVolume
            End If
         End If
      Next
      MEM(MemPos_Vol) = 0 ' build volume on device 0
   End If
   ' --- check specified device ---
   If Not MEMTF(MemPos_DevTable + (VOL * DevEntrySize)) Then
      MEM(MemPos_Status) = 3 ' device not open
      GoTo ErrorFound
   End If
   strPath = FileServerPath & "Device" & FormatNum("z2", VOL) & "\" & AdjustFilenameWindows(KEY)
   If Dir$(strPath, vbDirectory) <> "" Then GoTo CanOpenVolume
   ' --- build the volume - all volumes are valid ---
   On Error GoTo VolumeError
   MkDir strPath ' build the volume directory on file server
   On Error GoTo 0
   GoTo CanOpenVolume
VolumeError:
   ' --- volume not found ---
   On Error GoTo 0
   MEM(MemPos_Status) = 2 ' volume not found on specified device
   GoTo ErrorFound
CanOpenVolume:
   ' --- see if volume is already open ---
   For lngLoop = 0 To MaxVolume
      If MEMTF(MemPos_VolTable + (lngLoop * VolEntrySize)) Then ' volume is open
         If MEM(MemPos_VolTable + (lngLoop * VolEntrySize) + 1) = VOL Then ' same device
            If GetAlpha(MemPos_VolTable + (lngLoop * VolEntrySize) + 2) = KEY Then ' same name
               MEM(MemPos_Status) = 0 ' ok
               MEM(MemPos_Vol) = lngLoop ' volume number
               OPENVOLUME = True
               Exit Function
            End If
         End If
      End If
   Next
   ' --- find next empty slot ---
   For lngLoop = 0 To MaxVolume
      If Not MEMTF(MemPos_VolTable + (lngLoop * VolEntrySize)) Then ' volume is closed
         LET_MEMTF MemPos_VolTable + (lngLoop * VolEntrySize), True
         MEM(MemPos_VolTable + (lngLoop * VolEntrySize) + 1) = VOL
         LetAlpha MemPos_VolTable + (lngLoop * VolEntrySize) + 2, KEY
         MEM(MemPos_Status) = 0 ' ok
         MEM(MemPos_Vol) = lngLoop ' volume number
         OPENVOLUME = True
         Exit Function
      End If
   Next
   ' --- volume table full ---
   MEM(MemPos_Status) = 1 ' volume table full
ErrorFound:
   On Error GoTo 0
   MEM(MemPos_Vol) = 255 ' invalid volume
   OPENVOLUME = False
End Function

Public Sub CLOSEVOLUME()
   Dim lngLoop As Long
   ' -----------------
   ' --- check for bad parameters ---
   If VOL < 0 Or VOL > MaxVolume Then
      MEM(MemPos_Status) = 63 ' other error
      MEM(MemPos_Vol) = 255
      Exit Sub
   End If
   ' --- check if volume already closed ---
   If Not MEMTF(MemPos_VolTable + (VOL * VolEntrySize)) Then
      MEM(MemPos_Status) = 0 ' ok
      MEM(MemPos_Vol) = 255
      Exit Sub
   End If
   ' --- check if volume in use ---
   For lngLoop = 0 To MaxFile
      If MEMTF(MemPos_FileTable + (lngLoop * FileEntrySize)) Then
         If MEM(MemPos_FileTable + (lngLoop * FileEntrySize) + 1) = VOL Then
            MEM(MemPos_Status) = 10 ' resource in use
            Exit Sub
         End If
      End If
   Next
   For lngLoop = 0 To MaxTFA
      If MEMTF(MemPos_TFATable + (lngLoop * TFAEntrySize)) Then
         If MEM(MemPos_TFATable + (lngLoop * TFAEntrySize) + 1) = VOL Then
            MEM(MemPos_Status) = 10 ' resource in use
            Exit Sub
         End If
      End If
   Next
   ' --- close volume ---
   LET_MEMTF MemPos_VolTable + (VOL * VolEntrySize), False
   MEM(MemPos_Status) = 0
   MEM(MemPos_Vol) = 255
End Sub

Public Function OPENDIRECTORY(ByVal Value As Long) As Boolean
   ' --- VOL=255 allows the file to be opened across all volumes ---
   ' --- check for bad parameters ---
   If (Value < 0) Or (Value > MaxFile) Or _
      (VOL < 0) Or (VOL > MaxVolume And VOL <> 255) Or _
      (KEY = "") Or (Len(KEY) > 8) Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check if volume is open ---
   If VOL <> 255 Then
      If Not MEMTF(MemPos_VolTable + (VOL * VolEntrySize)) Then
         MEM(MemPos_Status) = 1 ' volume not open
         GoTo ErrorFound
      End If
   End If
   ' --- close file if already open ---
   If MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' open
      CLOSEFILE Value
   End If
   ' --- get file info ---
   Files(Value).FileName = KEY
   If VOL = 255 Then
      Files(Value).Device = 255
      Files(Value).Volume = ""
      Files(Value).AdjVolume = ""
   Else
      Files(Value).Device = MEM(MemPos_VolTable + (VOL * VolEntrySize) + 1)
      Files(Value).Volume = GetAlpha(MemPos_VolTable + (VOL * VolEntrySize) + 2)
      Files(Value).AdjVolume = AdjustFilenameSQL(Files(Value).Volume)
   End If
   Set Files(Value).CadolXrefs = GetCadolXrefs(AdjustFilenameSQL(KEY))
   If Files(Value).CadolXrefs Is Nothing Then GoTo ErrorFound
   ' --- store info in memory ---
   LET_MEMTF MemPos_FileTable + (Value * FileEntrySize), True ' open
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 1) = VOL ' volume number
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 2) = 5 ' directory file
   LetAlpha MemPos_FileTable + (Value * FileEntrySize) + 3, KEY ' filename
   ' --- done ---
   OPENDIRECTORY = True
   Exit Function
ErrorFound:
   OPENDIRECTORY = False
End Function

Public Function OPENDIRDEVICE(ByVal Value As Long) As Boolean
   ThrowError "OPENDIRDEVICE", "Command not implemented in IDRIS"
End Function

Public Function OPENDIRVOLUME(ByVal Value As Long) As Boolean
   ThrowError "OPENDIRVOLUME", "Command not implemented in IDRIS"
End Function

Public Function OPENDIRTFA(ByVal Value As Long) As Boolean
   ThrowError "OPENDIRTFA", "Command not implemented in IDRIS"
End Function

Public Function OPENDIRLIB(ByVal Value As Long) As Boolean
   ThrowError "OPENDIRLIB", "Command not implemented in IDRIS"
End Function

Public Function OPENDATA(ByVal Value As Long) As Boolean
   ' --- VOL=255 allows the file to be opened across all volumes ---
   ' --- check for bad parameters ---
   If (Value < 0) Or (Value > MaxFile) Or _
      (VOL < 0) Or (VOL > MaxVolume And VOL <> 255) Or _
      (KEY = "") Or (Len(KEY) > 8) Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check if volume is open ---
   If VOL <> 255 Then
      If Not MEMTF(MemPos_VolTable + (VOL * VolEntrySize)) Then
         MEM(MemPos_Status) = 1 ' volume not open
         GoTo ErrorFound
      End If
   End If
   ' --- close file if already open ---
   If MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' open
      CLOSEFILE Value
   End If
   ' --- get file info ---
   Files(Value).FileName = KEY
   If VOL = 255 Then
      Files(Value).Device = 255
      Files(Value).Volume = ""
      Files(Value).AdjVolume = ""
   Else
      Files(Value).Device = MEM(MemPos_VolTable + (VOL * VolEntrySize) + 1)
      Files(Value).Volume = GetAlpha(MemPos_VolTable + (VOL * VolEntrySize) + 2)
      Files(Value).AdjVolume = AdjustFilenameSQL(Files(Value).Volume)
   End If
   Set Files(Value).CadolXrefs = GetCadolXrefs(AdjustFilenameSQL(KEY))
   If Files(Value).CadolXrefs Is Nothing Then GoTo ErrorFound
   ' --- store info in memory ---
   LET_MEMTF MemPos_FileTable + (Value * FileEntrySize), True ' open
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 1) = VOL ' volume number
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 2) = 2 ' data file
   LetAlpha MemPos_FileTable + (Value * FileEntrySize) + 3, KEY ' filename
   ' --- done ---
   OPENDATA = True
   Exit Function
ErrorFound:
   OPENDATA = False
End Function

Public Function OPENSORTFILE(ByVal Value As Long) As Boolean
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Or _
      (VOL < 0) Or (VOL > MaxVolume) Or _
      (KEY = "") Or (Len(KEY) > 8) Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check if volume is open ---
   If Not MEMTF(MemPos_VolTable + (VOL * VolEntrySize)) Then
      MEM(MemPos_Status) = 1 ' volume not open
      GoTo ErrorFound
   End If
   ' --- close file if already open ---
   If MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' open
      CLOSEFILE Value
   End If
   ' --- note: sortfiles don't have cadol info! ---
   ' --- store info in memory ---
   LET_MEMTF MemPos_FileTable + (Value * FileEntrySize), True ' open
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 1) = VOL ' volume number
   MEM(MemPos_FileTable + (Value * FileEntrySize) + 2) = 6 ' pseudo file
   LetAlpha MemPos_FileTable + (Value * FileEntrySize) + 3, KEY ' filename
   ' --- done ---
   OPENSORTFILE = True
   Exit Function
ErrorFound:
   OPENSORTFILE = False
End Function

Public Sub CLOSEFILE(ByVal Value As Long)
   Dim lngLoop As Long
   ' -----------------
   If Value < 0 Or Value > MaxFile Then
      Exit Sub ' no error reported
   End If
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then
      Exit Sub ' already closed
   End If
   ' --- clear file info ---
   With Files(Value)
      .FileName = ""
      .Device = -1
      .Volume = ""
      .AdjVolume = ""
      Set .CadolXrefs = Nothing
      .LastKey = ""
      .LastRec = -1
      If .RecSet.State = adStateOpen Then
         .RecSet.Close
      End If
   End With
   ' --- clear memory slot ---
   LET_MEMTF MemPos_FileTable + (Value * FileEntrySize), False ' closed
   For lngLoop = 1 To FileEntrySize - 1 ' rest of bytes
      MEM(MemPos_FileTable + (Value * FileEntrySize) + lngLoop) = 0
   Next lngLoop
End Sub

' -------------------------------------
' --- Data Storage Input Statements ---
' -------------------------------------

Public Function READFILE(ByVal Value As Long) As Boolean
   ' -----------------------------------------------
   ' --- Inputs:
   ' ---    KEY, key value desired, "" = any
   ' ---    LockFlag=False, normal read
   ' ---    LockFlag=True, read and lock record
   ' -----------------------------------------------
   ' --- Outputs:
   ' ---    ErrorHandler, fatal error
   ' ---    ReadFile=True, record found
   ' ---       KEY, REC, R buffer, LENGTH
   ' ---       rsLockedRec when LockFlag=True
   ' ---    ReadFile=False, record not found
   ' ---       Status=0, not found
   ' ---       Status=2, not found, past end of file
   ' -----------------------------------------------
   Dim aData() As Byte
   Dim strSQL As String
   Dim strWhere As String
   Dim strKey As String
   Dim strTemp As String
   Dim lngLen As Long
   Dim lngLoop As Long
   Dim lngTempRec As Long
   Dim strSQLTable As String
   Dim oCadolXref As rtCadolXref
   ' ---------------------------
   CheckDoEvents
   If EXITING Then GoTo ErrorFound
   ' --- clear REC before starting ---
   REC = 0
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "READFILE", "File not open"
      GoTo ErrorFound
   End If
   ' --- check if already have a record locked ---
   If LockFlag And HasLockedRec Then
      ThrowError "READFILE", "Attempt to lock multiple records"
      GoTo ErrorFound
   End If
   ' --- read specified record ---
   With Files(Value)
      ' --- build sql query ---
      strSQL = ""
      strKey = KEY ' save for speed
      For Each oCadolXref In .CadolXrefs
         ' --- turn off LockFlag if ReadOnly and not a common table ---
         If ReadOnly And LockFlag Then
            If GetClientWhere(oCadolXref.SQLTableName) <> "" Then
               LockFlag = False
            End If
         End If
         ' --- build sql statement ---
         If strSQL <> "" Then strSQL = strSQL & "UNION "
         strSQL = strSQL & "SELECT REC, [KEY], "
         strSQL = strSQL & "'" & oCadolXref.SQLTableName & "' AS SQLTABLE, PACKED_DATA "
         strSQL = strSQL & "FROM [" & oCadolXref.SQLTableName & "] "
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
         If ClientList <> "" Then
            strTemp = GetClientWhere(oCadolXref.SQLTableName)
            If strTemp <> "" Then
               strSQL = strSQL & strWhere & strTemp
               strWhere = "AND"
            End If
         End If
         If strKey <> "" Then
            strSQL = strSQL & strWhere & " [KEY] = '" & FixSqlStr(strKey) & "' "
            strWhere = "AND"
            If HasLetters(strKey) Then
               strSQL = strSQL & strWhere & " CAST([KEY] AS VARBINARY(20)) = CAST('" & _
                        FixSqlStr(strKey) & "' AS VARBINARY(20)) "
               strWhere = "AND"
            End If
         End If
      Next
      strSQL = strSQL & "ORDER BY REC"
      ' --- close recordset if open ---
      If rsRecord.State = adStateOpen Then
         rsRecord.Close
      End If
      ' --- jump back here if read locked record failed ---
TryAgain:
      ' --- read record ---
      On Error GoTo SQLError
      ' --- this is static data. adUseClient is fine. ---
      rsRecord.CursorLocation = adUseClient
      rsRecord.CursorType = adOpenStatic
      rsRecord.LockType = adLockReadOnly
      rsRecord.CacheSize = 1
      rsRecord.MaxRecords = 1 ' same as "TOP 1"
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      rsRecord.ActiveConnection = cnSQL
      rsRecord.Open strSQL, , , , adCmdText
      rsRecord.ActiveConnection = Nothing
      On Error GoTo 0
      ' --- exit if none found ---
      If rsRecord.EOF Then
         rsRecord.Close
         If KEY = "" Then
            MEM(MemPos_Status) = 2 ' past end of file
         Else
            MEM(MemPos_Status) = 0 ' no more records
         End If
         GoTo ErrorFound
      End If
      ' --- handle reading a locked record ---
      If LockFlag Then
         lngTempRec = rsRecord.Fields("REC")
         strSQLTable = rsRecord.Fields("SQLTABLE")
         rsRecord.Close ' done with this recordset
         For Each oCadolXref In .CadolXrefs
            If oCadolXref.SQLTableName = strSQLTable Then Exit For
         Next
         If oCadolXref Is Nothing Then
            MEM(MemPos_Status) = 63 ' unknown error
            GoTo ErrorFound
         End If
         ' --- record may have been deleted ---
         If Not ReadLockedRec(Value, lngTempRec, oCadolXref) Then
            If EXITING Then GoTo ErrorFound
            If LOCKVAL = 1 Then GoTo ErrorFound
            GoTo TryAgain
         End If
      Else
         ' --- store record data in Cadol variables/buffers ---
         REC = rsRecord.Fields("REC")
         LET_KEY rsRecord.Fields("KEY")
         aData = rsRecord.Fields("PACKED_DATA")
         rsRecord.Close ' done with this recordset
         lngLen = UBound(aData) + 1
         LET_LENGTH lngLen
         ' --- init r buffer pointers ---
         INIT_R
         INIT_IR
         ' --- move record to R buffer ---
         For lngLoop = 0 To lngLen - 1
            MEM(MemPos_R + lngLoop) = aData(lngLoop)
         Next lngLoop
      End If
   End With
   ' --- return result ---
   READFILE = True
   Exit Function
ConnError:
   ThrowError "READFILE", "SQL Connection Error:"
   GoTo ErrorFound
SQLError:
   ThrowError "READFILE", "Cannot execute SQL query: " & vbCrLf & strSQL
   Resume ErrorFound
ErrorFound:
   READFILE = False
End Function

Public Function READKEY(ByVal Value As Long) As Boolean
   ' ------------------------------------------------
   ' --- Inputs:
   ' ---    KEY, key value desired, "" = any
   ' ---    REC, last record# read, 0 = start of file
   ' ---    LockFlag=False, normal read
   ' ---    LockFlag=True, read and lock record
   ' ------------------------------------------------
   ' --- Outputs:
   ' ---    ErrorHandler, fatal error
   ' ---    ReadKey=True, record found
   ' ---       KEY, REC, R buffer, LENGTH
   ' ---       rsLockedRec when LockFlag=True
   ' ---    ReadKey=False, record not found
   ' ---       Status=0, not found (never used)
   ' ---       Status=2, not found, past end of file
   ' ------------------------------------------------
   Dim aData() As Byte
   Dim strSQL As String
   Dim strWhere As String
   Dim strKey As String
   Dim strTemp As String
   Dim lngLen As Long
   Dim lngLoop As Long
   Dim lngTempRec As Long
   Dim strSQLTable As String
   Dim oCadolXref As rtCadolXref
   ' ---------------------------
   CheckDoEvents
   If EXITING Then GoTo ErrorFound
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "READKEY", "File not open"
      GoTo ErrorFound
   End If
   ' --- check if already have a record locked ---
   If LockFlag And HasLockedRec Then
      ThrowError "READKEY", "Attempt to lock multiple records"
      GoTo ErrorFound
   End If
   ' --- read specified record ---
   With Files(Value)
      ' --- check if reading the same set of records ---
      If KEY = .LastKey And REC = .LastRec And .RecSet.State = adStateOpen Then
         GoTo GetNextRec
      End If
      ' --- save requested key, not returned key ---
      .LastKey = KEY
      ' --- clear ReadKey storage variables ---
      .LastRec = -1
      ' --- build sql query ---
      strSQL = ""
      strKey = KEY ' save for speed
      For Each oCadolXref In .CadolXrefs
         ' --- turn off LockFlag if ReadOnly and not a common table ---
         If ReadOnly And LockFlag Then
            If GetClientWhere(oCadolXref.SQLTableName) <> "" Then
               LockFlag = False
            End If
         End If
         ' --- check for subquery file to make reports faster ---
         If SQLSubQueryFile <> "" Then
            If oCadolXref.SQLTableName <> UCase$(SQLSubQueryFile) Then GoTo NextSQLTable
         End If
         If strSQL <> "" Then strSQL = strSQL & "UNION "
         strSQL = strSQL & "SELECT REC, [KEY], "
         strSQL = strSQL & "'" & oCadolXref.SQLTableName & "' AS SQLTABLE, PACKED_DATA "
         strSQL = strSQL & "FROM [" & oCadolXref.SQLTableName & "] "
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
         If ClientList <> "" Then
            strTemp = GetClientWhere(oCadolXref.SQLTableName)
            If strTemp <> "" Then
               strSQL = strSQL & strWhere & strTemp
               strWhere = "AND"
            End If
         End If
         If strKey <> "" Then
            strSQL = strSQL & strWhere & " [KEY] = '" & FixSqlStr(strKey) & "' "
            strWhere = "AND"
            If HasLetters(strKey) Then
               strSQL = strSQL & strWhere & " CAST([KEY] AS VARBINARY(20)) = CAST('" & _
                        FixSqlStr(strKey) & "' AS VARBINARY(20)) "
               strWhere = "AND"
            End If
         End If
         If REC > 0 Then
            strSQL = strSQL & strWhere & " REC > " & Trim$(Str$(REC)) & " "
            strWhere = "AND"
         End If
         If SQLSubQuery <> "" Then
            strSQL = strSQL & strWhere & " " & SQLSubQuery & " "
            strWhere = "AND"
            If DebugFlag And DebugFlagLevel > 0 Then
               DebugMessage "SUBQUERY: " & SQLSubQuery
            End If
         End If
NextSQLTable:
      Next
      ' --- make sure records are in order ---
      strSQL = strSQL & "ORDER BY REC "
      ' --- get first record as fast as possible ---
      strSQL = strSQL & "OPTION (FAST 1) "
      ' --- clear subquery data ---
      SQLSubQueryFile = ""
      SQLSubQuery = ""
      ' --- jump back here if read locked record failed ---
TryAgain:
      On Error GoTo SQLError
      ' --- close recordset if open ---
      If .RecSet.State = adStateOpen Then
         .RecSet.Close
      End If
      ' --- read records ---
      ' --- this is static data. adUseClient is fine. ---
      .RecSet.CursorLocation = adUseClient
      .RecSet.CursorType = adOpenStatic
      .RecSet.LockType = adLockReadOnly
      If LockFlag Then
         .RecSet.CacheSize = 1 ' prevent missing record problems
      Else
         .RecSet.CacheSize = ReadKeyCacheSize
      End If
      .RecSet.MaxRecords = 0 ' all records
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .RecSet.ActiveConnection = cnSQL
      .RecSet.Open strSQL, , , , adCmdText
      On Error GoTo 0
      ' --- arrive here if still using same recordset ---
GetNextRec:
      ' --- exit if none found ---
      If .RecSet.EOF Then
         .RecSet.Close
         ' --- clear ReadKey storage variables ---
         .LastKey = ""
         .LastRec = -1
         ' --- set return status value ---
         If KEY = "" Then
            MEM(MemPos_Status) = 2 ' past end of file
         Else
            MEM(MemPos_Status) = 0 ' no more records
         End If
         GoTo ErrorFound
      End If
      ' --- handle reading a locked record ---
      If LockFlag Then
         ' --- get values needed to read a locked record ---
         lngTempRec = .RecSet.Fields("REC")
         strSQLTable = .RecSet.Fields("SQLTABLE")
         .RecSet.MoveNext ' get off desired record
         For Each oCadolXref In .CadolXrefs
            If oCadolXref.SQLTableName = strSQLTable Then Exit For
         Next
         If oCadolXref Is Nothing Then
            MEM(MemPos_Status) = 63 ' unknown error
            GoTo ErrorFound
         End If
         ' --- record may have been deleted ---
         If Not ReadLockedRec(Value, lngTempRec, oCadolXref) Then
            If EXITING Then GoTo ErrorFound
            If LOCKVAL = 1 Then GoTo ErrorFound
            GoTo TryAgain
         End If
         ' --- prepare for next readkey ---
         .LastRec = REC
      Else
         ' --- store record data in Cadol variables/buffers ---
         REC = .RecSet.Fields("REC")
         LET_KEY .RecSet.Fields("KEY")
         aData = .RecSet.Fields("PACKED_DATA")
         .RecSet.MoveNext ' done with this record
         lngLen = UBound(aData) + 1
         LET_LENGTH lngLen
         ' --- prepare for next readkey ---
         .LastRec = REC
         ' --- init r buffer pointers ---
         INIT_R
         INIT_IR
         ' --- move record to R buffer ---
         For lngLoop = 0 To lngLen - 1
            MEM(MemPos_R + lngLoop) = aData(lngLoop)
         Next lngLoop
      End If
   End With
   ' --- return result ---
   READKEY = True
   Exit Function
ConnError:
   ThrowError "READKEY", "SQL Connection Error:"
   GoTo ErrorFound
SQLError:
   ThrowError "READKEY", "Cannot execute SQL query: " & vbCrLf & strSQL
   Resume ErrorFound
ErrorFound:
   READKEY = False
End Function

Public Function READREC(ByVal Value As Long) As Boolean
   ' -------------------------------------------
   ' --- Inputs:
   ' ---    REC, record number desired
   ' ---    LockFlag=False, normal read
   ' ---    LockFlag=True, read and lock record
   ' -------------------------------------------
   ' --- Outputs:
   ' ---    ErrorHandler, fatal error
   ' ---    ReadRec=True, record found
   ' ---       KEY, REC, R buffer, LENGTH
   ' ---       rsLockedRec when LockFlag=True
   ' ---    ReadRec=False, record not found
   ' ---       Status=0, specified slot is empty
   ' ---       Status=4, REC is past end of file
   ' -------------------------------------------
   Dim aData() As Byte
   Dim strSQL As String
   Dim strTemp As String
   Dim strWhere As String
   Dim lngLen As Long
   Dim lngLoop As Long
   Dim lngTempRec As Long
   Dim strSQLTable As String
   Dim oCadolXref As rtCadolXref
   ' ---------------------------
   CheckDoEvents
   If EXITING Then GoTo ErrorFound
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "READREC", "File not open"
      GoTo ErrorFound
   End If
   ' --- check if already have a record locked ---
   If LockFlag And HasLockedRec Then
      ThrowError "READREC", "Attempt to lock multiple records"
      GoTo ErrorFound
   End If
   ' --- check for invalid record number. just exit, don't throw an error. ---
   If REC < 0 Then GoTo ErrorFound
   ' --- read specified record ---
   With Files(Value)
      ' --- build sql query ---
      strSQL = ""
      For Each oCadolXref In .CadolXrefs
         ' --- turn off LockFlag if ReadOnly and not a common table ---
         If ReadOnly And LockFlag Then
            If GetClientWhere(oCadolXref.SQLTableName) <> "" Then
               LockFlag = False
            End If
         End If
         ' --- build sql statement ---
         If strSQL <> "" Then strSQL = strSQL & "UNION "
         strSQL = strSQL & "SELECT REC, [KEY], "
         strSQL = strSQL & "'" & oCadolXref.SQLTableName & "' AS SQLTABLE, PACKED_DATA "
         strSQL = strSQL & "FROM [" & oCadolXref.SQLTableName & "] "
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
         If ClientList <> "" Then
            strTemp = GetClientWhere(oCadolXref.SQLTableName)
            If strTemp <> "" Then
               strSQL = strSQL & strWhere & strTemp
               strWhere = "AND"
            End If
         End If
         strSQL = strSQL & strWhere & " REC = " & Trim$(Str$(REC)) & " "
         strWhere = "AND"
      Next
      ' --- close recordset if open ---
      If rsRecord.State = adStateOpen Then
         rsRecord.Close
      End If
      ' --- read record ---
      On Error GoTo SQLError
      ' --- this is static data. adUseClient is fine. ---
      rsRecord.CursorLocation = adUseClient
      rsRecord.CursorType = adOpenStatic
      rsRecord.LockType = adLockReadOnly
      rsRecord.CacheSize = 1
      rsRecord.MaxRecords = 1 ' same as "TOP 1"
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      rsRecord.ActiveConnection = cnSQL
      rsRecord.Open strSQL, , , , adCmdText
      rsRecord.ActiveConnection = Nothing
      On Error GoTo 0
      ' --- exit if none found ---
      If rsRecord.EOF Then
         rsRecord.Close
         MEM(MemPos_Status) = 0 ' no more records
         GoTo ErrorFound
      End If
      ' --- handle reading a locked record ---
      If LockFlag Then
         lngTempRec = rsRecord.Fields("REC")
         strSQLTable = rsRecord.Fields("SQLTABLE")
         rsRecord.Close ' done with this recordset
         For Each oCadolXref In .CadolXrefs
            If oCadolXref.SQLTableName = strSQLTable Then Exit For
         Next
         If oCadolXref Is Nothing Then
            MEM(MemPos_Status) = 63 ' unknown error
            GoTo ErrorFound
         End If
         ' --- record may have been deleted ---
         If Not ReadLockedRec(Value, lngTempRec, oCadolXref) Then
            If EXITING Then GoTo ErrorFound
            If LOCKVAL = 1 Then GoTo ErrorFound
            MEM(MemPos_Status) = 0 ' no more records
            GoTo ErrorFound
         End If
      Else
         ' --- store record data in Cadol variables/buffers ---
         REC = rsRecord.Fields("REC")
         LET_KEY rsRecord.Fields("KEY")
         aData = rsRecord.Fields("PACKED_DATA")
         rsRecord.Close ' done with this recordset
         lngLen = UBound(aData) + 1
         LET_LENGTH lngLen
         ' --- init r buffer pointers ---
         INIT_R
         INIT_IR
         ' --- move record to R buffer ---
         For lngLoop = 0 To lngLen - 1
            MEM(MemPos_R + lngLoop) = aData(lngLoop)
         Next lngLoop
      End If
   End With
   ' --- return result ---
   READREC = True
   Exit Function
ConnError:
   ThrowError "READREC", "SQL Connection Error:"
   GoTo ErrorFound
SQLError:
   ThrowError "READREC", "Cannot execute SQL query: " & vbCrLf & strSQL
   Resume ErrorFound
ErrorFound:
   READREC = False
End Function

' ----------------------------------------
' --- Record Access Control Statements ---
' ----------------------------------------

Public Sub LOCKREC()
   ' --- next attempt to read will lock a record ---
   LockFlag = True
End Sub

Public Sub UNLOCKREC()
   ' --- check if record is currently locked ---
   If HasLockedRec Then
      On Error Resume Next
      If rsLockedRec.State = adStateOpen Then
         rsLockedRec.CancelUpdate
         rsLockedRec.Close
      End If
      On Error GoTo 0
      ReleaseAppLock LockedResource
      LockedResource = ""
      LockedFileNum = -1
      LockedRecNum = -1
      LockedRecLen = -1
      Set LockedCadolXref = Nothing
      HasLockedRec = False
   End If
   ' --- clear the lock flag ---
   LockFlag = False
End Sub

' --------------------------------------
' --- Data Storage Output Statements ---
' --------------------------------------

Public Function WRITEFILE(ByVal Value As Long) As Boolean
   Dim strSQL As String
   Dim lngRecLen As Long
   Dim oCadolXref As rtCadolXref
   Dim rsWrite As ADODB.Recordset
   ' ----------------------------
   CheckDoEvents
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "WRITEFILE", "File not open"
      GoTo ErrorFound
   End If
   ' --- don't write anything if in ReadOnly mode ---
   If ReadOnly Then
      If Not LockFlag Then GoTo Done
   End If
   ' --- find proper format for current record ---
   For Each oCadolXref In Files(Value).CadolXrefs
      If Files(Value).CadolXrefs.Count = 1 Then GoTo FoundFormat
      With oCadolXref
         If .CadolKey <> "" Then
            If Not MatchKeyPattern(KEY, .CadolKey) Then GoTo NextSQLFile
         End If
         If .CadolLength > 0 Then
            If GetNumeric(MemPos_W + .CadolByte, .CadolLength) <> .CadolValue Then
               GoTo NextSQLFile
            End If
         End If
      End With
      GoTo FoundFormat
NextSQLFile:
   Next
   MEM(MemPos_Status) = 255 ' cannot determine proper format (database damaged)
   GoTo ErrorFound
FoundFormat:
   Set rsWrite = New ADODB.Recordset
   ' --- build a sql query that returns no records ---
   strSQL = "SELECT TOP 0 * FROM [" & oCadolXref.SQLTableName & "]"
   ' --- open recordset and prepare for adding ---
   With rsWrite
      ' --- must be adUseServer for trigger issues ---
      .CursorLocation = adUseServer
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      If cnSQL Is Nothing Then GoTo ConnError
      If cnSQL.Errors.Count > 0 Then GoTo ConnError
      .ActiveConnection = cnSQL
      .Open strSQL, , , , adCmdText
      .AddNew
   End With
   ' --- unpack record ---
   lngRecLen = MEM(MemPos_WP)
   If lngRecLen = 0 Then lngRecLen = 256
   If Not UnpackFields(rsWrite, oCadolXref, MemPos_W, lngRecLen) Then
      ThrowError "WRITEFILE", "Error unpacking record"
      GoTo ErrorFound
   End If
   ' --- update header fields ---
   rsWrite.Fields("KEY") = KEY
   rsWrite.Fields("VOLUME") = Files(Value).AdjVolume
   rsWrite.Fields("DEVICE") = Files(Value).Device
   If oCadolXref.Multiple Then
      rsWrite.Fields("FILENAME") = Files(Value).FileName
   End If
   If UpdateRecField Then
      rsWrite.Fields("REC") = REC ' can only work if Rec is not the Identity field
      UpdateRecField = False
   End If
   ' --- update ChangedBy ---
   rsWrite.Fields("ChangedBy") = LoginID
   ' --- save data ---
   rsWrite.Update
   ' --- return the record number of the new record ---
   REC = rsWrite.Fields("REC")
   ' --- done with recordset ---
   rsWrite.Close
   ' --- init w buffer pointers after write ---
Done:
   INIT_W
   INIT_IW
   ' --- done ---
   WRITEFILE = True
   Exit Function
ConnError:
   ThrowError "WRITEFILE", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   UpdateRecField = False
   WRITEFILE = False
End Function

Public Function WRITEKEY(ByVal Value As Long) As Boolean
   WRITEKEY = WRITEFILE(Value)
End Function

Public Function WRITEREC(ByVal Value As Long) As Boolean
   Dim lngRecLen As Long
   Dim lngLoop As Long
   Dim aPacked() As Byte
   ' ----------------------
   CheckDoEvents
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "WRITEREC", "File not open"
      GoTo ErrorFound
   End If
   ' --- don't write anything if in ReadOnly mode ---
   If ReadOnly Then
      If Not LockFlag Then GoTo Done
   End If
   ' --- Check if doing WriteRec to add a new record ---
   If Not HasLockedRec Then
      UpdateRecField = True ' will update the REC field on the new record
      WRITEREC = WRITEFILE(Value)
      Exit Function
   End If
   ' --- check for errors ---
   If Value <> LockedFileNum Then
      SYSERROR "WRITING UNLOCKED RECORD"
      GoTo ErrorFound
   End If
   If REC <> rsLockedRec.Fields("REC") Then
      SYSERROR "WRITING UNLOCKED RECORD"
      GoTo ErrorFound
   End If
   If Files(Value).Device <> 255 Then
      If Files(Value).Device <> rsLockedRec.Fields("DEVICE") Then
         ThrowError "WRITEREC", "Device may not be changed using WriteRec"
         GoTo ErrorFound
      End If
   End If
   If Files(Value).AdjVolume <> "" Then
      If Files(Value).AdjVolume <> rsLockedRec.Fields("VOLUME") Then
         ThrowError "WRITEREC", "Volume may not be changed using WriteRec"
         GoTo ErrorFound
      End If
   End If
   ' --- get record length ---
   lngRecLen = MEM(MemPos_WP)
   If lngRecLen = 0 Then lngRecLen = 256
   ' --- check if record has changed ---
   If lngRecLen = rsLockedRec.Fields("PACKED_DATA").ActualSize Then
      aPacked = rsLockedRec.Fields("PACKED_DATA")
      For lngLoop = 0 To lngRecLen - 1
         If aPacked(lngLoop) <> MEM(MemPos_W + lngLoop) Then GoTo DoUnpack
      Next lngLoop
      If rsLockedRec.State <> adEditNone Then
         rsLockedRec.CancelUpdate
      End If
      GoTo DoUnlockRec
   End If
   ' --- unpack record ---
DoUnpack:
   If Not UnpackFields(rsLockedRec, LockedCadolXref, MemPos_W, lngRecLen) Then
      ThrowError "WRITEREC", "Error unpacking record"
      GoTo ErrorFound
   End If
   ' --- update ChangedBy ---
   rsLockedRec.Fields("ChangedBy") = LoginID
   ' --- update record back into SQL ---
   On Error GoTo ErrorFound
   rsLockedRec.Update
   On Error GoTo 0
   ' --- unlock record ---
DoUnlockRec:
   rsLockedRec.Close
   ReleaseAppLock LockedResource
   LockedResource = ""
   LockedFileNum = -1
   LockedRecNum = -1
   LockedRecLen = -1
   Set LockedCadolXref = Nothing
   HasLockedRec = False
   ' --- init w buffer pointers after write ---
Done:
   INIT_W
   INIT_IW
   ' --- return result ---
   WRITEREC = True
   Exit Function
ErrorFound:
   WRITEREC = False
End Function

Public Sub WRITEBACK()
   Dim lngLoop As Long
   Dim lngRecLen As Long
   Dim aPacked() As Byte
   ' -------------------
   ' --- MUSTEXIT is used to prevent infinite loops when an error occurs ---
   If Not EXITING Then
      MUSTEXIT = False ' clear error flag
   End If
   CheckDoEvents
   If EXITING Then
      MUSTEXIT = True ' turn on error flag
      Exit Sub
   End If
   ' --- don't write anything if in ReadOnly mode ---
   If ReadOnly Then
      If Not LockFlag Then GoTo Done
   End If
   ' --- check for errors ---
   If Not HasLockedRec Then
      ThrowError "WRITEBACK", "WRITING UNLOCKED RECORD"
      Exit Sub
   End If
   With Files(LockedFileNum)
      If .Device <> 255 Then
         If .Device <> rsLockedRec.Fields("DEVICE") Then
            ThrowError "WRITEBACK", "Device may not be changed using WriteBack"
            Exit Sub
         End If
      End If
      If .AdjVolume <> "" Then
         If .AdjVolume <> rsLockedRec.Fields("VOLUME") Then
            ThrowError "WRITEBACK", "Volume may not be changed using WriteBack"
            Exit Sub
         End If
      End If
      ' --- get record length ---
      lngRecLen = LockedRecLen
      If lngRecLen = 0 Then lngRecLen = 256
      ' --- check if record has changed ---
      If lngRecLen = rsLockedRec.Fields("PACKED_DATA").ActualSize Then
         aPacked = rsLockedRec.Fields("PACKED_DATA")
         For lngLoop = 0 To lngRecLen - 1
            If aPacked(lngLoop) <> MEM(MemPos_R + lngLoop) Then GoTo DoUnpack
         Next lngLoop
         If rsLockedRec.State <> adEditNone Then
            rsLockedRec.CancelUpdate
         End If
         GoTo DoneWith
      End If
      ' --- unpack record ---
DoUnpack:
      If Not UnpackFields(rsLockedRec, LockedCadolXref, MemPos_R, lngRecLen) Then
         ThrowError "WRITEBACK", "Error unpacking record"
         Exit Sub
      End If
      ' --- update ChangedBy ---
      rsLockedRec.Fields("ChangedBy") = LoginID
      ' --- update record back into SQL ---
      On Error GoTo ErrorFound
      rsLockedRec.Update
      On Error GoTo 0
DoneWith:
   End With
   ' --- unlock record ---
   rsLockedRec.Close
   ReleaseAppLock LockedResource
   LockedResource = ""
   LockedFileNum = -1
   LockedRecNum = -1
   LockedRecLen = -1
   Set LockedCadolXref = Nothing
   HasLockedRec = False
Done:
   Exit Sub
ErrorFound:
   ThrowError "WRITEBACK", "Unable to properly update record" & vbCrLf & Err.Description
   Exit Sub
End Sub

Public Function DELETE(ByVal Value As Long) As Boolean
   Dim strSQL As String
   Dim strTemp As String
   Dim strWhere As String
   Dim oCadolXref As rtCadolXref
   ' ---------------------------
   CheckDoEvents
   If EXITING Then GoTo ErrorFound
   ' --- check for bad parameters ---
   If Value < 0 Or Value > MaxFile Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check parameters ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then ' closed
      ThrowError "DELETE", "File not open"
      GoTo ErrorFound
   End If
   ' --- don't delete anything if in ReadOnly mode ---
   If ReadOnly Then
      If Not LockFlag Then GoTo Done
   End If
   ' --- check if locked record will be the one deleted ---
   If HasLockedRec And LockedFileNum = Value And LockedRecNum = REC Then
      On Error Resume Next
      If rsLockedRec.State = adStateOpen Then
         rsLockedRec.CancelUpdate
         rsLockedRec.Close
      End If
      On Error GoTo 0
   End If
   ' --- delete all matching records ---
   With Files(Value)
      ' --- build sql query ---
      strSQL = ""
      For Each oCadolXref In .CadolXrefs
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
         If ClientList <> "" Then
            strTemp = GetClientWhere(oCadolXref.SQLTableName)
            If strTemp <> "" Then
               strSQL = strSQL & strWhere & strTemp
               strWhere = "AND"
            End If
         End If
         strSQL = strSQL & strWhere & " REC = " & Trim$(Str$(REC)) & " "
         strWhere = "AND"
      Next
      ' --- execute the sql query ---
      If strSQL <> "" Then
         If cnSQL Is Nothing Then GoTo ConnError
         If cnSQL.Errors.Count > 0 Then GoTo ConnError
         cnSQL.Execute strSQL, , adCmdText
      End If
   End With
   ' --- release internal locks ---
   If HasLockedRec And LockedFileNum = Value And LockedRecNum = REC Then
      ReleaseAppLock LockedResource
      LockedResource = ""
      LockedFileNum = -1
      LockedRecNum = -1
      LockedRecLen = -1
      Set LockedCadolXref = Nothing
      LockFlag = False
      HasLockedRec = False
   End If
   ' --- done ---
Done:
   DELETE = True
   Exit Function
ConnError:
   ThrowError "DELETE", "SQL Connection Error:"
   GoTo ErrorFound
ErrorFound:
   DELETE = False
End Function

' --------------------------------------
' --- Buffer Manipulation Statements ---
' --------------------------------------

' --- INIT commands ---

Public Sub INIT_R()
   MEM(MemPos_RP) = 0
   MEM(MemPos_RP2) = MemPos_R \ 256
End Sub
Public Sub INIT_Z()
   MEM(MemPos_ZP) = 0
   MEM(MemPos_ZP2) = MemPos_Z \ 256
End Sub
Public Sub INIT_X()
   MEM(MemPos_XP) = 0
   MEM(MemPos_XP2) = MemPos_X \ 256
End Sub
Public Sub INIT_Y()
   MEM(MemPos_YP) = 0
   MEM(MemPos_YP2) = MemPos_Y \ 256
End Sub
Public Sub INIT_W()
   MEM(MemPos_WP) = 0
   MEM(MemPos_WP2) = MemPos_W \ 256
End Sub
Public Sub INIT_S()
   MEM(MemPos_SP) = 0
   MEM(MemPos_SP2) = MemPos_S \ 256
End Sub
Public Sub INIT_T()
   MEM(MemPos_TP) = 0
   MEM(MemPos_TP2) = MemPos_T \ 256
End Sub
Public Sub INIT_U()
   MEM(MemPos_UP) = 0
   MEM(MemPos_UP2) = MemPos_U \ 256
End Sub
Public Sub INIT_V()
   MEM(MemPos_VP) = 0
   MEM(MemPos_VP2) = MemPos_V \ 256
End Sub

Public Sub INIT_IR()
   MEM(MemPos_IRP) = 0
   MEM(MemPos_IRP2) = MemPos_R \ 256
End Sub
Public Sub INIT_IZ()
   MEM(MemPos_IZP) = 0
   MEM(MemPos_IZP2) = MemPos_Z \ 256
End Sub
Public Sub INIT_IX()
   MEM(MemPos_IXP) = 0
   MEM(MemPos_IXP2) = MemPos_X \ 256
End Sub
Public Sub INIT_IY()
   MEM(MemPos_IYP) = 0
   MEM(MemPos_IYP2) = MemPos_Y \ 256
End Sub
Public Sub INIT_IW()
   MEM(MemPos_IWP) = 0
   MEM(MemPos_IWP2) = MemPos_W \ 256
End Sub
Public Sub INIT_IS()
   MEM(MemPos_ISP) = 0
   MEM(MemPos_ISP2) = MemPos_S \ 256
End Sub
Public Sub INIT_IT()
   MEM(MemPos_ITP) = 0
   MEM(MemPos_ITP2) = MemPos_T \ 256
End Sub
Public Sub INIT_IU()
   MEM(MemPos_IUP) = 0
   MEM(MemPos_IUP2) = MemPos_U \ 256
End Sub
Public Sub INIT_IV()
   MEM(MemPos_IVP) = 0
   MEM(MemPos_IVP2) = MemPos_V \ 256
End Sub

' --- SET commands ---

Public Sub SET_R()
   MEM(MemPos_IRP) = MEM(MemPos_RP)
   MEM(MemPos_IRP2) = MEM(MemPos_RP2)
End Sub
Public Sub SET_Z()
   MEM(MemPos_IZP) = MEM(MemPos_ZP)
   MEM(MemPos_IZP2) = MEM(MemPos_ZP2)
End Sub
Public Sub SET_X()
   MEM(MemPos_IXP) = MEM(MemPos_XP)
   MEM(MemPos_IXP2) = MEM(MemPos_XP2)
End Sub
Public Sub SET_Y()
   MEM(MemPos_IYP) = MEM(MemPos_YP)
   MEM(MemPos_IYP2) = MEM(MemPos_YP2)
End Sub
Public Sub SET_W()
   MEM(MemPos_IWP) = MEM(MemPos_WP)
   MEM(MemPos_IWP2) = MEM(MemPos_WP2)
End Sub
Public Sub SET_S()
   MEM(MemPos_ISP) = MEM(MemPos_SP)
   MEM(MemPos_ISP2) = MEM(MemPos_SP2)
End Sub
Public Sub SET_T()
   MEM(MemPos_ITP) = MEM(MemPos_TP)
   MEM(MemPos_ITP2) = MEM(MemPos_TP2)
End Sub
Public Sub SET_U()
   MEM(MemPos_IUP) = MEM(MemPos_UP)
   MEM(MemPos_IUP2) = MEM(MemPos_UP2)
End Sub
Public Sub SET_V()
   MEM(MemPos_IVP) = MEM(MemPos_VP)
   MEM(MemPos_IVP2) = MEM(MemPos_VP2)
End Sub

' --- RESET commands ---

Public Sub RESET_R()
   MEM(MemPos_RP) = MEM(MemPos_IRP)
   MEM(MemPos_RP2) = MEM(MemPos_IRP2)
End Sub
Public Sub RESET_Z()
   MEM(MemPos_ZP) = MEM(MemPos_IZP)
   MEM(MemPos_ZP2) = MEM(MemPos_IZP2)
End Sub
Public Sub RESET_X()
   MEM(MemPos_XP) = MEM(MemPos_IXP)
   MEM(MemPos_XP2) = MEM(MemPos_IXP2)
End Sub
Public Sub RESET_Y()
   MEM(MemPos_YP) = MEM(MemPos_IYP)
   MEM(MemPos_YP2) = MEM(MemPos_IYP2)
End Sub
Public Sub RESET_W()
   MEM(MemPos_WP) = MEM(MemPos_IWP)
   MEM(MemPos_WP2) = MEM(MemPos_IWP2)
End Sub
Public Sub RESET_S()
   MEM(MemPos_SP) = MEM(MemPos_ISP)
   MEM(MemPos_SP2) = MEM(MemPos_ISP2)
End Sub
Public Sub RESET_T()
   MEM(MemPos_TP) = MEM(MemPos_ITP)
   MEM(MemPos_TP2) = MEM(MemPos_ITP2)
End Sub
Public Sub RESET_U()
   MEM(MemPos_UP) = MEM(MemPos_IUP)
   MEM(MemPos_UP2) = MEM(MemPos_IUP2)
End Sub
Public Sub RESET_V()
   MEM(MemPos_VP) = MEM(MemPos_IVP)
   MEM(MemPos_VP2) = MEM(MemPos_IVP2)
End Sub

' --- SKIP Numeric commands ---

Public Sub SKIP_R(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_RP2) * 256) + MEM(MemPos_RP) + Value
   MEM(MemPos_RP2) = lngPos \ 256
   MEM(MemPos_RP) = lngPos Mod 256
End Sub
Public Sub SKIP_Z(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_ZP2) * 256) + MEM(MemPos_ZP) + Value
   MEM(MemPos_ZP2) = lngPos \ 256
   MEM(MemPos_ZP) = lngPos Mod 256
End Sub
Public Sub SKIP_X(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_XP2) * 256) + MEM(MemPos_XP) + Value
   MEM(MemPos_XP2) = lngPos \ 256
   MEM(MemPos_XP) = lngPos Mod 256
End Sub
Public Sub SKIP_Y(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_YP2) * 256) + MEM(MemPos_YP) + Value
   MEM(MemPos_YP2) = lngPos \ 256
   MEM(MemPos_YP) = lngPos Mod 256
End Sub
Public Sub SKIP_W(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_WP2) * 256) + MEM(MemPos_WP) + Value
   MEM(MemPos_WP2) = lngPos \ 256
   MEM(MemPos_WP) = lngPos Mod 256
End Sub
Public Sub SKIP_S(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_SP2) * 256) + MEM(MemPos_SP) + Value
   MEM(MemPos_SP2) = lngPos \ 256
   MEM(MemPos_SP) = lngPos Mod 256
End Sub
Public Sub SKIP_T(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_TP2) * 256) + MEM(MemPos_TP) + Value
   MEM(MemPos_TP2) = lngPos \ 256
   MEM(MemPos_TP) = lngPos Mod 256
End Sub
Public Sub SKIP_U(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_UP2) * 256) + MEM(MemPos_UP) + Value
   MEM(MemPos_UP2) = lngPos \ 256
   MEM(MemPos_UP) = lngPos Mod 256
End Sub
Public Sub SKIP_V(ByVal Value As Long)
   Dim lngPos As Long
   ' ----------------
   lngPos = (MEM(MemPos_VP2) * 256) + MEM(MemPos_VP) + Value
   MEM(MemPos_VP2) = lngPos \ 256
   MEM(MemPos_VP) = lngPos Mod 256
End Sub

' --- SKIP Alpha commands ---

Public Sub SKIP_R_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = R_A
   Next lngLoop
End Sub
Public Sub SKIP_Z_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = Z_A
   Next lngLoop
End Sub
Public Sub SKIP_X_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = X_A
   Next lngLoop
End Sub
Public Sub SKIP_Y_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = Y_A
   Next lngLoop
End Sub
Public Sub SKIP_W_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = W_A
   Next lngLoop
End Sub
Public Sub SKIP_S_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = S_A
   Next lngLoop
End Sub
Public Sub SKIP_T_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = T_A
   Next lngLoop
End Sub
Public Sub SKIP_U_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = U_A
   Next lngLoop
End Sub
Public Sub SKIP_V_A(ByVal Value As Long)
   Dim lngLoop As Long
   Dim strDummy As String
   ' --------------------
   For lngLoop = 1 To Value
      strDummy = V_A
   Next lngLoop
End Sub

Public Sub MOVE(ByVal Size As Long, _
                ByVal FromPos As Long, ByVal FromOfs As Long, _
                ByVal ToPos As Long, ByVal ToOfs As Long)
   Dim lngCount As Long
   Dim lngActualFrom As Long
   Dim lngActualTo As Long
   ' -----------------------
   If Size < 0 Or Size > 256 Then
      ThrowError "MOVE", "Invalid size: " & Trim$(Str$(Size))
      Exit Sub
   End If
   ' --- adjust size ---
   If Size = 0 Then Size = 256
   ' --- get actual offset values from buffer pointers ---
   Select Case FromPos
      Case MemPos_R: lngActualFrom = (MEM(MemPos_RP2) * 256) + MEM(MemPos_RP)
      Case MemPos_Z: lngActualFrom = (MEM(MemPos_ZP2) * 256) + MEM(MemPos_ZP)
      Case MemPos_X: lngActualFrom = (MEM(MemPos_XP2) * 256) + MEM(MemPos_XP)
      Case MemPos_Y: lngActualFrom = (MEM(MemPos_YP2) * 256) + MEM(MemPos_YP)
      Case MemPos_W: lngActualFrom = (MEM(MemPos_WP2) * 256) + MEM(MemPos_WP)
      Case MemPos_S: lngActualFrom = (MEM(MemPos_SP2) * 256) + MEM(MemPos_SP)
      Case MemPos_T: lngActualFrom = (MEM(MemPos_TP2) * 256) + MEM(MemPos_TP)
      Case MemPos_U: lngActualFrom = (MEM(MemPos_UP2) * 256) + MEM(MemPos_UP)
      Case MemPos_V: lngActualFrom = (MEM(MemPos_VP2) * 256) + MEM(MemPos_VP)
      Case Else: lngActualFrom = FromPos
   End Select
   Select Case ToPos
      Case MemPos_R: lngActualTo = (MEM(MemPos_RP2) * 256) + MEM(MemPos_RP)
      Case MemPos_Z: lngActualTo = (MEM(MemPos_ZP2) * 256) + MEM(MemPos_ZP)
      Case MemPos_X: lngActualTo = (MEM(MemPos_XP2) * 256) + MEM(MemPos_XP)
      Case MemPos_Y: lngActualTo = (MEM(MemPos_YP2) * 256) + MEM(MemPos_YP)
      Case MemPos_W: lngActualTo = (MEM(MemPos_WP2) * 256) + MEM(MemPos_WP)
      Case MemPos_S: lngActualTo = (MEM(MemPos_SP2) * 256) + MEM(MemPos_SP)
      Case MemPos_T: lngActualTo = (MEM(MemPos_TP2) * 256) + MEM(MemPos_TP)
      Case MemPos_U: lngActualTo = (MEM(MemPos_UP2) * 256) + MEM(MemPos_UP)
      Case MemPos_V: lngActualTo = (MEM(MemPos_VP2) * 256) + MEM(MemPos_VP)
      Case Else: lngActualTo = ToPos
   End Select
   ' --- move specified number of chars ---
   lngCount = 0
   Do While lngCount < Size
      MEM(lngActualTo + ToOfs + lngCount) = MEM(lngActualFrom + FromOfs + lngCount)
      lngCount = lngCount + 1
   Loop
   ' --- adjust any buffer pointers needed ---
   If FromPos <> ToPos Then
      Select Case FromPos
         Case MemPos_R: SKIP_R FromOfs + Size
         Case MemPos_Z: SKIP_Z FromOfs + Size
         Case MemPos_X: SKIP_X FromOfs + Size
         Case MemPos_Y: SKIP_Y FromOfs + Size
         Case MemPos_W: SKIP_W FromOfs + Size
         Case MemPos_S: SKIP_S FromOfs + Size
         Case MemPos_T: SKIP_T FromOfs + Size
         Case MemPos_U: SKIP_U FromOfs + Size
         Case MemPos_V: SKIP_V FromOfs + Size
         ' --- no error if not a buffer ---
      End Select
   End If
   Select Case ToPos
      Case MemPos_R: SKIP_R ToOfs + Size
      Case MemPos_Z: SKIP_Z ToOfs + Size
      Case MemPos_X: SKIP_X ToOfs + Size
      Case MemPos_Y: SKIP_Y ToOfs + Size
      Case MemPos_W: SKIP_W ToOfs + Size
      Case MemPos_S: SKIP_S ToOfs + Size
      Case MemPos_T: SKIP_T ToOfs + Size
      Case MemPos_U: SKIP_U ToOfs + Size
      Case MemPos_V: SKIP_V ToOfs + Size
      ' --- no error if not a buffer ---
   End Select
End Sub

Public Function FLIP(ByVal Value As Currency) As Currency
   Dim curTemp As Currency
   ' ---------------------
   curTemp = ModPos(Value, 65536)
   curTemp = ((curTemp Mod 256) * 256) + (curTemp \ 256)
   FLIP = curTemp
End Function

Public Function FLOP(ByVal Value As Currency) As Currency
   Dim curTemp As Currency
   ' ---------------------
   curTemp = ModPos(Value, 65536)
   curTemp = ((curTemp Mod 256) * 256) + (curTemp \ 256)
   If curTemp >= 32768 Then
      curTemp = curTemp - 65536
   Else
      curTemp = curTemp
   End If
   FLOP = curTemp
End Function

Public Function NOSIGN(ByVal Value As Currency) As Currency
   Dim curTemp As Currency
   ' ---------------------
   curTemp = ModPos(Value, 65536)
   NOSIGN = curTemp
End Function

Public Function SIGNSET(ByVal Value As Currency) As Currency
   Dim curTemp As Currency
   ' ---------------------
   curTemp = ModPos(Value, 65536)
   If curTemp >= 32768 Then
      curTemp = curTemp - 65536
   Else
      curTemp = curTemp
   End If
   SIGNSET = curTemp
End Function

' -----------------------
' --- Sort Statements ---
' -----------------------

Public Sub INITSORT(ByVal Value As Long)
   Dim lngLoop As Long
   ' -----------------
   ' --- close any open sort file unless MUSFIP ---
   If MEM(MemPos_SortState) > 1 Then
      CloseSortFile
   End If
   ' --- check for MUSFIP ---
   If Value = -1 Then
      MEM(MemPos_SortState) = 1 ' MUSFIP sorting
      Exit Sub
   End If
   ' --- check for invalid file number ---
   If Value < 1 Or Value > MaxFile Then
      MEM(MemPos_Status) = 6 ' invalid file number or file closed
      Exit Sub
   End If
   ' --- even pseudo-files are open ---
   If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then
      MEM(MemPos_Status) = 6 ' invalid file number or file closed
      Exit Sub
   End If
   ' --- create temporary sort file ---
   If MEM(MemPos_SortState) = 0 Then ' not sorting
      For lngLoop = 0 To MAXMEMTAGS
         SortTags(lngLoop) = ""
         SortIndex(lngLoop) = 0
      Next lngLoop
      SortFileName = "" ' not using a file yet
   Else ' MUSFIP sorting
      SortFileName = FileServerPath & "MUSFIP\" & AdjustFilenameWindows(KEY) & ".smp"
      SortFileNum = FreeFile
      Open SortFileName For Output As #SortFileNum
   End If
   ' --- done ---
   UNLOCKREC
   MEM(MemPos_SortState) = 2 ' sort initialized
   SortTagSize = 0
   SortLineCount = 0
   FetchLineCount = 0
   MEM(MemPos_Status) = 0 ' ok
End Sub

Public Sub SORTNUM(ByVal Size As Long, ByVal Value As Currency)
   Dim strTemp As String
   ' -------------------
   If Size < 1 Or Size > 6 Then
      ThrowError "SORTNUM", "Invalid size: " & Trim$(Str$(Size))
      Exit Sub
   End If
   If Value < 0 Then
      strTemp = "-" & FormatNum("z14", 100000000000000@ - Value)
   Else
      strTemp = "0" & FormatNum("z14", Value)
   End If
   AddSortTag strTemp
End Sub

Public Sub SORTALPHA(ByVal Size As Long, ByVal Value As String)
   Dim strTemp As String
   ' -------------------
   If Size < 1 Or Size > 256 Then
      ThrowError "SORTALPHA", "Invalid size: " & Trim$(Str$(Size))
      Exit Sub
   End If
   If Len(Value) = Size Then
      strTemp = Value
   Else
      strTemp = Left$(Value & Space(Size), Size)
   End If
   AddSortTag strTemp
End Sub

Public Sub MERGE()
   Dim lngPID As Long
   Dim lngPHnd As Long
   Dim strCommand As String
   Dim strResultFile As String
   ' -------------------------
   ' --- check if wrong sort state ---
   If MEM(MemPos_SortState) < 2 Or MEM(MemPos_SortState) > 3 Then
      CloseSortFile
      MEM(MemPos_SortState) = 0 ' not sorting
      N = 3 ' sort not initialized
      SYSERROR "MERGE ERROR"
      Exit Sub
   End If
   ' --- check if memory sort ---
   If SortFileName = "" Then
      MemorySort 1, SortLineCount
      GoTo Done
   End If
   ' --- close sort file ---
   Close #SortFileNum
   DoEvents
   ' --- change extension to ".srt" for sorted output file ---
   strResultFile = Left$(SortFileName, Len(SortFileName) - 4) & ".srt"
   ' --- sort the file ---
   ' --- "FSort" freeware command is case-sensitive when using "/cy" and is faster than "Sort" ---
   strCommand = BinPath & "FSORT.EXE"
   If Dir$(strCommand) = "" Then
      ThrowError "MERGE", "Cannot find sorting program FSORT.EXE"
      Exit Sub
   End If
   ' --- these command-line parameters are specific to FSORT.EXE ---
   strCommand = strCommand & " """ & SortFileName & """ """ & strResultFile & """ /cy /ln"
   On Error GoTo ErrorFound
   lngPID = Shell(strCommand, vbHide)
   If lngPID = 0 Then GoTo ErrorFound
   WaitForTerminate lngPID
   On Error GoTo 0
   ' --- get rid of old sort file ---
   On Error Resume Next
   Kill SortFileName
   On Error GoTo 0
   ' --- change the sort filename ---
   SortFileName = strResultFile
   If Dir$(SortFileName) = "" Then
      ThrowError "MERGE", "File not found: " & SortFileName
      Exit Sub
   End If
   SortFileNum = FreeFile
   Open SortFileName For Input As #SortFileNum
   ' --- done ---
Done:
   UNLOCKREC
   MEM(MemPos_SortState) = 4 ' merge done
   MEM(MemPos_Status) = 0 ' ok
   Exit Sub
ErrorFound:
   MEM(MemPos_SortState) = 0 ' not sorting
   N = 64 ' unknown error
   SYSERROR "MERGE ERROR"
End Sub

Public Function FETCH(ByVal Value As Long) As Boolean
   Dim strLine As String
   ' -------------------
   CheckDoEvents
   ' --- check if no records sorted ---
   If MEM(MemPos_SortState) = 2 Then
      Close #SortFileNum
      DoEvents
      SortFileNum = 0
      MEM(MemPos_SortState) = 6 ' eof
      MEM(MemPos_Status) = 0 ' all records fetched
      GoTo ErrorFound
   End If
   ' --- check if wrong sort state ---
   If MEM(MemPos_SortState) < 4 Or MEM(MemPos_SortState) > 5 Then
      CloseSortFile
      MEM(MemPos_SortState) = 0 ' not sorting
      MEM(MemPos_Status) = 63 ' unknown error
      SYSERROR "FETCH ERROR"
      GoTo ErrorFound
   End If
   ' --- update sort state ---
   MEM(MemPos_SortState) = 5 ' fetching
   ' --- get sort tag from memory ---
   If SortFileName = "" Then
      If FetchLineCount >= SortLineCount Then
         MEM(MemPos_SortState) = 6 ' eof
         MEM(MemPos_Status) = 0 ' all records fetched
         GoTo ErrorFound
      End If
      FetchLineCount = FetchLineCount + 1
      strLine = SortTags(SortIndex(FetchLineCount))
   Else
      ' --- get sort tag from file ---
      Do
         ' --- check if at end of file ---
         If EOF(SortFileNum) Then
            Close #SortFileNum
            DoEvents
            SortFileNum = 0
            MEM(MemPos_SortState) = 6 ' eof
            MEM(MemPos_Status) = 0 ' all records fetched
            GoTo ErrorFound
         End If
         ' --- read one line until not blank ---
         Line Input #SortFileNum, strLine
         FetchLineCount = FetchLineCount + 1
      Loop Until strLine <> ""
   End If
   ' --- get sort tag and record number ---
   REC = Val(Right$(strLine, 14)) ' split off rec
   strLine = Left$(strLine, Len(strLine) - 14) ' get actual sort tag
   ' --- store the sort tag ---
   If Value = -1 Then ' move to Z buffer
      INIT_Z
      LET_Z_A strLine
      INIT_Z
      INIT_IZ
   Else
      LET_C strLine
      ' --- check if file number is invalid ---
      If (Value < 0) Or (Value > MaxFile) Then
         MEM(MemPos_Status) = 63 ' other error
         GoTo ErrorFound
      End If
      ' --- check if file is open ---
      If Not MEMTF(MemPos_FileTable + (Value * FileEntrySize)) Then
         MEM(MemPos_Status) = 63 ' other error
         GoTo ErrorFound
      End If
      ' --- check if can read record ---
      If Not READREC(Value) Then
         MEM(MemPos_Status) = 5 ' record not found
         GoTo ErrorFound
      End If
   End If
   ' --- done ---
   MEM(MemPos_Status) = 0 ' ok
   FETCH = True
   Exit Function
ErrorFound:
   FETCH = False
End Function

Public Sub INITFETCH()
   ' --- check if not sorting ---
   If MEM(MemPos_SortState) < 4 And MEM(MemPos_SortState) <> 1 Then
      MEM(MemPos_Status) = 3 ' invalid sequence
      Exit Sub
   End If
   ' --- close any open sort file unless MUSFIP ---
   If MEM(MemPos_SortState) > 1 Then
      If SortFileName <> "" Then
         ' --- don't use CloseSortFile, as it deletes the file ---
         Close #SortFileNum
         DoEvents
         SortFileNum = 0
      End If
   End If
   ' --- build MUSFIP file name ---
   If MEM(MemPos_SortState) = 1 Then
      SortFileName = FileServerPath & "MUSFIP\" & AdjustFilenameWindows(KEY) & ".srt"
   End If
   ' --- open file for input ---
   If SortFileName <> "" Then
      If Dir$(SortFileName) = "" Then
         ThrowError "INITFETCH", "File not found: " & SortFileName
         Exit Sub
      End If
      SortFileNum = FreeFile
      Open SortFileName For Input As #SortFileNum
   End If
   ' --- done ---
   UNLOCKREC
   FetchLineCount = 0
   MEM(MemPos_SortState) = 4 ' merged file is ready
   MEM(MemPos_Status) = 0 ' ok
End Sub

' -----------------------------
' --- Conversion Statements ---
' -----------------------------

Public Function PACK(ByVal Size As Long, ByVal FromPos As Long, ByVal FromOfs As Long) As String
   Dim bChar As Byte
   Dim lngCount As Long
   Dim strResult As String
   Dim lngActualFrom As Long
   ' -----------------------
   If Size < 0 Or Size > 256 Then
      ThrowError "PACK", "Invalid size: " & Trim$(Str$(Size))
      Exit Function
   End If
   ' --- get actual offset values from buffer pointers ---
   Select Case FromPos
      Case MemPos_R: lngActualFrom = (MEM(MemPos_RP2) * 256) + MEM(MemPos_RP)
      Case MemPos_Z: lngActualFrom = (MEM(MemPos_ZP2) * 256) + MEM(MemPos_ZP)
      Case MemPos_X: lngActualFrom = (MEM(MemPos_XP2) * 256) + MEM(MemPos_XP)
      Case MemPos_Y: lngActualFrom = (MEM(MemPos_YP2) * 256) + MEM(MemPos_YP)
      Case MemPos_W: lngActualFrom = (MEM(MemPos_WP2) * 256) + MEM(MemPos_WP)
      Case MemPos_S: lngActualFrom = (MEM(MemPos_SP2) * 256) + MEM(MemPos_SP)
      Case MemPos_T: lngActualFrom = (MEM(MemPos_TP2) * 256) + MEM(MemPos_TP)
      Case MemPos_U: lngActualFrom = (MEM(MemPos_UP2) * 256) + MEM(MemPos_UP)
      Case MemPos_V: lngActualFrom = (MEM(MemPos_VP2) * 256) + MEM(MemPos_VP)
      Case Else: lngActualFrom = FromPos
   End Select
   ' --- move specified number of chars ---
   lngCount = 0 ' will be actual length of string
   strResult = ""
   Do While lngCount < Size
      bChar = MEM(lngActualFrom + FromOfs + lngCount)
      If bChar <> 0 Then ' skip nulls
         strResult = strResult & Chr$(bChar Mod 128)
      End If
      lngCount = lngCount + 1
      If bChar < 128 Then Exit Do
   Loop
   ' --- adjust any buffer pointers needed ---
   Select Case FromPos
      Case MemPos_R: SKIP_R FromOfs + lngCount
      Case MemPos_Z: SKIP_Z FromOfs + lngCount
      Case MemPos_X: SKIP_X FromOfs + lngCount
      Case MemPos_Y: SKIP_Y FromOfs + lngCount
      Case MemPos_W: SKIP_W FromOfs + lngCount
      Case MemPos_S: SKIP_S FromOfs + lngCount
      Case MemPos_T: SKIP_T FromOfs + lngCount
      Case MemPos_U: SKIP_U FromOfs + lngCount
      Case MemPos_V: SKIP_V FromOfs + lngCount
      ' --- no error if not a buffer ---
   End Select
   MEM(MemPos_Length) = 0
   PACK = RTrim$(strResult)
End Function

Public Function CONVERT(ByVal NumericFmt As String, ByVal FromPos As Long, ByVal FromOfs As Long) As Boolean
   Dim strValue As String
   Dim curResult As Currency
   ' -----------------------
   ' --- get value and move proper pointers ---
   strValue = PACK(LenFormat(NumericFmt), FromPos, FromOfs)
   ' --- check if value can be converted ---
   If Not ConvertNum(NumericFmt, strValue, curResult) Then
      CONVERT = False
      Exit Function
   End If
   NUMERIC_RESULT = curResult
   CONVERT = True
End Function

' -----------------------
' --- Misc Statements ---
' -----------------------

Public Sub NOP()
   ' --- does nothing ---
End Sub

Public Sub DELAY(ByVal Value As Currency)
   If Value > 0 Then
      ' --- Value is in 1/100 seconds ---
      Sleep Value * 10 ' change to milliseconds
   End If
End Sub

' -----------------------------------
' --- Terminal Control Statements ---
' -----------------------------------

Public Sub HOME()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "HOME"
End Sub

Public Sub RESETSCREEN()
   If MEMTF(MemPos_PrintOn) Then
      ' --- print to printer ---
      If CompressMultiFF Then
         If MEMTF(MemPos_PageHasData) Then
            LET_MEMTF MemPos_FFPending, True
            LET_MEMTF MemPos_PageHasData, False
         End If
      Else
         ' --- finish last line ---
         If MEMTF(MemPos_LineHasData) Then
            Print #PrinterFileNum,
         End If
         Print #PrinterFileNum, Chr$(12); ' Ctrl-L
         LET_MEMTF MemPos_PageHasData, False
         LET_MEMTF MemPos_LineHasData, False
      End If
   Else
      ' --- display to screen ---
      If MEMTF(MemPos_Background) Then Exit Sub ' not an error
      SendToServer "SCREEN" & vbTab & "RESET"
   End If
End Sub

Public Sub CLEARSCREEN()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "CLEAR"
End Sub

Public Sub CR()
   If MEMTF(MemPos_PrintOn) Then
      ' --- print to printer ---
      CheckFormFeed
      Print #PrinterFileNum, Chr$(13); ' Ctrl-M
      LET_MEMTF MemPos_PageHasData, True
      LET_MEMTF MemPos_LineHasData, True
   Else
      ' --- display to screen ---
      If MEMTF(MemPos_Background) Then Exit Sub ' not an error
      SendToServer "CURSOR" & vbTab & "CR"
   End If
End Sub

Public Sub CRD(ByVal Value As Long)
   Dim lngLoop As Long
   ' -----------------
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   For lngLoop = 1 To Value
      SendToServer "SCREEN" & vbTab & "ATT" & vbTab & "0"
      SendToServer "CURSOR" & vbTab & "NL"
      SendToServer "SCREEN" & vbTab & "ATT" & vbTab & "6"
   Next lngLoop
End Sub

Public Sub TABSET()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "TABSET"
End Sub

Public Sub TABCLEAR()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "TABCLEAR"
End Sub

Public Sub TABCANCEL()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "TABCANCEL"
End Sub

Public Sub KLOCK()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "KEYBOARD" & vbTab & "LOCK" & vbTab & "ON"
End Sub

Public Sub KFREE()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "KEYBOARD" & vbTab & "LOCK" & vbTab & "OFF"
End Sub

Public Sub CURSORAT(ByVal Row As Long, ByVal Column As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "AT" & vbTab & Trim$(Str$(Row)) & vbTab & Trim$(Str$(Column))
   MEM(MemPos_Char) = 255 ' implicit stay
End Sub

Public Sub MOVEUP(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "UP" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub MOVEDOWN(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "DOWN" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub MOVERIGHT(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "RIGHT" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub MOVELEFT(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "LEFT" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub TABCURSOR(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "CURSOR" & vbTab & "TAB" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub BELL(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SOUND" & vbTab & "BELL" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub LINEINSERT(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "LINEINSERT" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub LINEDELETE(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "LINEDELETE" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub CHARINSERT(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "CHARINSERT" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub CHARDELETE(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "CHARDELETE" & vbTab & Trim$(Str$(Value))
End Sub

' --------------------------------------------------
' --- Terminal Attribute and Graphics Statements ---
' --------------------------------------------------

Public Sub ATT(ByVal Value As Long)
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "ATT" & vbTab & Trim$(Str$(Value))
End Sub

Public Sub GRAPHON()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "GRAPH" & vbTab & "ON"
End Sub

Public Sub GRAPHOFF()
   If MEMTF(MemPos_Background) Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then Exit Sub
   SendToServer "SCREEN" & vbTab & "GRAPH" & vbTab & "OFF"
End Sub

Public Sub NL(ByVal Value As Long)
   Dim lngLoop As Long
   ' -----------------
   If Value <= 0 Then Exit Sub
   If MEMTF(MemPos_PrintOn) Then
      ' --- print to printer ---
      CheckFormFeed
      For lngLoop = 1 To Value
         Print #PrinterFileNum,
      Next
      LET_MEMTF MemPos_PageHasData, True
      LET_MEMTF MemPos_LineHasData, False
   Else
      ' --- display to screen ---
      If MEMTF(MemPos_Background) Then Exit Sub ' not an error
      SendToServer "CURSOR" & vbTab & "NL" & vbTab & Trim$(Str$(Value))
   End If
End Sub

Public Sub DISPLAYSPACE(ByVal Value As Long)
   Dim lngSaveChar As Long
   ' ---------------------
   If Value <= 0 Then Exit Sub
   lngSaveChar = CHAR ' CHAR value must be preserved
   DISPLAYSTRING Space(Value)
   LET_CHAR lngSaveChar
End Sub

Public Sub FF(ByVal Value As Long)
   ' --- ignore multiple Formfeeds and only do one ---
   RESETSCREEN
End Sub

Public Sub PRINTVT(ByVal Value As Long)
   ThrowError "PRINTVT", "Command not implemented in IDRIS"
End Sub

' --------------------------
' --- Text File Commands ---
' --------------------------

Public Function OPENTFA() As Boolean
   Dim lngLoop As Long
   Dim strPath As String
   ' -------------------
   ' --- check for bad parameters ---
   If (VOL < 0) Or (VOL > MaxVolume) Or (KEY = "") Or (Len(KEY) > 8) Then
      MEM(MemPos_Status) = 63 ' other error
      GoTo ErrorFound
   End If
   ' --- check specified volume ---
   If Not MEMTF(MemPos_VolTable + (VOL * VolEntrySize)) Then
      MEM(MemPos_Status) = 3 ' volume not open
      GoTo ErrorFound
   End If
   ' --- check if tfa exists ---
   strPath = FileServerPath & "DEVICE" & _
             FormatNum("z2", MEM(MemPos_VolTable + (VOL * VolEntrySize) + 1)) & _
             "\" & AdjustFilenameWindows(GetAlpha(MemPos_VolTable + (VOL * VolEntrySize) + 2)) & _
             "\" & AdjustFilenameWindows(KEY)
   If Dir$(strPath, vbDirectory) = "" Then
      MEM(MemPos_Status) = 2 ' tfa not found on specified volume
      GoTo ErrorFound
   End If
   ' --- see if tfa is already open ---
   For lngLoop = 0 To MaxTFA
      If MEMTF(MemPos_TFATable + (lngLoop * TFAEntrySize)) Then ' tfa is open
         If MEM(MemPos_TFATable + (lngLoop * TFAEntrySize) + 1) = VOL Then ' same volume
            If GetAlpha(MemPos_TFATable + (lngLoop * TFAEntrySize) + 2) = KEY Then ' same name
               MEM(MemPos_Status) = 0 ' ok
               LET_TFA lngLoop ' tfa number
               OPENTFA = True
               Exit Function
            End If
         End If
      End If
   Next
   ' --- find next empty slot ---
   For lngLoop = 0 To MaxTFA
      If Not MEMTF(MemPos_TFATable + (lngLoop * TFAEntrySize)) Then ' tfa is closed
         LET_MEMTF MemPos_TFATable + (lngLoop * TFAEntrySize), True ' tfa is open
         MEM(MemPos_TFATable + (lngLoop * TFAEntrySize) + 1) = VOL ' save volume number
         LetAlpha MemPos_TFATable + (lngLoop * TFAEntrySize) + 2, KEY ' save tfa name
         MEM(MemPos_Status) = 0 ' ok
         LET_TFA lngLoop ' tfa number
         OPENTFA = True
         Exit Function
      End If
   Next
   ' --- tfa table full ---
   MEM(MemPos_Status) = 1 ' tfa table full
ErrorFound:
   LET_TFA 255 ' invalid tfa
   OPENTFA = False
End Function

Public Function OPENCHANNEL(ByVal Value As Long) As Boolean
   Dim lngVol As Long
   Dim strFileName As String
   ' -----------------------
   ' --- check parameters ---
   If TFA < 0 Or TFA > MaxTFA Then
      If TFA <> 252 And TFA <> 253 Then
         MEM(MemPos_Status) = 1 ' tfa not open
         GoTo ErrorFound
      End If
   End If
   If TFA >= 0 And TFA <= MaxTFA Then
      If Not MEMTF(MemPos_TFATable + (TFA * TFAEntrySize)) Then
         MEM(MemPos_Status) = 1 ' tfa not open
         GoTo ErrorFound
      End If
   End If
   If Value < 0 Or Value > MaxChannel Then
      MEM(MemPos_Status) = 13 ' invalid channel number
      GoTo ErrorFound
   End If
   If MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then
      MEM(MemPos_Status) = 12 ' channel is already open
      GoTo ErrorFound
   End If
   ' --- get channel file name ---
   If TFA = 252 Or TFA = 253 Then
      strFileName = GetAlpha((256 * MEM(MemPos_ZP2)) + MEM(MemPos_ZP)) ' don't move buffer pointers
      If strFileName = "" Then
         MEM(MemPos_Status) = 16 ' illegal file name
         GoTo ErrorFound
      End If
      If Left$(strFileName, 1) = "/" Or Left$(strFileName, 1) = "\" Then
         strFileName = Mid$(strFileName, 2) ' remove leading "/"
      End If
      strFileName = Replace(strFileName, "/", "\") ' adjust unix to dos
      strFileName = FileServerPath & strFileName ' add full path
   Else
      If InvalidFilename(KEY) Then
         MEM(MemPos_Status) = 16 ' illegal file name
         GoTo ErrorFound
      End If
      lngVol = MEM(MemPos_TFATable + (TFA * TFAEntrySize) + 1) ' volume
      strFileName = FileServerPath & "DEVICE" & _
                    FormatNum("z2", MEM(MemPos_VolTable + (lngVol * VolEntrySize) + 1)) & _
                    "\" & AdjustFilenameWindows(GetAlpha(MemPos_VolTable + (lngVol * VolEntrySize) + 2)) & _
                    "\" & AdjustFilenameWindows(GetAlpha(MemPos_TFATable + (TFA * TFAEntrySize) + 2)) & _
                    "\" & AdjustFilenameWindows(KEY) ' create actual filename
   End If
   ' --- if LockFlag is true, acts like Create Channel ---
   If Not LockFlag Then
      If Dir$(strFileName) = "" Then
         MEM(MemPos_Status) = 2 ' file does not exist
         GoTo ErrorFound
      End If
   End If
   ' --- open/create channel file ---
   ChannelFileNums(Value) = FreeFile
   ChannelPaths(Value) = strFileName
   MEM(MemPos_Status) = 3 ' file exists and is in use
   On Error GoTo ErrorFound
   If LockFlag Then
      Open ChannelPaths(Value) For Binary Access Read Write Lock Read Write As #ChannelFileNums(Value)
   Else
      Open ChannelPaths(Value) For Binary Access Read As #ChannelFileNums(Value)
   End If
   On Error GoTo 0
   ' --- mark channel as open ---
   LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize), True
   MEM(MemPos_ChanTable + (Value * ChanEntrySize) + 1) = TFA
   If LockFlag Then
      LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize) + 2, True ' locked
      LockFlag = False
   Else
      LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize) + 2, False ' not locked
   End If
   MEM(MemPos_Status) = 0 ' ok
   OPENCHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   ChannelFileNums(Value) = 0
   ChannelPaths(Value) = ""
   UNLOCKREC
   OPENCHANNEL = False
End Function

Public Function CREATECHANNEL(ByVal Value As Long) As Boolean
   Dim lngVol As Long
   Dim strFileName As String
   ' -----------------------
   ' --- check parameters ---
   If TFA < 0 Or TFA > MaxTFA Then
      If TFA <> 252 And TFA <> 253 Then
         MEM(MemPos_Status) = 1 ' tfa not open
         GoTo ErrorFound
      End If
   End If
   If TFA >= 0 And TFA <= MaxTFA Then
      If Not MEMTF(MemPos_TFATable + (TFA * TFAEntrySize)) Then
         MEM(MemPos_Status) = 1 ' tfa not open
         GoTo ErrorFound
      End If
   End If
   If Value < 0 Or Value > MaxChannel Then
      MEM(MemPos_Status) = 13 ' invalid channel number
      GoTo ErrorFound
   End If
   If MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then
      MEM(MemPos_Status) = 12 ' channel is already open
      GoTo ErrorFound
   End If
   ' --- get channel file name ---
   If TFA = 252 Or TFA = 253 Then
      strFileName = GetAlpha((256 * MEM(MemPos_ZP2)) + MEM(MemPos_ZP)) ' don't move buffer pointers
      If strFileName = "" Then
         MEM(MemPos_Status) = 16 ' illegal file name
         GoTo ErrorFound
      End If
      If Left$(strFileName, 1) = "/" Or Left$(strFileName, 1) = "\" Then
         strFileName = Mid$(strFileName, 2) ' remove leading "/"
      End If
      strFileName = Replace(strFileName, "/", "\") ' adjust unix to dos
      strFileName = FileServerPath & strFileName ' add full path
   Else
      If InvalidFilename(KEY) Then
         MEM(MemPos_Status) = 16 ' illegal file name
         GoTo ErrorFound
      End If
      lngVol = MEM(MemPos_TFATable + (TFA * TFAEntrySize) + 1) ' volume
      strFileName = FileServerPath & "DEVICE" & _
                    FormatNum("z2", MEM(MemPos_VolTable + (lngVol * VolEntrySize) + 1)) & _
                    "\" & AdjustFilenameWindows(GetAlpha(MemPos_VolTable + (lngVol * VolEntrySize) + 2)) & _
                    "\" & AdjustFilenameWindows(GetAlpha(MemPos_TFATable + (TFA * TFAEntrySize) + 2)) & _
                    "\" & AdjustFilenameWindows(KEY) ' create actual filename
   End If
   ' --- open/create channel file ---
   ChannelFileNums(Value) = FreeFile
   ChannelPaths(Value) = strFileName
   MEM(MemPos_Status) = 0 ' ok if it works
   If Dir$(strFileName) <> "" Then
      MEM(MemPos_Status) = 4 ' file already exists
   End If
   On Error GoTo CannotOpen
   Open ChannelPaths(Value) For Binary Access Read Write Lock Read Write As #ChannelFileNums(Value)
   On Error GoTo 0
   ' --- mark channel as open and locked ---
   LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize), True
   MEM(MemPos_ChanTable + (Value * ChanEntrySize) + 1) = TFA
   LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize) + 2, True
   CREATECHANNEL = True
   Exit Function
CannotOpen:
   MEM(MemPos_Status) = 3 ' file exists and is in use
ErrorFound:
   On Error GoTo 0
   ChannelFileNums(Value) = 0
   ChannelPaths(Value) = ""
   CREATECHANNEL = False
End Function

Public Function DELETECHANNEL(ByVal Value As Long) As Boolean
   Dim strFileName As String
   ' -----------------------
   If Value < 0 Or Value > MaxChannel Then
      MEM(MemPos_Status) = 13 ' invalid channel number
      GoTo ErrorFound
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   strFileName = ChannelPaths(Value) ' save for delete
   CLOSECHANNEL Value
   MEM(MemPos_Status) = 7 ' unable to delete file
   On Error GoTo ErrorFound
   If Dir$(strFileName) = "" Then GoTo ErrorFound
   Kill strFileName ' delete the file
   On Error GoTo 0
   MEM(MemPos_Status) = 0 ' ok
   DELETECHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   DELETECHANNEL = False
End Function

Public Function RENAMECHANNEL(ByVal Value As Long) As Boolean
   ' --- renaming a channel will re-open the new channel, but the ---
   ' --- pointer gets set back to the beginning of the file.      ---
   Dim lngPos As Long
   Dim strFileName As String
   Dim strNewFileName As String
   ' --------------------------
   If Value < 0 Or Value > MaxChannel Then
      MEM(MemPos_Status) = 13 ' invalid channel number
      GoTo ErrorFound
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   If InvalidFilename(KEY) Then
      MEM(MemPos_Status) = 16 ' illegal file name
      GoTo ErrorFound
   End If
   strFileName = ChannelPaths(Value) ' save for rename
   lngPos = InStrRev(strFileName, "\")
   strNewFileName = Left$(strFileName, lngPos) & AdjustFilenameWindows(KEY)
   If Dir$(strNewFileName) <> "" Then
      MEM(MemPos_Status) = 4 ' new filename already exists
      GoTo ErrorFound
   End If
   ' --- save current location in file ---
   lngPos = Loc(ChannelFileNums(Value))
   ' --- close the channel so it can be renamed ---
   CLOSECHANNEL Value
   MEM(MemPos_Status) = 7 ' unable to rename file
   On Error GoTo ErrorFound
   Name strFileName As strNewFileName ' rename the file
   On Error GoTo 0
   MEM(MemPos_Status) = 3 ' file exists and is in use
   On Error GoTo ErrorFound
   ChannelFileNums(Value) = FreeFile
   ChannelPaths(Value) = strNewFileName
   Open ChannelPaths(Value) For Binary Access Read Write Lock Read Write As #ChannelFileNums(Value)
   Seek #ChannelFileNums(Value), lngPos ' restore current pointer
   On Error GoTo 0
   MEM(MemPos_Status) = 0 ' ok
   RENAMECHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   RENAMECHANNEL = False
End Function

Public Sub CLOSETFA()
   Dim lngLoop As Long
   ' -----------------
   ' --- check for bad parameters ---
   If TFA < 0 Or TFA > MaxTFA Then
      MEM(MemPos_Status) = 63 ' other error
      MEM(MemPos_TFA) = 255
      Exit Sub
   End If
   ' --- check if tfa already closed ---
   If Not MEMTF(MemPos_TFATable + (TFA * TFAEntrySize)) Then
      MEM(MemPos_Status) = 0 ' ok
      MEM(MemPos_TFA) = 255
      Exit Sub
   End If
   ' --- check if tfa in use ---
   For lngLoop = 0 To MaxChannel
      If MEMTF(MemPos_ChanTable + (lngLoop * ChanEntrySize)) Then
         If MEM(MemPos_ChanTable + (lngLoop * ChanEntrySize) + 1) = TFA Then
            MEM(MemPos_Status) = 10 ' resource in use
            Exit Sub
         End If
      End If
   Next
   ' --- close tfa ---
   LET_MEMTF MemPos_TFATable + (TFA * TFAEntrySize), False
   MEM(MemPos_Status) = 0 ' ok
   MEM(MemPos_TFA) = 255
End Sub

Public Sub CLOSECHANNEL(ByVal Value As Long)
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "CLOSECHANNEL", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Sub
   End If
   ' --- ok if already closed ---
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' already closed
      MEM(MemPos_Status) = 6 ' already closed
      Exit Sub
   End If
   ' --- close the current channel ---
   Close #ChannelFileNums(Value)
   DoEvents
   ChannelFileNums(Value) = 0
   ChannelPaths(Value) = ""
   LET_MEMTF MemPos_ChanTable + (Value * ChanEntrySize), False
   MEM(MemPos_Status) = 0 ' ok
End Sub

Public Function READCHANNEL(ByVal Value As Long, ByVal ToBuff As Long, ByVal UntilBuff As Long) As Boolean
   Dim bTemp As Byte
   Dim lngTemp As Long
   Dim ToOfs As Long
   Dim CharCount As Long
   Dim HasConversion As Boolean
   ' --------------------------
   CheckDoEvents
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "ReadChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- do conversion if TFA <> 252 ---
   HasConversion = (MEM(MemPos_ChanTable + (Value * ChanEntrySize) + 1) <> 252)
   ' --- prepare to get data ---
   InitBuffer ToBuff
   ' --- check length of data to read ---
   If UntilBuff < 0 Then
      CharCount = 255 ' maximum chars to read
   Else
      CharCount = LENGTH ' maximum chars to read
   End If
   MEM(MemPos_Length) = 0 ' number of chars read
   ToOfs = 0
NextChar:
   If CharCount <= 0 Then
      MEM(MemPos_Term) = 255 ' no ending found
      GoTo Done
   End If
   If EOF(ChannelFileNums(Value)) Then
      MEM(MemPos_Status) = 9 ' end of file
      MEM(MemPos_Term) = 255 ' no terminator
      GoTo ErrorFound
   End If
   ' --- get one character ---
   On Error GoTo ErrorFound
   ' ### can this be sped up by reading an entire buffer at once? ###
   Get #ChannelFileNums(Value), , bTemp
   On Error GoTo 0
   ' --- do conversion if needed ---
   If HasConversion Then
      If bTemp > 127 Then bTemp = bTemp - 128 ' move into 0-127 range
      If bTemp = 10 Then bTemp = 30 ' change to record separator
      If bTemp = 13 And Not EOF(ChannelFileNums(Value)) Then
         Get #ChannelFileNums(Value), , bTemp
         If bTemp > 127 Then bTemp = bTemp - 128 ' move into 0-127 range
         If bTemp <> 10 Then
            On Error GoTo ErrorFound
            Seek #ChannelFileNums(Value), ChannelFileNums(Value) - 1
            On Error GoTo 0
            bTemp = 13 ' return to original
         Else
            bTemp = 30 ' change to record separator
         End If
      End If
   End If
   ' --- check if have reached end of string ---
   If UntilBuff < 0 Then
      ' --- check standard terminator list ---
      If bTemp = 30 Then
         MEM(MemPos_Term) = bTemp
         GoTo Done
      End If
      If bTemp = 3 Or bTemp = 4 Or (bTemp >= 10 And bTemp <= 13) Then
         If HasConversion Then bTemp = bTemp + 128
         MEM(MemPos_Term) = bTemp
         GoTo Done
      End If
   Else
      ' --- check specified terminator list ---
      lngTemp = 0
      Do While (lngTemp <= 255) And (MEM(UntilBuff + lngTemp) <> 255)
         If bTemp = MEM(UntilBuff + lngTemp) Then
            MEM(MemPos_Term) = bTemp
            GoTo Done
         End If
         lngTemp = lngTemp + 1
      Loop
   End If
   ' --- convert character ---
   If HasConversion Then
      ' --- correct invalid characters here ---
      If bTemp < 32 Or bTemp > 126 Then
         bTemp = 32 ' change invalid chars to spaces
      End If
      bTemp = bTemp + 128 ' make unterminated
   End If
   ' --- store character ---
   MEM(ToBuff + ToOfs) = bTemp
   ToOfs = ToOfs + 1
   MEM(MemPos_Length) = LENGTH + 1
   If LENGTH = 255 Then GoTo Done
   CharCount = CharCount - 1
   GoTo NextChar
Done:
   On Error GoTo 0
   CheckDoEvents
   MEM(MemPos_Status) = 0 ' ok
   READCHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   CheckDoEvents
   READCHANNEL = False
End Function

Public Function WRITECHANNEL(ByVal Value As Long, ByVal FromBuff As Long) As Boolean
   Dim bTemp As Byte
   Dim FromOfs As Long
   Dim CharCount As Long
   Dim HasConversion As Boolean
   ' --------------------------
   CheckDoEvents
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "WriteChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- get character count ---
   CharCount = BufferPos(FromBuff) - FromBuff
   If CharCount < 1 Or CharCount > 255 Then
      ThrowError "WriteChannel", "Invalid number of characters: " & Trim$(Str$(CharCount))
      Exit Function
   End If
   ' --- do conversion if TFA <> 252 ---
   HasConversion = (MEM(MemPos_ChanTable + (Value * ChanEntrySize) + 1) <> 252)
   ' --- prepare for writing ---
   InitBuffer FromBuff
   FromOfs = 0
NextChar:
   If CharCount <= 0 Then GoTo Done
   bTemp = MEM(FromBuff + FromOfs)
   If HasConversion Then
      bTemp = ModPos(bTemp, 128) ' change to normal ascii
   End If
   MEM(MemPos_Status) = 63 ' unknown error
   On Error GoTo ErrorFound
   ' ### can this be sped up by writing an entire buffer at once? ###
   Put #ChannelFileNums(Value), , bTemp
   On Error GoTo 0
   FromOfs = FromOfs + 1
   CharCount = CharCount - 1
   GoTo NextChar
Done:
   MEM(MemPos_Status) = 63 ' unknown error
   On Error GoTo ErrorFound
   ' --- write out terminator ---
   If TERM <> 255 Then
      If TERM = 30 And HasConversion Then ' record separator
         bTemp = 13 ' cr
         Put #ChannelFileNums(Value), , bTemp
         bTemp = 10 ' lf
         Put #ChannelFileNums(Value), , bTemp
      Else
         bTemp = TERM
         If bTemp = 132 Then ' end of transmission
            bTemp = 138 ' change to linefeed
         End If
         If HasConversion Then
            bTemp = ModPos(bTemp, 128)
         End If
         Put #ChannelFileNums(Value), , bTemp
      End If
   End If
   ' --- Done ---
   On Error GoTo 0
   CheckDoEvents
   MEM(MemPos_Status) = 0 ' ok
   WRITECHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   CheckDoEvents
   WRITECHANNEL = False
End Function

Public Function BACKSPACECHANNEL(ByVal Value As Long, ByVal UntilBuff As Long) As Boolean
   Dim bTemp As Byte
   Dim CharCount As Long
   Dim lngTemp As Long
   Dim EOLCount As Long
   Dim HasConversion As Boolean
   ' --------------------------
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "BackspaceChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- do conversion if TFA <> 252 ---
   HasConversion = (MEM(MemPos_ChanTable + (Value * ChanEntrySize) + 1) <> 252)
   ' --- prepare for backspace ---
   On Error GoTo ErrorFound
   CharCount = LENGTH ' number of chars to skip
   MEM(MemPos_Length) = 0 ' number of chars skipped
NextChar:
   If CharCount <= 0 Then
      MEM(MemPos_Term) = 255 ' no ending found
      GoTo Done
   End If
   If Seek(ChannelFileNums(Value)) <= 1 Then
      GoTo BOFFound
   End If
   ' --- move back one character ---
   Seek #ChannelFileNums(Value), Seek(ChannelFileNums(Value)) - 1
   ' --- get character ---
   Get #ChannelFileNums(Value), , bTemp
   ' --- do conversion if needed ---
   If HasConversion Then
      If bTemp > 127 Then bTemp = bTemp - 128 ' move into 0-127 range
      If bTemp = 10 And Seek(ChannelFileNums(Value)) > 2 Then
         Seek #ChannelFileNums(Value), Seek(ChannelFileNums(Value)) - 2
         Get #ChannelFileNums(Value), , bTemp
         If bTemp > 127 Then bTemp = bTemp - 128 ' move into 0-127 range
         If bTemp <> 13 Then
            Get #ChannelFileNums(Value), , bTemp ' get original again
         Else
            bTemp = 30 ' record separator
         End If
      End If
      If bTemp = 10 Then bTemp = 30 ' unix linefeed to record separator
   End If
   ' --- don't check terminator on first char ---
   If LENGTH > 0 Then
      ' --- check if have found the end of a string ---
      If UntilBuff < 0 Then
         ' --- check standard terminator list ---
         If bTemp = 30 Then
            MEM(MemPos_Term) = bTemp
            GoTo Done
         End If
         If bTemp = 3 Or bTemp = 4 Or (bTemp >= 10 And bTemp <= 13) Then
            MEM(MemPos_Term) = bTemp + 128
            GoTo Done
         End If
      Else
         ' --- check specified terminator list ---
         lngTemp = 0
         Do While (lngTemp <= 255) And (MEM(UntilBuff + lngTemp) <> 255)
            If bTemp = MEM(UntilBuff + lngTemp) Then
               MEM(MemPos_Term) = bTemp
               GoTo Done
            End If
            lngTemp = lngTemp + 1
         Loop
      End If
   End If
   MEM(MemPos_Length) = LENGTH + 1
   CharCount = CharCount - 1
   ' --- move back before this character ---
   Seek #ChannelFileNums(Value), Seek(ChannelFileNums(Value)) - 1
   GoTo NextChar
Done:
   ' --- check if phantom linefeed needs gobbling ---
   If HasConversion And TERM = 30 And Not EOF(ChannelFileNums(Value)) Then
      Get #ChannelFileNums(Value), , bTemp ' check for linefeed
      If bTemp > 127 Then bTemp = bTemp - 128 ' move into 0-127 range
      If bTemp <> 10 Then ' undo gobble
         Seek #ChannelFileNums(Value), Seek(ChannelFileNums(Value)) - 1
      End If
   End If
   On Error GoTo 0
   MEM(MemPos_Status) = 0 ' ok
   BACKSPACECHANNEL = True
   Exit Function
BOFFound:
   On Error GoTo 0
   MEM(MemPos_Term) = 255 ' no terminator found
   MEM(MemPos_Status) = 9 ' beginning of file
   BACKSPACECHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   BACKSPACECHANNEL = False
End Function

Public Function REWINDCHANNEL(ByVal Value As Long) As Boolean
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "RewindChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- move pointer to beginning of file ---
   On Error GoTo ErrorFound
   Seek #ChannelFileNums(Value), 1
   On Error GoTo 0
   MEM(MemPos_Status) = 0 ' ok
   REWINDCHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   REWINDCHANNEL = False
End Function

Public Function WINDCHANNEL(ByVal Value As Long) As Boolean
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "WindChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- move pointer one past the end of the file ---
   On Error GoTo ErrorFound
   Seek #ChannelFileNums(Value), LOF(ChannelFileNums(Value)) + 1
   On Error GoTo 0
   MEM(MemPos_Status) = 0 ' ok
   WINDCHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   WINDCHANNEL = False
End Function

Public Function EOFCHANNEL(ByVal Value As Long) As Boolean
   Dim lngSize As Long
   ' -----------------
   ' --- check for invalid parameters ---
   If Value < 0 Or Value > MaxChannel Then
      ThrowError "EOFChannel", "Invalid Channel Number: " & Trim$(Str$(Value))
      Exit Function
   End If
   If Not MEMTF(MemPos_ChanTable + (Value * ChanEntrySize)) Then ' channel not open
      MEM(MemPos_Status) = 6 ' channel not open
      GoTo ErrorFound
   End If
   ' --- if already at EOF, then nothing to do. ---
   If EOF(ChannelFileNums(Value)) Then GoTo Done
   ' --- otherwise, resize the file ---
   On Error GoTo ErrorFound
   MEM(MemPos_Status) = 7 ' illegal operation
   ' --- save current point position ---
   lngSize = Seek(ChannelFileNums(Value))
   Close #ChannelFileNums(Value)
   DoEvents
   ' --- truncate the file one byte before pointer ---
   SetFileSize ChannelPaths(Value), lngSize - 1
   ' --- get new file number in case original is in use ---
   ChannelFileNums(Value) = FreeFile
   ' --- check if channel was originally open locked ---
   If MEMTF(MemPos_ChanTable + (Value * ChanEntrySize) + 2) Then
      Open ChannelPaths(Value) For Binary Access Read Write Lock Read Write As #ChannelFileNums(Value)
   Else
      Open ChannelPaths(Value) For Binary Access Read As #ChannelFileNums(Value)
   End If
   ' --- move pointer back ---
   Seek #ChannelFileNums(Value), lngSize
   On Error GoTo 0
Done:
   MEM(MemPos_Status) = 0 ' ok
   EOFCHANNEL = True
   Exit Function
ErrorFound:
   On Error GoTo 0
   EOFCHANNEL = False
End Function

' ------------------------------
' --- Special IDRIS Commands ---
' ------------------------------

Public Function GETLOGINID() As String
   If LoginID = "" Then
      ThrowError "GETLOGINID", "*** LoginID not known! ***"
      Exit Function
   End If
   GETLOGINID = UCase$(LoginID)
End Function

Public Function GETDATESTR() As String
   GETDATESTR = Format$(Date, "mm/dd/yyyy")
End Function

Public Function GETTIMESTR() As String
   GETTIMESTR = Format$(Now, "hh:mm:ss")
End Function

Public Function GETDATEVAL() As Currency
   GETDATEVAL = Val(Format$(Date, "yyyymmdd"))
End Function

Public Function GETTIMEVAL() As Currency
   GETTIMEVAL = Val(Format$(Now, "hhmm"))
End Function

Public Function GETTIMEHUND() As Currency
   GETTIMEHUND = Int(Timer * 100)
End Function

Public Function GETCOMPANYNAME() As String
   GETCOMPANYNAME = "CUSTOM DISABILITY SOLUTIONS"
End Function

Public Function GETCOMPANYNAMELEN() As Currency
   GETCOMPANYNAMELEN = Len(GETCOMPANYNAME)
End Function

Public Function GETCOMPANYINITIALS() As String
   GETCOMPANYINITIALS = "CDS"
End Function

Public Function GETCLIENTLIST() As String
   GETCLIENTLIST = ClientList
End Function

Public Function GETREADONLY() As Integer
   If ReadOnly Then
      GETREADONLY = TRUEVAL
   Else
      GETREADONLY = FALSEVAL
   End If
End Function

Public Function GETENVIRONMENT() As String
   GETENVIRONMENT = UCase$(Trim$(EnvName))
End Function

Public Sub SETSUBQUERY(ByVal Value As String)
   If SQLSubQuery <> "" Then
      SQLSubQuery = SQLSubQuery & " AND " & Value
   Else
      SQLSubQuery = Value
   End If
End Sub

Public Sub SETSUBQUERYFILE(ByVal Value As String)
   SQLSubQueryFile = Value
End Sub

Public Sub EXECSQL(ByVal Value As String)
   On Error GoTo SQLError
   If cnSQL Is Nothing Then GoTo ConnError
   If cnSQL.Errors.Count > 0 Then GoTo ConnError
   ' --- turn off timer before running stored procedure ---
   rtFormMain.SQLTimer.Enabled = False
   cnSQL.Execute Value, , adCmdText
   On Error GoTo 0
   ' --- turn timer back on ---
   rtFormMain.SQLTimer.Enabled = True
   Exit Sub
ConnError:
   ThrowError "EXECSQL", "SQL Connection Error:"
   GoTo ErrorFound
SQLError:
   ThrowError "EXECSQL", "Error executing SQL command:" & vbCrLf & Value
   Resume ErrorFound
ErrorFound:
   ' --- turn timer back on ---
   rtFormMain.SQLTimer.Enabled = True
End Sub

Public Sub SYSERROR(ByVal Msg As String)
   MUSTEXIT = True
   ' --- this is a trappable error ---
   If HasGosubStackItem(WHEN_ERROR_TYPEVAL) Then
      ClearGosubStackTo WHEN_ERROR_TYPEVAL
      Exit Sub
   End If
   ' --- no error handling, so fatal error ---
   FATALERROR Msg
End Sub

Public Sub FATALERROR(ByVal Msg As String)
   ' --- this is a non-trappable fatal error ---
   Dim strMsg As String
   ' ------------------
   MUSTEXIT = True
   EXITING = True
   DebugMessage "Exiting from FATALERROR"
   strMsg = Msg
   If Left$(strMsg, 3) <> "***" Then
      strMsg = "*** " & strMsg & " ***"
   End If
   ThrowError Mid$(App.EXEName, 5) & ":" & Format$(PROG, "000"), strMsg
End Sub

Public Sub EXITRUNTIME()
   ' --- must set flags before sending message ---
   MUSTEXIT = True
   EXITING = True
   DebugMessage "Exiting from EXITRUNTIME"
   ' --- send message that runtime is done ---
   If Not MEMTF(MemPos_Background) Then
      WAITTOEXIT = True
      SendToServer "APPLICATION" & vbTab & "END"
   End If
End Sub

' ----------------------
' --- Debug commands ---
' ----------------------

Public Sub DBUG(ByVal ILCode As String)
   ' ------------------------------------------------------------------
   ' --- This command is used to allow breaking, stepping, breakpoints,
   ' --- and watchpoints. When the IDRIS code is compiled with the
   ' --- option to include debugging code, each command is preceeded by
   ' --- 'DBUG "...ilcode..." : ', with whatever IL code represents the
   ' --- current command. This DBUG command is run before the actual VB
   ' --- line of code is run, so the breakpoint is before the command.
   ' --- During the BREAK period, the client sends any other commands
   ' --- to perform any other debugging features.
   ' ------------------------------------------------------------------
   Dim lngLoop As Long
   Dim oStackEntry As rtStackEntry
   ' -----------------------------
   If EXITING Then Exit Sub
   ' --- save the IL code for viewing, even if not BREAKing.    ---
   ' --- might be interrupting a running ENTER or EDIT command. ---
   LastILCode = ILCode
   ' --- update the current JumpPoint ---
   CurrJumpPoint = Val(Left$(ILCode, InStr(ILCode, " ") - 1))
   ' --- check if doing a one-step command ---
   If DebugOneStep Then
      DebugOneStep = False
      GoTo DoBreak
   End If
   ' --- check if at a breakpoint ---
   If Not (Breakpoints Is Nothing) Then
      If Breakpoints.Count > 0 Then
         For lngLoop = 1 To Breakpoints.Count
            Set oStackEntry = Breakpoints.Item(lngLoop)
            With oStackEntry
               If .DevNum = CurrDevNum And .VolName = CurrVolName And .LibName = CurrLibName Then
                  If .ProgNum = PROG Then
                     If .JumpNum = CurrJumpPoint Then
                        ' --- remove the breakpoint ---
                        Breakpoints.Remove lngLoop
                        GoTo DoBreak
                     End If
                  End If
               End If
            End With
         Next
      End If
   End If
   Exit Sub
DoBreak:
   ' --- break until "DEBUG GO" received ---
   BREAK
End Sub

Public Sub BREAK()
   ' ----------------------------------------------------------------------
   ' --- This command will pause the IDRIS program until the Client program
   ' --- sends a "DEBUG GO" message. It puts a full memory dump into a file
   ' --- that is reused each time. The client may send other commands back,
   ' --- such as "DEBUG ONESTEP", setting breakpoints, or altering variable
   ' --- values.
   ' ----------------------------------------------------------------------
   If EXITING Then Exit Sub
   ' --- send debug data to client program ---
   SendDebugData
   ' --- wait until exiting breakmode ---
   Do While InBreakMode And (Not EXITING)
      Sleep 1
      DoEvents
   Loop
End Sub

Public Sub SendDebugData()
   If EXITING Then Exit Sub
   ' --- check if common file path doesn't exist ---
   On Error Resume Next
   If Dir$(CommonFilePath, vbDirectory) = "" Then
      MkDir CommonFilePath
      Err.Clear
      On Error Resume Next
   End If
   If Dir$(CommonFilePath, vbDirectory) = "" And AltCommonFilePath <> "" Then
      CommonFilePath = AltCommonFilePath
      If Dir$(CommonFilePath, vbDirectory) = "" Then
         MkDir CommonFilePath
         Err.Clear
         On Error Resume Next
      End If
   End If
   If Dir$(CommonFilePath, vbDirectory) = "" Then
      ThrowError "BREAK", "CommonFilePath not found: " & CommonFilePath
      GoTo ErrorFound
   End If
   On Error GoTo ErrorFound
   ' --- store memory into temp file, using the same filename each time ---
   If BreakFilename = "" Then
      BreakFilename = GetTempFile(CommonFilePath, "MEM")
   End If
   SaveMemory BreakFilename
   InBreakMode = True
   If LastILCode = "" Then LastILCode = "BREAK"
   SendToServer "DEBUG" & vbTab & "BREAK" & vbTab & BreakFilename & vbTab & LastILCode
   Exit Sub
ErrorFound:
   Exit Sub
End Sub

' ------------------------------
' --- Custom Window commands ---
' ------------------------------

Public Sub WININIT(ByVal ClassName As String, ByVal WindowName As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WININIT" & vbTab & UCase$(ClassName) & vbTab & _
                UCase$(WindowName) & vbTab & cnSQL.ConnectionString
   CustomWindowProcessing
End Sub

Public Sub WINSHOW(ByVal WindowName As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINSHOW" & vbTab & UCase$(WindowName)
   CustomWindowProcessing
End Sub

Public Function WINSTATUS(ByVal WindowName As String) As String
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINSTATUS" & vbTab & UCase$(WindowName)
   CustomWindowProcessing
   WINSTATUS = CustomWindowResult
End Function

Public Sub WINSAVE(ByVal WindowName As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINSAVE" & vbTab & UCase$(WindowName)
   CustomWindowProcessing
End Sub

Public Sub WINSAVEALL()
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINSAVEALL"
   CustomWindowProcessing
End Sub

Public Sub WINVALIDATE(ByVal WindowName As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINVALIDATE" & vbTab & UCase$(WindowName)
   CustomWindowProcessing
End Sub

Public Sub WINVALIDATEALL()
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINVALIDATEALL"
   CustomWindowProcessing
End Sub

Public Function WINGETATTR(ByVal WindowName As String, ByVal AttributeName As String) As String
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINGETATTR" & vbTab & UCase$(WindowName) & vbTab & UCase$(AttributeName)
   CustomWindowProcessing
   WINGETATTR = CustomWindowResult
End Function

Public Sub WINSETATTR(ByVal WindowName As String, ByVal AttributeName As String, ByVal AttributeValue As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINSETATTR" & vbTab & UCase$(WindowName) & vbTab & UCase$(AttributeName) & vbTab & AttributeValue
   CustomWindowProcessing
End Sub

Public Sub WINUNLOAD(ByVal WindowName As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINUNLOAD" & vbTab & UCase$(WindowName)
   CustomWindowProcessing
End Sub

Public Sub WINUNLOADALL()
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINUNLOADALL"
   CustomWindowProcessing
End Sub

Public Sub WINEXEC(ByVal WindowName As String, ByVal Command As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINEXEC" & vbTab & UCase$(WindowName) & vbTab & Trim$(Command)
   CustomWindowProcessing
End Sub

Public Sub WINEXECALL(ByVal Command As String)
   CustomWindowProcessDone = False
   SendToServer "CUSTOMWINDOW" & vbTab & "WINEXECALL" & vbTab & Trim$(Command)
   CustomWindowProcessing
End Sub
