Attribute VB_Name = "rtMain"
' ---------------------------
' --- rtMain - 08/24/2009 ---
' ---------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 08/24/2009 - SBakker - 11336
'            - Changed the ConnectionTimeout to zero (infinite). Long-running
'              saves of complex policies were exceeding the default timeout.
' 02/17/2009 - SBAKKER - 11235
'            - Added check to make sure the Client/Server is still connected
'              during the SQL "Ping" check.
' 10/08/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
' 06/27/2008 - SBAKKER - URD 11118
'            - Added command line info to error messages to be able to diagnose
'              the "Client Port not specified" error which started appearing.
' 07/17/2007 - SBAKKER - URD 10969
'            - Create new User number if record is missing from [%USER] table.
'              This can happen if SQLAgent or SQLServer is restarted. Only
'              delete User number if LoginID matches.
' 02/08/2007 - SBAKKER - URD 10898
'            - Added UserAuthorized function to prevent users from accessing
'              IDRIS during EOM Calc time.
' 01/22/2007 - SBAKKER - URD 9739
'            - Turn on "ON ERROR RESUME NEXT" when closing down the library, to
'              prevent fatal errors when it is already trying to exit.
'            - Turn on timer for checking SQL Connection being broken.
' 01/18/2006 - Only check for "CommonFilePath Not Specified". Don't check if the
'              directory exists at this point.
'            - Added AltCommonFilePath variable to hold an alternate (and
'              hopefully existing) directory if the CommonFilePath isn't found.
' 10/07/2005 - Moved check for missing CommonFilePath after connection to Client
'              so that the error shows up on the client's end.
' 10/07/2005 - Wait for "APPLICATION END" message after an error thrown here.
' ------------------------------------------------------------------------------

' --- Command-line arguments:
' ---    /INI=ini filename
' ---    /ENV=environment
' ---    /START=devnum:volume:library:prognum:jumpnum
' ---    /MEM=memory filename
' ---    /SCF=devnum:volume
' ---    /BG
' ---    /DEBUG
' ---    /HOSTIP=ip value
' ---    /PORT=port value
' ---    /CLIENT=client list, 4 digits each, comma separated
' ---    /READONLY

Public Sub Main()
   Dim lngTemp As Long
   Dim strTemp As String
   Dim strTemp2 As String
   Dim strTemp3 As String
   Dim strErrMsg As String
   Dim strSQL As String
   Dim Tokens() As String
   Dim strStartProg As String
   Dim objItem As rtStackEntry
   Dim strMemFilename As String
   ' --------------------------
   Randomize
   ' --- initialize necessary variables ---
   On Error GoTo ErrorFound
   InitRuntime
   InitMemory
   SpawnTarget = ""
   rtFormMainLoaded = False
   rtDebugLogLoaded = False
   ' --- get INI filename ---
   IniFilename = CommandArgValue("/INI")
   If IniFilename = "" Then
      IniFilename = "CONNECT.INI"
      If Dir$(FixPath(App.Path) & IniFilename) <> "" Then
         IniFilename = FixPath(App.Path) & IniFilename
      ElseIf Dir$(FixPath(App.Path) & "..\..\BIN\" & IniFilename) <> "" Then
         IniFilename = FixPath(App.Path) & "..\..\BIN\" & IniFilename
      ElseIf Dir$(FixPath(App.Path) & "..\..\..\BIN\" & IniFilename) <> "" Then
         IniFilename = FixPath(App.Path) & "..\..\..\BIN\" & IniFilename
      End If
   End If
   If Mid$(IniFilename, 2, 1) <> ":" And Left$(IniFilename, 2) <> "\\" Then
      IniFilename = FixPath(App.Path) & IniFilename
   End If
   If Dir$(IniFilename) = "" Then
      strErrMsg = "INI File not found: " & IniFilename
      GoTo ErrorFound
   End If
   ' --- get environment ---
   EnvName = UCase$(CommandArgValue("/ENV"))
   If EnvName = "" Then
      EnvName = GetINIString(APP_NAME, "DefaultEnv", "", IniFilename)
   End If
   If EnvName = "" Then EnvName = "PROD"
   ' --- check if debugging ---
   DebugFlag = CommandArgFound("/DEBUG")
   If Not DebugFlag Then
      strTemp = GetINIString(APP_NAME & EnvName, "RuntimeDebugFlag", "", IniFilename)
      DebugFlag = (UCase$(strTemp) = "TRUE")
   End If
   strTemp = GetINIString(APP_NAME & EnvName, "RuntimeDebugLevel", "0", IniFilename)
   DebugFlagLevel = Val(strTemp)
   ' --- open debug log file ---
   If DebugFlag And DebugFlagLevel = 99 Then
      strTemp = GetTempFile("C:\TEMP", "LOG")
      strTemp = Replace(strTemp, ".tmp", ".txt")
      DebugLogFile = FreeFile
      Open strTemp For Output As #DebugLogFile
      DoEvents
      Kill Replace(strTemp, ".txt", ".tmp")
   End If
   ' --- show debug log screen ---
   If DebugFlag And DebugFlagLevel < 99 Then
      rtDebugLog.Show
   End If
   ' --- load necessary info from INI file ---
   BinPath = FixPath(GetINIString(APP_NAME & EnvName, "BinPath", "", IniFilename))
   FileServerPath = FixPath(GetINIString(APP_NAME & EnvName, "FileServerPath", "", IniFilename))
   LibraryPath = FixPath(GetINIString(APP_NAME & EnvName, "LibServerPath", "", IniFilename))
   TempPath = FixPath(GetINIString(APP_NAME & EnvName, "TempPath", "", IniFilename))
   UDLFilename = GetINIString(APP_NAME & EnvName, "UDL", "", IniFilename)
   ErrorFilename = GetINIString(APP_NAME & EnvName, "ErrorFile", "", IniFilename)
   CommonFilePath = GetINIString(APP_NAME & EnvName, "CommonFilePath", "", IniFilename)
   AltCommonFilePath = GetINIString(APP_NAME & EnvName, "AltCommonFilePath", "", IniFilename)
   ' --- validate INI file info ---
   strErrMsg = "BinPath not specified or not found"
   If BinPath = "" Then GoTo ErrorFound
   If Dir$(BinPath, vbDirectory) = "" Then GoTo ErrorFound
   strErrMsg = "FileServerPath not specified or not found"
   If FileServerPath = "" Then GoTo ErrorFound
   If Dir$(FileServerPath, vbDirectory) = "" Then GoTo ErrorFound
   strErrMsg = "LibraryPath not specified or not found"
   If LibraryPath = "" Then GoTo ErrorFound
   If Dir$(LibraryPath, vbDirectory) = "" Then GoTo ErrorFound
   strErrMsg = "TempPath not specified or not found"
   If TempPath = "" Then GoTo ErrorFound
   If Dir$(TempPath, vbDirectory) = "" Then GoTo ErrorFound
   strErrMsg = "TempPath cannot contain spaces"
   If InStr(TempPath, " ") > 0 Then GoTo ErrorFound
   strErrMsg = "UDLFilename not specified or not found"
   If UDLFilename = "" Then GoTo ErrorFound
   If Dir$(UDLFilename) = "" Then GoTo ErrorFound
   strErrMsg = "ErrorFilename not specified or not found"
   If ErrorFilename = "" Then GoTo ErrorFound
   ' --- get starting program ---
   strErrMsg = "/START parameter not specified"
   strStartProg = UCase$(CommandArgValue("/START"))
   If strStartProg = "" Then ' build out of exe path
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
               strStartProg = strTemp2 & ":" & strTemp & ":" & strTemp3 & ":0:0"
               DebugMessage "Built StartProg: " & strStartProg
            End If
         End If
      End If
   End If
   If strStartProg = "" Then GoTo ErrorFound
   DebugMessage "Starting Program = " & strStartProg
   ' --- add starting program to stack ---
   DebugMessage "Add Starting Program to Stack..."
   Set objItem = New rtStackEntry
   objItem.FromString strStartProg
   ' --- save starting info ---
   DebugMessage "Save Starting Info..."
   CurrDevNum = objItem.DevNum
   CurrVolName = objItem.VolName
   CurrLibName = objItem.LibName
   ' --- connect to SQL database ---
   DebugMessage "Connect to SQL Database..."
   Set cnSQL = New ADODB.Connection
   strErrMsg = "Error connecting to SQL"
   With cnSQL
      .ConnectionString = "File Name=" & UDLFilename
      .CommandTimeout = 0 ' infinite
      .ConnectionTimeout = 0 ' infinite
      .Open
   End With
   ' --- check if called from another runtime and already have a memory file ---
   strMemFilename = CommandArgValue("/MEM")
   If strMemFilename <> "" Then
      DebugMessage "Open Memory File """ & strMemFilename & """"
      strErrMsg = "Cannot load Memory file: " & strMemFilename
      If Not LoadMemory(strMemFilename) Then GoTo ErrorFound
      ' --- create file objects for "open" files ---
      DebugMessage "Create File Objects for Open Files..."
      If Not OpenAllFiles Then ' must be able to re-open all files
         strErrMsg = "Error Re-opening all files"
         GoTo ErrorFound
      End If
      ' --- open print file ---
      If PrinterFileName <> "" Then
         If Dir$(PrinterFileName) = "" Then
            strErrMsg = "File not found: " & PrinterFileName
            GoTo ErrorFound
         End If
         PrinterFileNum = FreeFile
         Open PrinterFileName For Append As #PrinterFileNum
      End If
      ' --- open sort file ---
      If SortFileName <> "" Then
         If MEM(MemPos_SortState) = 2 Or MEM(MemPos_SortState) = 3 Then
            If Dir$(SortFileName) = "" Then
               strErrMsg = "File not found: " & SortFileName
               GoTo ErrorFound
            End If
            SortFileNum = FreeFile
            Open SortFileName For Append As #SortFileNum
         End If
         If MEM(MemPos_SortState) = 4 Or MEM(MemPos_SortState) = 5 Then
            If Dir$(SortFileName) = "" Then
               strErrMsg = "File not found: " & SortFileName
               GoTo ErrorFound
            End If
            SortFileNum = FreeFile
            Open SortFileName For Input As #SortFileNum
         End If
         If MEM(MemPos_SortState) = 5 Then
            lngTemp = 0
            Do While lngTemp < FetchLineCount
               Line Input #SortFileNum, strTemp ' ignore previously read lines
               lngTemp = lngTemp + 1
            Loop
         End If
      End If
   Else
      ' --- need to get a user number and set original user ---
      strErrMsg = "Error getting User Number"
      lngTemp = GetUserNum
      If lngTemp < 1 Or lngTemp >= 32767 Then GoTo ErrorFound
      LetNumeric MemPos_User, 2, lngTemp
      LetNumeric MemPos_Orig, 2, lngTemp
      DebugMessage "User Number = " & Trim$(Str$(USER))
      ' --- open device 0, /sysvol, /scf ---
      DebugMessage "Open 0:/SYSVOL:/SCF..."
      strErrMsg = "Cannot open Device 0"
      If Not OPENDEVICE(0) Then
         If STATUS <> 2 Then GoTo ErrorFound ' 2 = already open
      End If
      LET_VOL 0
      LET_KEY "/SYSVOL"
      strErrMsg = "Cannot open /SYSVOL"
      If Not OPENVOLUME Then GoTo ErrorFound
      LET_KEY "/SCF"
      strErrMsg = "Cannot open /SCF on /SYSVOL"
      If Not OPENDATA(0) Then GoTo ErrorFound
      ' --- check if need to open another volume's /scf ---
      strTemp = CommandArgValue("/SCF")
      If strTemp <> "" Then
         DebugMessage "Open " & strTemp & ":/SCF..."
         strErrMsg = "Invalid /SCF parameter (not DEV:VOLUME)"
         If InStr(strTemp, ":") = 0 Then GoTo ErrorFound
         Tokens = Split(strTemp, ":")
         If UBound(Tokens) <> 1 Then GoTo ErrorFound ' must be "dev:volume"
         If Not IsNumeric(Tokens(0)) Then GoTo ErrorFound
         If Val(Tokens(0)) < 0 Or Val(Tokens(0)) > MaxDevice Then GoTo ErrorFound
         If Tokens(1) = "" Or Len(Tokens(1)) > 8 Then GoTo ErrorFound
         strErrMsg = "Cannot open Device " & Tokens(0)
         If Not OPENDEVICE(Val(Tokens(0))) Then
            If STATUS <> 2 Then GoTo ErrorFound ' 2 = already open
         End If
         LET_VOL Val(Tokens(0))
         LET_KEY Tokens(1)
         strErrMsg = "Cannot open " & Tokens(1)
         If Not OPENVOLUME Then GoTo ErrorFound
         LET_KEY "/SCF"
         strErrMsg = "Cannot open /SCF on " & Tokens(1)
         If Not OPENDATA(0) Then GoTo ErrorFound
      End If
   End If
   ' --- check if starting at a different point than Program 0 ---
   If objItem.ProgNum <> 0 Or objItem.JumpNum <> 0 Then
      GosubStack.Add objItem
   End If
   ' --- check if background process ---
   If CommandArgFound("/BG") Then
      LET_MEMTF MemPos_Background, True
      If USER = 0 Then
         LetNumeric MemPos_User, 2, GetUserNum
      End If
      If ORIG = 0 Then
         LetNumeric MemPos_Orig, 2, USER
      End If
      If LoginID = "" Then LoginID = "background"
      If MachineName = "" Then MachineName = "background"
      ' --- create background print file if none specified ---
      If PrinterFileName = "" Then
         AssignBackgroundPrintFile
      End If
   Else
      LET_MEMTF MemPos_Background, False
   End If
   ' -------------------------------------
   ' --- Communicate with Host program ---
   ' -------------------------------------
   DebugMessage "Communicate with Host Program..."
   ' --- find background connection info ---
   strErrMsg = "Server HostIP not specified"
   BackgroundIP = GetINIString(APP_NAME & EnvName, "HostIP", "", IniFilename)
   If BackgroundIP = "" Then GoTo ErrorFound
   strErrMsg = "Server BackgroundPort not specified"
   BackgroundPort = Val(GetINIString(APP_NAME & EnvName, "BackgroundPort", "", IniFilename))
   If BackgroundPort <= 0 Then GoTo ErrorFound
   ' --- find client/specified connection info ---
   HostIP = CommandArgValue("/HOSTIP")
   PortVal = Val(CommandArgValue("/PORT"))
   If Not MEMTF(MemPos_Background) Then
      strErrMsg = "Client HostIP not specified ***" & vbCrLf & "*** Parameters: " & Command()
      If HostIP = "" Then GoTo ErrorFound
      strErrMsg = "Client Port not specified ***" & vbCrLf & "*** Parameters: " & Command()
      If PortVal <= 0 Then GoTo ErrorFound
   Else
      If HostIP = "" Then HostIP = BackgroundIP
      If PortVal <= 0 Then PortVal = BackgroundPort
   End If
   ' --- only check ClientList and ReadOnly if they are not set yet ---
   If ClientList = "" Then
      ClientList = CommandArgValue("/CLIENT")
      If ClientList <> "" Then
         DebugMessage "ClientList = """ & ClientList & """"
         strErrMsg = "Error preparing Client List ***" & vbCrLf & "*** Parameters: " & Command()
         PrepareClientLists
      End If
   End If
   If Not ReadOnly Then
      ReadOnly = CommandArgFound("/READONLY")
   End If
   ' --- load the form with the sockets ---
   DebugMessage "Load rtFormMain..."
   Load rtFormMain
   ' --- wait until server handshaking is complete ---
   Do While (Not ReadyToRun) Or (LoginID = "") Or (MachineName = "")
      Sleep 10 ' prevents too much cpu usage
      DoEvents
      If EXITING Then
         GoTo Done ' ### not currently an error ###
      End If
   Loop
   ' --- have LoginID, validate that the user is authorized to connect ---
   If Not MEMTF(MemPos_Background) Then
      If Not UserAuthorized Then
         strErrMsg = GETLOGINID & " is not authorized to use IDRIS at this time"
         GoTo ErrorFound
      End If
   End If
   ' --- do non-critical validation after connected ---
   If Not MEMTF(MemPos_Background) Then
      ' --- can't do a slave print in the background ---
      strErrMsg = "CommonFilePath not specified: """ & CommonFilePath & """"
      If CommonFilePath = "" Then GoTo ErrorFound
   End If
   ' --- get default printer number once we have a LoginID ---
   strErrMsg = "Error getting Default Printer Number"
   LetByte MemPos_PrtNum, GetPRTNUM
   ' --- update user info when loginid is known ---
   strErrMsg = "Error updating User Info"
   If Not UpdateUserInfo Then
      ' --- need to get a new user number and set original user ---
      strErrMsg = "Error getting User Number"
      lngTemp = GetUserNum
      If lngTemp < 1 Or lngTemp >= 32767 Then GoTo ErrorFound
      LetNumeric MemPos_User, 2, lngTemp
      LetNumeric MemPos_Orig, 2, lngTemp
      DebugMessage "New User Number = " & Trim$(Str$(USER))
      strErrMsg = "Error updating User Info"
      If Not UpdateUserInfo Then GoTo ErrorFound
   End If
   ' --- enable sql disconnect testing timer ---
   rtFormMain.SQLTimer.Enabled = True
   ' --- run Cadol programs ---
   DebugMessage "Run Cadol Programs..."
   On Error GoTo 0
   ProgControl
   ' --- wait until switching/exitruntime handshaking is done ---
   DebugMessage "Before 'Do While SWITCHING Or WAITTOEXIT'..."
   Do While SWITCHING Or WAITTOEXIT
      Sleep 10 ' prevents too much cpu usage
      DoEvents
   Loop
   GoTo Done
   ' ---------------------
   ' --- Handle errors ---
   ' ---------------------
ErrorFound:
   On Error GoTo 0
   ThrowError APP_NAME & ":rtMain:Main", strErrMsg
   Do While WAITTOEXIT
      Sleep 10 ' prevents too much cpu usage
      DoEvents
   Loop
   GoTo Done
   ' -----------------------
   ' --- Exiting library ---
   ' -----------------------
Done:
   On Error Resume Next
   DebugMessage "Done, closing down..."
   ' --- close the form with the sockets ---
   If rtFormMainLoaded Then
      DebugMessage "Unloading rtFormMain..."
      Unload rtFormMain
   End If
   ' --- release user number if not spawning ---
   If Not (cnSQL Is Nothing) Then
      If cnSQL.State = adStateOpen And SpawnTarget = "" And USER <> 0 Then
         DebugMessage "Deleting User Number..."
         strSQL = "DELETE FROM [%USER] "
         strSQL = strSQL & "WHERE [USERNUM] = " & Trim$(Str$(USER)) & " "
         strSQL = strSQL & "AND [LOGINID] = '" & UCase$(LoginID) & "' "
         cnSQL.Execute strSQL
      End If
   End If
   ' --- close sql connection ---
   If Not (cnSQL Is Nothing) Then
      DebugMessage "Closing SQL Connection..."
      If cnSQL.State = adStateOpen Then
         cnSQL.Close
      End If
      Set cnSQL = Nothing
   End If
   ' --- close debug screen ---
   If DebugFlag Then
      If rtDebugLogLoaded Then
         DebugMessage "*** Library Exectution Complete ***"
         DebugMessage "<Close window to exit>"
         Do While rtDebugLog.Visible
            Sleep 20
            DoEvents
         Loop
         Unload rtDebugLog
      End If
      If DebugLogFile > 0 Then
         Close #DebugLogFile
         DebugLogFile = 0
      End If
   End If
End Sub
