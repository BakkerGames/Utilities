VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form rtFormMain 
   Caption         =   "IDRIS Quantum Runtime"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer SQLTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3240
      Top             =   1380
   End
   Begin VB.Timer TickTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2670
      Top             =   1380
   End
   Begin VB.Timer TimeoutTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2130
      Top             =   1380
   End
   Begin MSWinsockLib.Winsock wsToServer 
      Left            =   1590
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "rtFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -------------------------------
' --- rtFormMain - 02/17/2009 ---
' -------------------------------

Option Explicit

' ------------------------------------------------------------------------------
' 02/17/2009 - SBAKKER - 11235
'            - Added check to make sure the Client/Server is still connected
'              during the SQL "Ping" check.
' 09/17/2008 - SBAKKER - URD 11164
'            - Made changes recommended by CodeAdvisor.
' 06/27/2008 - SBAKKER - URD 11118
'            - Throw an error when a connection timeout occurs. This will allow
'              them to be logged and tracked, which wasn't happening before.
'            - Added command line info and error number to wsToServer_Error.
' 01/22/2007 - SBAKKER - URD 9739
'            - Added timer to check if the SQL connection has been broken, to
'              prevent hanging runtime processes.
' 01/30/2006 - Removed "DebugMessage '*** SWITCHING...'" messages. The problem
'              these were tracking has been corrected and is not needed anymore.
' ------------------------------------------------------------------------------

' --------------------------
' --- Internal variables ---
' --------------------------

Private blnParsingTick As Boolean

Private Sub Form_Load()
   rtFormMainLoaded = True
   DebugMessage "rtFormMain:Form_Load"
   ReadyToRun = False
   ServerSendComplete = True
   ' --- start timeout timer - must be before wsToServer.Connect ---
   If InsideIDE Or DebugFlag Then
      TimeoutTimer.Enabled = False
   Else
      TimeoutTimer.Interval = 30000 ' 30 seconds
      TimeoutTimer.Enabled = True
   End If
   TickTimer.Enabled = False
   blnParsingTick = False
   ' --- connect to server ---
   wsToServer.Connect HostIP, PortVal
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DebugMessage "rtFormMain:Form_Unload"
   TimeoutTimer.Enabled = False
   TickTimer.Enabled = False
   If wsToServer.State <> sckClosed Then
      On Error Resume Next
      DebugMessage "*** wsToServer.Close in Form_Unload"
      wsToServer.Close
      On Error GoTo 0
      DoEvents
   End If
   ' --- clear internal flags ---
   ReadyToRun = False
   MUSTEXIT = True
   EXITING = True
   DebugMessage "Exiting from Form_Unload"
   SWITCHING = False
   WAITTOEXIT = False
   rtFormMainLoaded = False
End Sub

' ------------------------
' --- WinSocket events ---
' ------------------------

Private Sub wsToServer_Connect()
   TimeoutTimer.Enabled = False ' connected
   ReadyToRun = True
   ' --- immediately send info for /WHO if known ---
   If LoginID <> "" Then
      SendToServer "SERVER" & vbTab & "LOGINID" & vbTab & UCase$(LoginID)
   End If
   If MachineName <> "" Then
      SendToServer "SERVER" & vbTab & "MACHINENAME" & vbTab & UCase$(MachineName)
   End If
   SendToServer "SERVER" & vbTab & "USERNUM" & vbTab & Trim$(Str$(USER))
   SendToServer "SERVER" & vbTab & "ORIGNUM" & vbTab & Trim$(Str$(ORIG))
   ' --- signal that everything is ready ---
   If Not MEMTF(MemPos_Background) Then
      SendToServer "KEYBOARD" & vbTab & "INFO" & vbTab & KBuff.ToString
      SendToServer "APPLICATION" & vbTab & "CONNECTED" & vbTab & GetLibNameFromApp
   End If
   DebugMessage "wsToServer_Connect is complete"
End Sub

Private Sub wsToServer_DataArrival(ByVal bytesTotal As Long)
   Dim strData As String
   ' -------------------
   ' --- give socket a chance to finish connecting ---
   If wsToServer.State = sckConnecting Then
      DoEvents
   End If
   ' --- check for invalid states ---
   DebugMessage "wsToServer_DataArrival, Bytes = " & Trim$(Str$(bytesTotal))
   If wsToServer.State = sckClosed Then Exit Sub
   If wsToServer.State = sckError Then Exit Sub
   ' --- get data from socket ---
   wsToServer.GetData strData, vbString
   PendingInput = PendingInput & strData
   TickTimer.Enabled = True
End Sub

Private Sub wsToServer_SendComplete()
   DebugMessage "wsToServer_SendComplete"
   ServerSendComplete = True
End Sub

Private Sub wsToServer_Close()
   DebugMessage "wsToServer_Close"
   If SWITCHING Then
      ThrowError "wsToServer_Close", "Closing while waiting for SWITCHREADY message"
      Exit Sub
   End If
   ServerSendComplete = True
   ' --- may need to close socket from this side ---
   If wsToServer.State = sckClosing Then
      DebugMessage "*** wsToServer.Close in wsToServer_Close"
      wsToServer.Close
      DoEvents
   End If
   ' --- exit application ---
   ReadyToRun = False
   MUSTEXIT = True
   EXITING = True
   DebugMessage "Exiting from wsToServer_Close"
   SWITCHING = False
   WAITTOEXIT = False
   ' --- force an escape ---
   LET_ESCVAL 0 ' perform normal escape
End Sub

Private Sub wsToServer_Error(ByVal Number As Integer, _
                             Description As String, _
                             ByVal Scode As Long, _
                             ByVal Source As String, _
                             ByVal HelpFile As String, _
                             ByVal HelpContext As Long, _
                             CancelDisplay As Boolean)
   DebugMessage "wsToServer_Error: " & Source & " - " & Description
   ServerSendComplete = True
   CancelDisplay = True
   TimeoutTimer.Enabled = False
   ' --- close socket ---
   If wsToServer.State <> sckClosed Then
      On Error Resume Next
      DebugMessage "*** wsToServer.Close in wsToServer_Error"
      wsToServer.Close
      On Error GoTo 0
      DoEvents
   End If
   ' --- check for timeout errors ---
   If Number = 10053 Then
      TimeoutTimer.Enabled = True
      ' --- re-connect to server ---
      On Error GoTo NoReconnect
      wsToServer.Connect HostIP, PortVal
      On Error GoTo 0
      DoEvents
      Exit Sub
   End If
   GoTo DoneTimeout
NoReconnect:
   Resume DoneTimeout
DoneTimeout:
   ' --- exit application ---
   ReadyToRun = False
   If Number <> 10061 Then ' server not running
      ThrowError Source, Trim$(Str$(Number)) & " - " & Description & " ***" & vbCrLf & "*** Parameters: " & Command()
   End If
   MUSTEXIT = True
   EXITING = True
   DebugMessage "Exiting from wsToServer_Error"
   SWITCHING = False
   WAITTOEXIT = False
   ' --- force an escape ---
   LET_ESCVAL 0 ' perform normal escape
End Sub

' --------------------
' --- Timer events ---
' --------------------

Private Sub TimeoutTimer_Timer()
   DebugMessage "TimeoutTimer_Timer"
   TimeoutTimer.Enabled = False
   ThrowError App.EXEName, "Timeout error connecting to client ***" & vbCrLf & "*** Parameters: " & Command()
End Sub

Private Sub TickTimer_Timer()
   Dim strData As String
   ' -------------------
   If blnParsingTick Then Exit Sub
   blnParsingTick = True
   TickTimer.Enabled = False
   ' --- parse any pending input messages ---
   If PendingInput <> "" Then
      ParseRuntimeCommand
   End If
   ' --- send any pending output messages to server ---
   If PendingOutput <> "" Then
      If DebugFlag And DebugFlagLevel > 0 Then
         DebugMessage "Has PendingOutput"
      End If
      If Not (wsToServer Is Nothing) Then
         If wsToServer.State = sckConnected Then
            strData = PendingOutput
            PendingOutput = ""
            ServerSendComplete = False
            If DebugFlag And DebugFlagLevel > 0 Then
               DebugMessage "SENT: " & strData
            End If
            wsToServer.SendData strData
            Do While (Not ServerSendComplete) And (wsToServer.State = sckConnected) And (Not EXITING)
               DoEvents
            Loop
         End If
      End If
   End If
   ' --- check if timer needs to be re-enabled ---
   TickTimer.Enabled = ((PendingInput <> "") Or (PendingOutput <> ""))
   blnParsingTick = False
   DoEvents
End Sub

Private Sub SQLTimer_Timer()
   Dim ErrorMsg As String
   On Error GoTo ErrorFound
   ' --- Check if still connected to SQL server ---
   ErrorMsg = "SQL connection is no longer available"
   If cnSQL Is Nothing Then GoTo ErrorFound
   If cnSQL.Errors.Count > 0 Then GoTo ErrorFound
   If cnSQL.State <> adStateOpen Then GoTo ErrorFound
   cnSQL.Execute "Ping"
   ' --- Check if still connected to client ---
   ErrorMsg = "Client connection is no longer available"
   If Not MEMTF(MemPos_Background) And Not (wsToServer Is Nothing) Then
      If wsToServer.State = sckClosed Then GoTo ErrorFound
      If wsToServer.State = sckError Then GoTo ErrorFound
   End If
   ' --- Ok ---
   On Error GoTo 0
   Exit Sub
ErrorFound:
   On Error Resume Next
   cnSQL.Close
   Set cnSQL = Nothing
   On Error GoTo 0
   ThrowError "SQLTimer", ErrorMsg
End Sub
