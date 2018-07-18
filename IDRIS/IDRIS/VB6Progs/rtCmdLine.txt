Attribute VB_Name = "rtCmdLine"
' ------------------------------
' --- rtCmdLine - 12/12/2006 ---
' ------------------------------

' ------------------------------------------------------------------------------
' 12/12/2006 - Removed unused variable "lngTemp".
' 10/15/2004 - Changed CommandSplitArgs to handle quotes inside strings.
' 09/25/2002 - Added CommandArgQuoted to return command arguments surrounded by
'              quotes if any spaces exist in the argument.
' 08/31/2001 - Added CommandShiftReset. Allows shifted parameters to be reset
'              back to their original positions.
'            - Added CommandArgFound and CommandArgValue functions. The first
'              will return True or False based on whether the command line arg
'              specified exists or not. The second will return the second half
'              of a command line arg in the format "/xxx=yyy".
' 08/24/2001 - Added CommandArgShift and mintArgOffset. This lets an optional
'              command-line parameter be processed, and then shifted off the
'              list so the rest of the parameters will be in their expected
'              locations.
'            - Subtract mintArgOffset from CommandNumArgs so it indicates the
'              highest argument number.
' 03/05/2001 - added parsing of parameters in single or double quotes.
' 11/07/2000 - added modCmdLine module.
' ------------------------------------------------------------------------------

Option Explicit

Private mblnSplitYet As Boolean
Private mintArgCount As Long
Private mintArgOffset As Long
Private mstrArgs() As String

Public Function CommandNumArgs() As Long
   CommandSplitArgs
   CommandNumArgs = mintArgCount - mintArgOffset
End Function

Public Function CommandArg(ByVal ArgNum As Long) As String
   CommandSplitArgs
   If ArgNum + mintArgOffset > 0 And ArgNum + mintArgOffset <= mintArgCount Then
      CommandArg = mstrArgs(ArgNum + mintArgOffset - 1)
   Else
      CommandArg = ""
   End If
End Function

Public Function CommandArgQuoted(ByVal ArgNum As Long) As String
   Dim strTemp As String
   ' -------------------
   strTemp = CommandArg(ArgNum)
   If InStr(strTemp, " ") > 0 Then strTemp = """" & strTemp & """"
   CommandArgQuoted = strTemp
End Function

Public Sub CommandArgShift()
   ' --- This will force CommandArg to shift the argument numbers back by one. ---
   ' --- This will make the first argument go from 1 to 0 to -1 to -2, etc.    ---
   ' --- The arguments are still accessible, if you know about the shifting.   ---
   ' --- Use CommandShiftReset to return the args to their original positions. ---
   CommandSplitArgs
   mintArgOffset = mintArgOffset + 1
End Sub

Public Sub CommandShiftReset()
   ' --- Reset the offset used for CommandArg. This will make all the shifted ---
   ' --- parameters available again. Sometimes the parameters need to be used ---
   ' --- more than once, such as when entering and when leaving a program.    ---
   CommandSplitArgs
   mintArgOffset = 0
End Sub

Public Function CommandArgFound(ByVal Arg As String) As Boolean
   ' --- This looks for an exactly matching argument on the command line. ---
   ' --- Used for finding flags to enables features, like "/DEBUG".       ---
   Dim lngLoop As Long
   ' --------------------
   CommandSplitArgs
   For lngLoop = 0 To mintArgCount - 1
      If UCase$(mstrArgs(lngLoop)) = UCase$(Arg) Then
         CommandArgFound = True
         Exit Function
      End If
   Next lngLoop
   CommandArgFound = False
End Function

Public Function CommandArgValue(ByVal Arg As String) As String
   ' --- This will return the second half of a command line parameter. ---
   ' --- The first half must match "ARG=", where ARG is the argument   ---
   ' --- sent to this function. The argument and parameters are all    ---
   ' --- uppercased before comparison and the return value is trimmed. ---
   Dim lngLoop As Long
   Dim strTempArg As String
   Dim lngTempLen As Long
   ' -----------------------
   CommandSplitArgs
   strTempArg = UCase$(Arg) + "="
   lngTempLen = Len(strTempArg)
   For lngLoop = 0 To mintArgCount - 1
      If Left$(UCase$(mstrArgs(lngLoop)), lngTempLen) = strTempArg Then
         CommandArgValue = Trim$(Mid$(mstrArgs(lngLoop), lngTempLen + 1))
         Exit Function
      End If
   Next lngLoop
   CommandArgValue = ""
End Function

Private Sub CommandSplitArgs()
   Dim strChar As String
   Dim strLine As String
   Dim lngLoop As Long
   Dim blnInArg As Boolean
   Dim blnInVal As Boolean
   Dim blnInQuote As Boolean
'  -------------------------
   If mblnSplitYet Then Exit Sub
   mblnSplitYet = True
   mintArgOffset = 0 ' not shifted
   mintArgCount = 0 ' none yet
   blnInArg = False
   blnInVal = False
   blnInQuote = False
   strLine = Command() ' get command line arguments
   For lngLoop = 1 To Len(strLine)
      strChar = Mid$(strLine, lngLoop, 1)
      If strChar = """" Then
         blnInQuote = Not blnInQuote ' does not store quote char
      ElseIf (strChar = " " Or strChar = vbTab) And (Not blnInQuote) Then
         blnInArg = False
      Else
         If Not blnInArg Then
            mintArgCount = mintArgCount + 1
            ReDim Preserve mstrArgs(mintArgCount - 1)
            blnInArg = True
         End If
         mstrArgs(mintArgCount - 1) = mstrArgs(mintArgCount - 1) & strChar
      End If
   Next lngLoop
End Sub
