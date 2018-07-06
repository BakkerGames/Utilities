Attribute VB_Name = "rtIniFile"
' ------------------------------
' --- rtIniFile - 10/03/2005 ---
' ------------------------------

' -----------------------------------------------------------------------------
' 10/03/2005 - Changed "As Any" to proper data types.
' 02/17/2004 - Only add AppPath to EXEName when EXEName is not a full path.
' 08/24/2001 - Added FixAppPath to properly return it with a trailing "\".
' 05/30/2001 - Added "ByVal" to all parameters of GetINIString/SaveINIString.
'              Somehow got left out, and caused failures when building params.
' 05/22/2001 - Added optional ServerName parameter to GetIniName. It allows the
'              UseINI item to have different values based on the server name,
'              such as "UseIniDevel", "UseIniAccept", and "UseIniProd".
' 05/04/2001 - Only add extension to filename in GetIniName if one is not
'              specified. This allows files with other extensions to be used
'              as Ini files with these routines, and for the calling programs
'              to specify the full name of the file (ie "Connect.ini").
' 01/17/2001 - allow a name to be sent to the GetIniName routine. It will
'              be used as the basis for the Ini file name if not null. useful
'              if multiple apps are accessing the same Ini file.
' 09/20/2000 - added PathCase function to this module so it is self-contained.
' 05/30/2000 - added ability to redirect ini file using [Paths] UseINI=
'              trim strings before writing in SaveINIString.
'              trim strings before returning from GetINIString.
' -----------------------------------------------------------------------------

Option Explicit

' API Function declarations

Public Declare Function GetPrivateProfileString Lib "kernel32" _
                           Alias "GetPrivateProfileStringA" _
                          (ByVal lpApplicationName As String, _
                           ByVal lpKeyName As String, _
                           ByVal lpDefault As String, _
                           ByVal lpReturnedString As String, _
                           ByVal nSize As Long, _
                           ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" _
                           Alias "WritePrivateProfileStringA" _
                          (ByVal lpApplicationName As String, _
                           ByVal lpKeyName As String, _
                           ByVal lpString As String, _
                           ByVal lpFileName As String) As Long

Public Function GetINIString(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strDefault As String, _
                             ByVal strFileName As String) As String
   ' --- Retrieve a string from an INI file ---
   Dim strBuffer As String
   Dim lngResult As Long
'  ------------------------
   strBuffer = String$(255, 0) ' Initialize the receive buffer with nulls
   ' returns the length of the string in lngResult
   lngResult = GetPrivateProfileString(Trim$(strSection), _
                                       Trim$(strKey), _
                                       "", _
                                       strBuffer, _
                                       Len(strBuffer), _
                                       strFileName)
   ' If no value was found return the default value
   If Trim$(Left$(strBuffer, lngResult)) = "" Then
      GetINIString = strDefault
   Else
      GetINIString = Trim$(Left$(strBuffer, lngResult))
   End If
End Function

Public Function SaveINIString(ByVal strSection As String, _
                              ByVal strKey As String, _
                              ByVal strValue As String, _
                              ByVal strFileName As String) As Boolean
   ' --- Save a string to an INI file ---
   ' --- If strSection doesn't exist, it will be created automatically ---
   Dim lngResult As Long
'  ------------------------
   ' 0 = failed, nonzero = succeeded
   lngResult = WritePrivateProfileString(Trim$(strSection), _
                                         Trim$(strKey), _
                                         Trim$(strValue), _
                                         strFileName)
   SaveINIString = (lngResult <> 0)
End Function

Public Function GetININame(Optional ByVal EXEName As String = "", _
                           Optional ByVal ServerName As String = "") As String
   ' --- this will return the name of the ini file based on the ---
   ' --- current application name. it can also be redirected to ---
   ' --- another filename, such as when the original ini is     ---
   ' --- located on a network drive or a write-protected media. ---
   ' --- If an EXEName is sent, it will be instead of the name  ---
   ' --- of the current executable. Useful for multiple progs   ---
   ' --- accessing the same Ini file.
   Dim strTemp As String
   Dim strTemp2 As String
'  ------------------------
   EXEName = Trim$(EXEName)
   If EXEName = "" Then EXEName = App.EXEName ' use name of current app
   ' --- only add AppPath if EXEName is not a full path ---
   If Left$(EXEName, 2) = "\\" Or Mid$(EXEName, 2, 1) = ":" Then
      strTemp = PathCase(EXEName)
   Else
      strTemp = PathCase(FixAppPath) & PathCase(EXEName)
   End If
   If InStr(strTemp, ".") = 0 Then ' has no extention
      strTemp = strTemp & ".ini"
   End If
   ' --- check if the ini file really resides on another drive ---
   strTemp2 = GetINIString("Paths", "UseINI" & ServerName, "", strTemp)
   If strTemp2 <> "" Then strTemp = strTemp2
   ' --- return ini path and filename ---
   GetININame = strTemp
End Function

Private Function PathCase(ByVal S As String) As String
   Dim blnLastLetter As Boolean
   Dim strTemp As String
   Dim lngTemp As Long
   Dim chrTemp As String
'  ----------------------------
   strTemp = ""
   blnLastLetter = False
   For lngTemp = 1 To Len(S)
      chrTemp = UCase$(Mid$(S, lngTemp, 1))
      If blnLastLetter Then
         strTemp = strTemp & LCase$(chrTemp)
      Else
         strTemp = strTemp & chrTemp
      End If
      blnLastLetter = (chrTemp >= "A" And chrTemp <= "Z") Or (chrTemp = ".")
   Next lngTemp
   PathCase = strTemp
End Function

Private Function FixAppPath() As String
   If Right$(App.Path, 1) <> "\" Then
      FixAppPath = App.Path & "\"
   Else
      FixAppPath = App.Path
   End If
End Function
