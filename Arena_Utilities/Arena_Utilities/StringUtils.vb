' -----------------------------------
' --- StringUtils.vb - 07/22/2016 ---
' -----------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/22/2016 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrEmpty().
' 01/08/2016 - SBakker
'            - Added StringToSQLNull() to return a single-quoted string (ex. 'string') or NULL. It
'              uses StringToSQL first to double up any single quotes.
' 04/02/2015 - SBakker
'            - Added function StringToMixedCase() which uppercases the first letter of each word and
'              lowercases the rest of the word. Don't use on data that has been manually cased, such
'              as "ABC Company" and "John MacDonald".
' 03/26/2015 - SBakker
'            - Added more information to "String is too long and would be truncated" error.
' 03/16/2015 - SBakker
'            - Added StringToTSV() to return a string with tabs converted to spaces.
'            - Added optional SurroundingQuotes parameter to StringToCSV(), so it return the shortest
'              valid CSV-compatable string if False. Otherwise it will always surround with quotes.
'            - Added other known Unicode-to-ASCII conversions to StringToASCII().
'            - Added StringToANSI() to return a string with all non-ANSI (not 0-255) chars and invalid
'              ANSI chars replaced with "?" or space. This logic is separated so it can be called by
'              other StringTo??? functions. CR, LF, Tab are left alone, but other control chars are
'              changed to spaces.
'            - Added LeftZeroTrim() to trim leading zeroes from a string. Will always leave a single
'              zero before a decimal point and in the result if the string is only zeroes.
' 03/13/2015 - SBakker
'            - Added StringToCSV() for handling quotes and other chars properly when building a CSV
'              file. It returns the value already surrounded by double quotes.
' 03/12/2015 - SBakker
'            - Don't replace vbTab with Space in StringToSQL(). Tabs are sometimes necessary, and
'              removing them should be done in the calling program instead of here.
' 09/30/2014 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrWhiteSpace().
' 07/02/2014 - SBakker
'            - Fixed so that StringToCADOL() can also handle Nothing and properly return "".
' 05/12/2014 - SBakker
'            - Fixed so that StringToASCII() can handle Nothing and properly return "".
'            - Added new function StringToCADOL() to handle everything needed to save a string into a
'              CADOL alpha field - ASCII, ToUpper, and TrimEnd.
' 04/23/2014 - SBakker
'            - Added functions LeftPad(), RightPad(), and LeftZeroFill() for Long integers.
' 03/18/2014 - SBakker
'            - Added function LeftZeroFill() to add leading zeros to a string.
'            - Added Integer versions of LeftPad(), RightPad(), and LeftZeroFill() so the calling
'              program doesn't have to convert the value to a string itself.
' 03/12/2014 - SBakker
'            - Added LeftPad() to add leading spaces to a string. RightPad() adds trailing spaces.
' 08/02/2013 - SBakker - URD 12104
'            - Added function StringToASCII() for making sure IDRIS KEY fields are in ASCII.
' 07/19/2013 - SBakker
'            - Added EmptyStringToNothing() function, for easily converting "" to Nothing during save.
' 05/14/2012 - SBakker - Bug ArenaP2-293
'            - Change all non-CR/LF control characters (including Tabs) to single spaces in
'              StringToSQL. They all cause problems down the line. Also change any non-ASCII
'              Unicode characters to "?", as they can't be handled at this time.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 11/15/2010 - SBakker
'            - Added new functions ContainsIgnoreCase and ReplaceIgnoreCase. These are for
'              centralized handling of strings that might have case differences. The functs
'              String.Contains and String.Replace don't have CaseInsensitive options.
'            - Made sure that ReplaceIgnoreCase can't go into an infinite loop if the
'              ReplaceValue contains the SearchValue, by jumping past the replace position
'              for the next search.
' 03/29/2010 - SBakker
'            - Added new function HasWildcards().
' 01/25/2010 - SBakker
'            - Added new function WildcardsToSQL().
' 01/21/2010 - SBakker
'            - Moved WildcardChars to here so it can be used by all classes.
' ----------------------------------------------------------------------------------------------------

Imports System.Text

Public Class StringUtils

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Public Shared WildcardChars As Char() = {"?"c, "*"c, "_"c, "%"c}

    ''' <summary>
    ''' Returns a string value, suitable for sending to a SQL table.
    ''' </summary>
    Public Shared Function StringToSQL(ByVal Value As String) As String
        Dim Result As New StringBuilder
        ' -----------------------------
        ' --- Let StringToANSI() do most of the work ---
        For Each CurrChar As Char In StringToANSI(Value)
            Try
                If CurrChar = "'"c Then
                    ' --- Single quotes surround strings in SQL, change to two single quote chars ---
                    Result.Append("''")
                Else
                    ' --- Leave as is ---
                    Result.Append(CurrChar)
                End If
            Catch ex As Exception
                ' --- Change errors to "?" ---
                Result.Append("?")
            End Try
        Next
        Return Result.ToString
    End Function

    Public Shared Function StringToSQLNull(ByVal Value As String) As String
        If String.IsNullOrEmpty(Value) Then
            Return "NULL"
        End If
        Return "'" + StringToSQL(Value) + "'"
    End Function

    ''' <summary>
    ''' Returns a string surrounded by double-quotes, suitable for sending to a Comma Separated Value file.
    ''' </summary>
    Public Shared Function StringToCSV(ByVal Value As String) As String
        ' --- Always surround with double quotes ---
        Return StringToCSV(Value, True)
    End Function

    ''' <summary>
    ''' Returns a string optionally surrounded by double-quotes, suitable for sending to a Comma Separated Value file.
    ''' </summary>
    Public Shared Function StringToCSV(ByVal Value As String, ByVal SurroundingQuotes As Boolean) As String
        Dim HadDQuoteComma As Boolean = False
        Dim Result As New StringBuilder
        ' -----------------------------
        ' --- Let StringToANSI() do most of the work ---
        For Each CurrChar As Char In StringToANSI(Value)
            Try
                If CurrChar = """"c Then
                    ' --- Double quotes surround strings in CSV, change to two double quote chars ---
                    Result.Append("""""")
                    HadDQuoteComma = True
                ElseIf CurrChar = ","c Then
                    ' --- Output commaa but set flag ---
                    Result.Append(CurrChar)
                    HadDQuoteComma = True
                ElseIf CurrChar = vbCr OrElse CurrChar = vbLf OrElse CurrChar = vbTab Then
                    ' --- Change CR/LF/TAB to spaces ---
                    Result.Append(" ")
                Else
                    ' --- Leave as is ---
                    Result.Append(CurrChar)
                End If
            Catch ex As Exception
                ' --- Change errors to "?" ---
                Result.Append("?")
            End Try
        Next
        If SurroundingQuotes OrElse HadDQuoteComma OrElse Result.Length = 0 Then
            Return """" + Result.ToString + """"
        Else
            Return Result.ToString
        End If
    End Function

    ''' <summary>
    ''' Returns a string without tabs, suitable for sending to a Tab Separated Value file.
    ''' </summary>
    Public Shared Function StringToTSV(ByVal Value As String) As String
        Dim Result As New StringBuilder
        ' -----------------------------
        ' --- Let StringToANSI() do most of the work ---
        For Each CurrChar As Char In StringToANSI(Value)
            Try
                If CurrChar = vbCr OrElse CurrChar = vbLf OrElse CurrChar = vbTab Then
                    ' --- CR/LF/TAB need to be changed to spaces ---
                    Result.Append(" ")
                Else
                    ' --- Leave as is ---
                    Result.Append(CurrChar)
                End If
            Catch ex As Exception
                ' --- Change errors to "?" ---
                Result.Append("?")
            End Try
        Next
        Return Result.ToString
    End Function

    ''' <summary>
    ''' Returns a string with only ANSI characters, 0-255, no Unicode.
    ''' </summary>
    Public Shared Function StringToANSI(ByVal Value As String) As String
        Dim CurrAsc As Integer
        Dim Result As New StringBuilder
        ' -----------------------------
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        For Each CurrChar As Char In Value
            Try
                CurrAsc = Asc(CurrChar)
                If CurrAsc < 0 OrElse CurrAsc > 255 Then
                    ' --- Unicode characters can't be handled at this time ---
                    Result.Append("?")
                ElseIf CurrChar = vbCr OrElse CurrChar = vbLf OrElse CurrChar = vbTab Then
                    ' --- Leave CR/LF/TAB unchanged ---
                    Result.Append(CurrChar)
                ElseIf CurrAsc < 32 Then
                    ' --- All other control chars cause issues and need to be replaced here ---
                    Result.Append(" ")
                ElseIf CurrAsc = 173 Then
                    ' --- Soft Hyphen converts to Hyphen ---
                    Result.Append("-")
                ElseIf CurrAsc = 127 OrElse
                       CurrAsc = 129 OrElse
                       CurrAsc = 141 OrElse
                       CurrAsc = 143 OrElse
                       CurrAsc = 144 OrElse
                       CurrAsc = 157 OrElse
                       CurrAsc = 160 Then
                    ' --- All Windows-1252 unused chars and NBSP change to single space ---
                    Result.Append(" ")
                Else
                    ' --- Leave as is ---
                    Result.Append(CurrChar)
                End If
            Catch ex As Exception
                ' --- Change errors to "?" ---
                Result.Append("?")
            End Try
        Next
        Return Result.ToString
    End Function

    ''' <summary>
    ''' Add leading spaces to the specified long integer.
    ''' </summary>
    Public Shared Function LeftPad(ByVal Value As Long, ByVal NumChars As Integer) As String
        Return LeftPad(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add leading spaces to the specified integer.
    ''' </summary>
    Public Shared Function LeftPad(ByVal Value As Integer, ByVal NumChars As Integer) As String
        Return LeftPad(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add leading spaces to the specified string.
    ''' </summary>
    Public Shared Function LeftPad(ByVal Value As String, ByVal NumChars As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If NumChars < 1 Then
            Throw New SystemException(FuncName + vbCrLf + "Invalid length specified: " + NumChars.ToString)
        End If
        If Value.Length > NumChars Then
            Throw New SystemException(FuncName + vbCrLf + "String is too long and would be truncated: " + NumChars.ToString + ", """ + Value + """")
        End If
        If Value.Length = NumChars Then
            Return Value
        End If
        Return StrDup(NumChars - Value.Length, " "c) + Value
    End Function

    ''' <summary>
    ''' Add trailing spaces to the specified long integer.
    ''' </summary>
    Public Shared Function RightPad(ByVal Value As Long, ByVal NumChars As Integer) As String
        Return RightPad(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add trailing spaces to the specified integer.
    ''' </summary>
    Public Shared Function RightPad(ByVal Value As Integer, ByVal NumChars As Integer) As String
        Return RightPad(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add trailing spaces to the specified string.
    ''' </summary>
    Public Shared Function RightPad(ByVal Value As String, ByVal NumChars As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If NumChars < 1 Then
            Throw New SystemException(FuncName + vbCrLf + "Invalid length specified: " + NumChars.ToString)
        End If
        If Value.Length > NumChars Then
            Throw New SystemException(FuncName + vbCrLf + "String is too long and would be truncated: " + NumChars.ToString + ", """ + Value + """")
        End If
        If Value.Length = NumChars Then
            Return Value
        End If
        Return Value + StrDup(NumChars - Value.Length, " "c)
    End Function

    ''' <summary>
    ''' Add leading zeros to the specified long integer.
    ''' </summary>
    Public Shared Function LeftZeroFill(ByVal Value As Long, ByVal NumChars As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If Value < 0 Then
            Throw New SystemException(FuncName + vbCrLf + "Value cannot be negative: " + Value.ToString)
        End If
        Return LeftZeroFill(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add leading zeros to the specified integer.
    ''' </summary>
    Public Shared Function LeftZeroFill(ByVal Value As Integer, ByVal NumChars As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If Value < 0 Then
            Throw New SystemException(FuncName + vbCrLf + "Value cannot be negative: " + Value.ToString)
        End If
        Return LeftZeroFill(Value.ToString, NumChars)
    End Function

    ''' <summary>
    ''' Add leading zeros to the specified string.
    ''' </summary>
    Public Shared Function LeftZeroFill(ByVal Value As String, ByVal NumChars As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If NumChars < 1 Then
            Throw New SystemException(FuncName + vbCrLf + "Invalid length specified: " + NumChars.ToString)
        End If
        If Value.Length > NumChars Then
            Throw New SystemException(FuncName + vbCrLf + "String is too long and would be truncated: " + NumChars.ToString + ", """ + Value + """")
        End If
        If Value.Length = NumChars Then
            Return Value
        End If
        Return StrDup(NumChars - Value.Length, "0"c) + Value
    End Function

    Public Shared Function DigitsOnly(ByVal Value As String) As Boolean
        If String.IsNullOrEmpty(Value) Then
            Return False
        End If
        For CharNum As Integer = 0 To Value.Length - 1
            If Value.Chars(CharNum) < "0"c OrElse Value.Chars(CharNum) > "9"c Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Shared Function WildcardsToSQL(ByVal Value As String) As String
        If Value.IndexOf("*"c) >= 0 OrElse Value.IndexOf("?"c) >= 0 Then
            Return Value.Replace("*"c, "%"c).Replace("?"c, "_"c)
        Else
            Return Value
        End If
    End Function

    Public Shared Function HasWildcards(ByVal Value As String) As Boolean
        For Each CurrChar As Char In WildcardChars
            If Value.IndexOf(CurrChar) >= 0 Then Return True
        Next
        Return False
    End Function

    Public Shared Function ContainsIgnoreCase(ByVal Source As String, ByVal SearchValue As String) As Boolean
        Return Source.ToUpper.Contains(SearchValue.ToUpper)
    End Function

    Public Shared Function ReplaceIgnoreCase(ByVal Source As String, ByVal SearchValue As String, ByVal ReplaceValue As String) As String
        Dim Result As String = Source
        Dim StartPos As Integer = Result.ToUpper.IndexOf(SearchValue.ToUpper, 0)
        Do While StartPos >= 0
            Result = Left(Result, StartPos) + ReplaceValue + Right(Result, Result.Length - StartPos - SearchValue.Length)
            StartPos = Result.ToUpper.IndexOf(SearchValue.ToUpper, StartPos + ReplaceValue.Length)
        Loop
        Return Result
    End Function

    Public Shared Function EmptyStringToNothing(ByVal Value As String) As String
        If String.IsNullOrEmpty(Value) Then Return Nothing
        Return Value
    End Function

    ''' <summary>
    ''' Converts all accented letters to ASCII, or unknown chars to "?"
    ''' </summary>
    Public Shared Function StringToASCII(ByVal Value As String) As String
        Dim CurrAscW As Integer
        Dim Result As New StringBuilder
        ' -----------------------------
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        For Each CurrChar As Char In Value
            CurrAscW = AscW(CurrChar)
            If CurrChar = vbCr OrElse CurrChar = vbLf OrElse CurrChar = vbTab Then
                ' --- Change CR/LF/TAB to space ---
                CurrChar = " "c
            ElseIf CurrAscW < 32 Then
                ' --- Other control chars to "?" ---
                CurrChar = "?"c
            ElseIf CurrAscW = 173 Then
                ' --- Soft Hyphen converts to Hyphen ---
                CurrChar = "-"c
            ElseIf CurrAscW = 130 OrElse CurrAscW = 145 OrElse CurrAscW = 146 OrElse CurrAscW = 139 OrElse CurrAscW = 155 Then
                ' --- Curved and Angle Single Quotes ---
                CurrChar = "'"c
            ElseIf CurrAscW = 132 OrElse CurrAscW = 147 OrElse CurrAscW = 148 OrElse CurrAscW = 171 OrElse CurrAscW = 187 Then
                ' --- Curved and Angle Double Quotes ---
                CurrChar = """"c
            ElseIf CurrAscW = 160 Then
                ' --- Non-Break Space to Space ---
                CurrChar = " "c
            ElseIf CurrAscW >= 127 Then
                Select Case CurrChar
                    Case "À"c : CurrChar = "A"c
                    Case "Á"c : CurrChar = "A"c
                    Case "Â"c : CurrChar = "A"c
                    Case "Ã"c : CurrChar = "A"c
                    Case "Ä"c : CurrChar = "A"c
                    Case "Å"c : CurrChar = "A"c
                    Case "à"c : CurrChar = "a"c
                    Case "á"c : CurrChar = "a"c
                    Case "â"c : CurrChar = "a"c
                    Case "ã"c : CurrChar = "a"c
                    Case "ä"c : CurrChar = "a"c
                    Case "å"c : CurrChar = "a"c
                    Case "Ç"c : CurrChar = "C"c
                    Case "ç"c : CurrChar = "c"c
                    Case "È"c : CurrChar = "E"c
                    Case "É"c : CurrChar = "E"c
                    Case "Ê"c : CurrChar = "E"c
                    Case "Ë"c : CurrChar = "E"c
                    Case "è"c : CurrChar = "e"c
                    Case "é"c : CurrChar = "e"c
                    Case "ê"c : CurrChar = "e"c
                    Case "ë"c : CurrChar = "e"c
                    Case "Ì"c : CurrChar = "I"c
                    Case "Í"c : CurrChar = "I"c
                    Case "Î"c : CurrChar = "I"c
                    Case "Ï"c : CurrChar = "I"c
                    Case "ì"c : CurrChar = "i"c
                    Case "í"c : CurrChar = "i"c
                    Case "î"c : CurrChar = "i"c
                    Case "ï"c : CurrChar = "i"c
                    Case "Ñ"c : CurrChar = "N"c
                    Case "ñ"c : CurrChar = "n"c
                    Case "Ò"c : CurrChar = "O"c
                    Case "Ó"c : CurrChar = "O"c
                    Case "Ô"c : CurrChar = "O"c
                    Case "Õ"c : CurrChar = "O"c
                    Case "Ö"c : CurrChar = "O"c
                    Case "Ø"c : CurrChar = "O"c
                    Case "ò"c : CurrChar = "o"c
                    Case "ó"c : CurrChar = "o"c
                    Case "ô"c : CurrChar = "o"c
                    Case "õ"c : CurrChar = "o"c
                    Case "ö"c : CurrChar = "o"c
                    Case "ø"c : CurrChar = "o"c
                    Case "Š"c : CurrChar = "S"c
                    Case "š"c : CurrChar = "s"c
                    Case "Ù"c : CurrChar = "U"c
                    Case "Ú"c : CurrChar = "U"c
                    Case "Û"c : CurrChar = "U"c
                    Case "Ü"c : CurrChar = "U"c
                    Case "ù"c : CurrChar = "u"c
                    Case "ú"c : CurrChar = "u"c
                    Case "û"c : CurrChar = "u"c
                    Case "ü"c : CurrChar = "u"c
                    Case "Ÿ"c : CurrChar = "Y"c
                    Case "Ý"c : CurrChar = "Y"c
                    Case "ÿ"c : CurrChar = "y"c
                    Case "ý"c : CurrChar = "y"c
                    Case "Ž"c : CurrChar = "Z"c
                    Case "ž"c : CurrChar = "z"c
                    Case Else : CurrChar = "?"c
                End Select
            End If
            Result.Append(CurrChar)
        Next
        Return Result.ToString
    End Function

    ''' <summary>
    ''' Returns StringToASCII, ToUpper, and TrimEnd on Value
    ''' </summary>
    Public Shared Function StringToCADOL(ByVal Value As String) As String
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        Return StringToASCII(Value).ToUpper.TrimEnd
    End Function

    ''' <summary>
    ''' Trim leading zeros from the specified string. Does not trim leading spaces.
    ''' </summary>
    Public Shared Function LeftZeroTrim(ByVal Value As String) As String
        Dim CurrChar As Char
        Dim StartPos As Integer = 0
        ' -------------------------
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        For CurrIndex As Integer = 0 To Value.Length - 2 ' Don't check last char
            CurrChar = Value(CurrIndex)
            If CurrChar = "."c Then
                If StartPos > 0 Then
                    StartPos -= 1 ' Leave leading zero before decimal point
                End If
                Exit For
            ElseIf CurrChar = "0"c Then
                StartPos = CurrIndex + 1
            Else
                Exit For
            End If
        Next
        Return Value.Substring(StartPos)
    End Function

    ''' <summary>
    ''' Returns string with first letter of each word uppercase and the rest lowercase
    ''' </summary>
    Public Shared Function StringToMixedCase(ByVal Value As String) As String
        Dim InWord As Boolean = False
        Dim Result As New StringBuilder
        ' -----------------------------
        For Each CurrChar As Char In Value
            If Char.IsLetter(CurrChar) Then
                If Not InWord Then
                    InWord = True
                    Result.Append(Char.ToUpper(CurrChar))
                Else
                    Result.Append(Char.ToLower(CurrChar))
                End If
            Else
                InWord = False
                Result.Append(CurrChar)
            End If
        Next
        Return Result.ToString
    End Function

End Class
