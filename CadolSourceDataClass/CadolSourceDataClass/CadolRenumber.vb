' ----------------------------------
' --- CadolRenumber - 04/09/2014 ---
' ----------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/09/2014 - SBakker
'            - Switched from being a Module to being a Public Class with Public Shared functions.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 10/19/2010 - SBakker
'            - Removed leading spaces from lines with line numbers. Somehow that messes up
'              the numbering, so there may be duplicates afterwards.
' 09/15/2010 - SBakker
'            - Added graceful handling of duplicate line number in a source file.
' 09/13/2010 - SBakker
'            - Fixed minor issues with turning on Option Strict.
' ----------------------------------------------------------------------------------------------------

Imports System.Text

Public Class CadolRenumber

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Public Shared Function RenumberCadolProgram(ByVal Program As String) As String

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        ' goto_line indicates whether a line contains line numbers:
        '    0 = unknown
        '    1 = line does not contain a line number
        '    2 = last number on line is a line number
        '    3 = all remaining numbers are line numbers (on-goto)

        ' check_token controls the search for more information about the line:
        '    0 = no search needed
        '    1 = if/when found - check until positive match for another command
        '    2 = then/else found - check next token only
        '    3 = on found - check until "go" or "goto" found
        '    4 = next token must be "channel" for there to be a line number
        '    5 = next token must be "release" for there to be a line number

        Dim Lines() As String
        Dim Result As New StringBuilder
        Dim TempPos As Integer
        Dim LineNum As Integer
        Dim CharNum As Integer
        Dim CurrChar As Char
        Dim LastChar As Char
        Dim ThisLine As String
        Dim OldLineNum As Integer
        Dim NewLineNum As Integer
        Dim Token As String
        Dim LastToken As String
        Dim TempToken As String
        Dim InComment As Boolean
        Dim InQuote As Boolean
        Dim GotoFlag As Integer
        Dim CheckFlag As Integer
        Dim QuoteChar As Char
        Dim IsLineNum As Boolean
        Dim SpaceCount As Integer
        Dim OldNew As New Generic.Dictionary(Of Integer, Integer)
        ' -------------------------------------------------------

        ' --- build a list of lines ---
        Lines = Program.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))

        ' --- get old line numbers, assign new line number ---
        NewLineNum = 0
        For LineNum = 0 To Lines.GetUpperBound(0) - 1 ' last line is extra blank
            ThisLine = Lines(LineNum).Trim
            If ThisLine = "" Then Continue For
            If ThisLine.Chars(0) >= "0"c AndAlso ThisLine.Chars(0) <= "9"c Then
                If Lines(LineNum).StartsWith(" ") Then
                    Lines(LineNum) = Lines(LineNum).Trim ' remove spaces before line numbers
                End If
                TempPos = ThisLine.IndexOf(" ")
                If TempPos < 0 Then
                    Token = ThisLine ' line only has line number
                Else
                    Token = ThisLine.Substring(0, TempPos)
                End If
                If Not IsNumeric(Token) Then
                    Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Non-numeric Line Number Label Found: " + Token)
                End If
                OldLineNum = CInt(Token)
                If OldLineNum < 0 OrElse OldLineNum >= 999999 Then
                    Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Invalid Line Number Found: " + OldLineNum.ToString)
                End If
                NewLineNum += 10
                Try
                    OldNew.Add(OldLineNum, NewLineNum)
                Catch ex As Exception
                    Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Duplicate line number found: " + OldLineNum.ToString)
                End Try
            End If
        Next

        ' --- change to new line numbers ---
        For LineNum = 0 To Lines.GetUpperBound(0) - 1 ' last line is extra blank

            ' --- skip blank lines ---
            If Lines(LineNum) = "" Then
                Result.Append(vbCrLf)
                Continue For
            End If

            ' --- initialize values ---
            Token = ""
            LastToken = ""
            IsLineNum = False
            SpaceCount = 0
            InComment = False
            InQuote = False
            GotoFlag = 0
            CheckFlag = 0
            QuoteChar = Nothing
            LastChar = Nothing

            For CharNum = 0 To Lines(LineNum).Length - 1

                ' --- get this character ---
                CurrChar = Lines(LineNum).Chars(CharNum)

                ' --- handle comments first ---
                If InComment Then
                    Result.Append(CurrChar)
                    Continue For
                End If

                ' --- handle quotes next ---
                If InQuote Then
                    If CurrChar = QuoteChar Then
                        InQuote = False
                    End If
                    Result.Append(CurrChar)
                    Continue For
                End If

                ' --- check for beginning of line comment ---
                If SpaceCount = CharNum AndAlso (CurrChar = "*"c OrElse CurrChar = "!"c OrElse CurrChar = "."c) Then
                    InComment = True
                    ' --- fill in saved spaces ---
                    If SpaceCount > 0 Then
                        Result.Append(Space(SpaceCount))
                        SpaceCount = 0
                    End If
                    ' --- output character ---
                    Result.Append(CurrChar)
                    Continue For
                End If

                ' --- check for a letter, number, or underline ---
                If Char.IsLetterOrDigit(CurrChar) OrElse CurrChar = "_"c Then
                    ' --- check for leading line numbers ---
                    If CharNum = 0 AndAlso Char.IsDigit(CurrChar) Then
                        IsLineNum = True
                    End If
                    ' --- add character to token ---
                    Token += CurrChar
                    Continue For
                End If

                ' --- token has ended, handle it now ---
                If Token <> "" Then
                    TempToken = Token.ToUpper
                    ' --- handle line number label ---
                    If IsLineNum Then
                        If Not IsNumeric(Token) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Non-numeric Line Number Label Found: " + Token)
                        End If
                        OldLineNum = CInt(Token)
                        If Not OldNew.ContainsKey(OldLineNum) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Line Number Label Unexpectedly Missing: " + Token)
                        End If
                        NewLineNum = OldNew(OldLineNum)
                        LastToken = NewLineNum.ToString
                        Token = ""
                        IsLineNum = False
                    ElseIf GotoFlag = 3 Then
                        ' --- handle ON GOTO line numbers ---
                        If Not IsNumeric(Token) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Non-numeric Line Number Found: " + Token)
                        End If
                        OldLineNum = CInt(Token)
                        If Not OldNew.ContainsKey(OldLineNum) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Line Number Unexpectedly Missing: " + Token)
                        End If
                        NewLineNum = OldNew(OldLineNum)
                        LastToken = NewLineNum.ToString
                        Token = ""
                    Else
                        ' --- check for ON-GOTO command ---
                        If CheckFlag = 3 AndAlso GotoFlag = 0 Then
                            If TempToken = "GO" OrElse TempToken = "GOTO" Then
                                GotoFlag = 2
                            End If
                        End If
                        ' --- check for multi-token commands that the 2nd token is CHANNEL ---
                        If CheckFlag = 4 AndAlso GotoFlag = 0 Then
                            If TempToken = "CHANNEL" Then
                                GotoFlag = 2 ' ending line number
                            Else
                                GotoFlag = 1 ' no line number
                            End If
                        End If
                        ' --- check for multi-token commands that the 2nd token is TERMINAL ---
                        If CheckFlag = 5 AndAlso GotoFlag = 0 Then
                            If TempToken = "TERMINAL" Then
                                GotoFlag = 2 ' ending line number
                            Else
                                GotoFlag = 1 ' no line number
                            End If
                        End If
                        ' --- check for normal conditions ---
                        If CheckFlag <= 2 AndAlso GotoFlag = 0 Then
                            ' --- commands that may have line numbers ---
                            If TempToken = "IF" Then CheckFlag = 1
                            If TempToken = "WHEN" Then CheckFlag = 1
                            If TempToken = "ON" Then CheckFlag = 3
                            If TempToken = "BACKSPACE" Then CheckFlag = 4
                            If TempToken = "RENAME" Then CheckFlag = 4
                            If TempToken = "RELEASE" Then CheckFlag = 5
                            ' --- commands that always have line numbers ---
                            If TempToken = "ASSIGN" Then GotoFlag = 2
                            If TempToken = "CONTROL" Then GotoFlag = 2
                            If TempToken = "CONVERT" Then GotoFlag = 2
                            If TempToken = "CREATE" Then GotoFlag = 2
                            If TempToken = "DELETE" Then GotoFlag = 2
                            If TempToken = "EOF" Then GotoFlag = 2
                            If TempToken = "FETCH" Then GotoFlag = 2
                            If TempToken = "GOS" Then GotoFlag = 2
                            If TempToken = "GO" Then GotoFlag = 2
                            If TempToken = "GOTO" Then GotoFlag = 2
                            If TempToken = "OPEN" Then GotoFlag = 2
                            If TempToken = "READ" Then GotoFlag = 2
                            If TempToken = "REWIND" Then GotoFlag = 2
                            If TempToken = "WIND" Then GotoFlag = 2
                            If TempToken = "WRITE" Then GotoFlag = 2
                            ' --- idris commands ---
                            If TempToken = "OPENSORTFILE" Then GotoFlag = 2
                        End If
                        ' --- check if after then/else ---
                        If CheckFlag = 2 AndAlso GotoFlag = 0 Then
                            GotoFlag = 1 ' no line number
                        End If
                        ' --- check for "THEN" or "ELSE" at start of line ---
                        If GotoFlag = 0 Then
                            If TempToken = "THEN" Then CheckFlag = 2
                            If TempToken = "ELSE" Then CheckFlag = 2
                        End If
                        ' --- posibilities are exhausted, must be no line number for this command ---
                        If CheckFlag = 0 AndAlso GotoFlag = 0 Then
                            GotoFlag = 1 ' no line number
                        End If
                    End If
                    ' --- check for "write back" - has no line number ---
                    If LastToken.ToUpper = "WRITE" And TempToken = "BACK" Then
                        GotoFlag = 0
                    End If
                    ' --- output last token ---
                    If LastToken <> "" Then
                        Result.Append(LastToken)
                        LastToken = ""
                    End If
                    ' --- fill in saved spaces ---
                    If SpaceCount > 0 Then
                        Result.Append(Space(SpaceCount))
                        SpaceCount = 0
                    End If
                    ' --- clear token ---
                    LastToken = Token
                    Token = ""
                End If

                ' --- check for end-of-line comment --- 
                If CurrChar = "!"c Then
                    InComment = True
                    ' --- check for ending line number ---
                    If GotoFlag = 2 Then
                        If Not IsNumeric(LastToken) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Non-numeric Line Number Found: " + LastToken)
                        End If
                        OldLineNum = CInt(LastToken)
                        If Not OldNew.ContainsKey(OldLineNum) Then
                            Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Line Number Unexpectedly Missing: " + LastToken)
                        End If
                        NewLineNum = OldNew(OldLineNum)
                        LastToken = NewLineNum.ToString
                        Result.Append(LastToken)
                        LastToken = ""
                    End If
                    ' --- output last token ---
                    If LastToken <> "" Then
                        Result.Append(LastToken)
                        LastToken = ""
                    End If
                    ' --- fill in saved spaces ---
                    If SpaceCount > 0 Then
                        Result.Append(Space(SpaceCount))
                        SpaceCount = 0
                    End If
                    ' --- output character ---
                    Result.Append(CurrChar)
                    Continue For
                End If

                ' --- check for spaces between tokens ---
                If CurrChar = " "c Then
                    SpaceCount += 1 ' don't output until later
                    Continue For
                End If

                ' --- output last token ---
                If LastToken <> "" Then
                    Result.Append(LastToken)
                    LastToken = ""
                End If
                ' --- fill in saved spaces ---
                If SpaceCount > 0 Then
                    Result.Append(Space(SpaceCount))
                    SpaceCount = 0
                End If
                ' --- add in this character ---
                Result.Append(CurrChar)

                ' --- check for quotes ---
                If CurrChar = """"c OrElse CurrChar = "%"c OrElse CurrChar = "$"c OrElse CurrChar = "'"c Then
                    InQuote = True
                    QuoteChar = CurrChar
                End If

            Next

            ' --- check for "write back" - has no line number ---
            If LastToken.ToUpper = "WRITE" And Token.ToUpper = "BACK" Then
                GotoFlag = 0
            End If

            ' --- check for ending line number ---
            If IsLineNum OrElse (Token <> "" AndAlso Not InComment AndAlso GotoFlag = 2) Then
                If Not IsNumeric(Token) Then
                    Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Non-numeric Line Number Found: " + Token)
                End If
                OldLineNum = CInt(Token)
                If Not OldNew.ContainsKey(OldLineNum) Then
                    Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Line Number Unexpectedly Missing: " + Token)
                End If
                NewLineNum = OldNew(OldLineNum)
                Token = NewLineNum.ToString
            End If

            ' --- output last token ---
            If LastToken <> "" Then
                Result.Append(LastToken)
                LastToken = ""
            End If

            ' --- check for final token ---
            If Token <> "" Then
                ' --- fill in saved spaces ---
                If SpaceCount > 0 Then
                    Result.Append(Space(SpaceCount))
                    SpaceCount = 0
                End If
                ' --- output new line number ---
                Result.Append(Token)
                Token = ""
            End If

            ' --- output line separator ---
            Result.Append(vbCrLf)

        Next

        Return Result.ToString

    End Function

End Class
