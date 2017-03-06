' ----------------------------------
' --- CadolReformat - 04/09/2014 ---
' ----------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/09/2014 - SBakker
'            - Catch badly formatted Cadol programs which cause negative indents. Usually will mean a
'              missing "THEN" or an extra block terminator.
'            - Switched from being a Module to being a Public Class with Public Shared functions.
' 04/21/2011 - SBakker
'            - Make sure that indentation never goes below 0.
' 09/13/2010 - SBakker
'            - Fixed minor issues with turning on Option Strict.
'            - Add a "nop" command to lines with only a line number.
' ----------------------------------------------------------------------------------------------------

Imports System.Text

Public Class CadolReformat

    Public Shared Function ReformatCadolProgram(ByVal Program As String) As String

        Dim Result As New StringBuilder
        Dim TempPos As Integer
        Dim TempPos2 As Integer
        Dim Lines() As String
        Dim LineNum As Integer
        Dim IndentLevel As Integer
        Dim NextIndent As Integer
        Dim LabelValue As String
        Dim CommandValue As String
        Dim CommentValue As String
        Dim CharNum As Integer
        Dim CurrChar As Char
        Dim InQuote As Boolean
        Dim QuoteChar As Char
        Dim ThisLine As String
        Dim Changed As Boolean
        ' ---------------------------------

        ' --- build a list of lines ---
        Lines = Program.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))

        ' --- indent lines and align comments ---
        IndentLevel = 0
        NextIndent = 0
        Changed = False
        For LineNum = 0 To Lines.GetUpperBound(0) - 1 ' last line is extra blank
            ThisLine = Lines(LineNum).Trim
            IndentLevel = NextIndent
            If ThisLine = "" Then GoTo DoneLine
            If ThisLine.StartsWith("*") Then
                ThisLine = Lines(LineNum).TrimEnd ' keep leading spaces
                GoTo DoneLine
            End If
            If ThisLine.StartsWith(".") Then GoTo DoneLine
            LabelValue = ""
            CommandValue = ""
            CommentValue = ""
            If ThisLine.Chars(0) >= "0"c AndAlso ThisLine.Chars(0) <= "9"c Then
                TempPos = ThisLine.IndexOf(" ")
                If TempPos < 0 Then ' line number without a command
                    ' --- add a "nop" command ---
                    Lines(LineNum) = Lines(LineNum).TrimEnd + " nop"
                    ThisLine += " nop"
                    TempPos = ThisLine.IndexOf(" ")
                End If
                LabelValue = ThisLine.Substring(0, TempPos)
                ThisLine = ThisLine.Substring(TempPos + 1).TrimStart
            End If
            If ThisLine.IndexOf("!"c) >= 0 Then
                InQuote = False
                For CharNum = 0 To ThisLine.Length - 1
                    CurrChar = ThisLine.Chars(CharNum)
                    If CurrChar = """"c Or CurrChar = "'"c Or CurrChar = "%"c Or CurrChar = "$"c Then
                        If Not InQuote Then
                            InQuote = True
                            QuoteChar = CurrChar
                        ElseIf CurrChar = QuoteChar Then
                            InQuote = False
                        End If
                    End If
                    If Not InQuote AndAlso CurrChar = "!"c Then
                        CommentValue = ThisLine.Substring(CharNum)
                        ThisLine = ThisLine.Substring(0, CharNum).TrimEnd
                        Exit For
                    End If
                Next
            End If
            TempPos = ThisLine.IndexOf(" "c)
            If TempPos > 0 Then
                CommandValue = ThisLine.Substring(0, TempPos).Trim.ToUpper
            Else
                CommandValue = ThisLine.ToUpper
            End If
            Select Case CommandValue
                Case "FOR", "REPEAT", "WHILE"
                    NextIndent += 1
                Case "NEXT", "UNTIL", "DO"
                    IndentLevel -= 1
                    NextIndent -= 1
                    If IndentLevel < 0 Then
                        Throw New SystemException("Badly formatted Cadol program: LineNum = " + (LineNum + 1).ToString + ", Command = " + CommandValue)
                    End If
                Case "IF"
                    If ThisLine.ToUpper.EndsWith(" THEN") Then
                        NextIndent += 1
                    End If
                Case "THEN"
                    If ThisLine.ToUpper = "THEN" Then
                        NextIndent += 1
                    End If
                Case "ELSE"
                    If ThisLine.ToUpper = "ELSE" Then
                        IndentLevel -= 1
                        If IndentLevel < 0 Then
                            Throw New SystemException("Badly formatted Cadol program: LineNum = " + (LineNum + 1).ToString + ", Command = " + CommandValue)
                        End If
                    End If
                Case "ENDIF"
                    IndentLevel -= 1
                    NextIndent -= 1
                    If IndentLevel < 0 Then
                        Throw New SystemException("Badly formatted Cadol program: LineNum = " + (LineNum + 1).ToString + ", Command = " + CommandValue)
                    End If
                Case "END"
                    If ThisLine.ToUpper = "END IF" Then
                        IndentLevel -= 1
                        NextIndent -= 1
                        If IndentLevel < 0 Then
                            Throw New SystemException("Badly formatted Cadol program: LineNum = " + (LineNum + 1).ToString + ", Command = " + CommandValue)
                        End If
                    End If
            End Select
            TempPos2 = (IndentLevel * 3) - LabelValue.Length
            If TempPos2 <= 0 AndAlso LabelValue <> "" Then TempPos2 = 1
            If TempPos2 < 0 Then TempPos2 = 0
            ThisLine = LabelValue + Space(TempPos2) + ThisLine
            If CommentValue <> "" Then
                TempPos2 = ThisLine.Length + 1
                If TempPos2 < 32 Then TempPos2 = 32
                TempPos2 = ((TempPos2 + 7) \ 8) * 8
                ThisLine = ThisLine + Space(TempPos2 - ThisLine.Length) + CommentValue
            End If
DoneLine:
            If ThisLine <> Lines(LineNum) Then
                Changed = True
            End If
            Result.Append(ThisLine)
            Result.Append(vbCrLf)
        Next

        ' --- done ---
        Return Result.ToString

    End Function

End Class
