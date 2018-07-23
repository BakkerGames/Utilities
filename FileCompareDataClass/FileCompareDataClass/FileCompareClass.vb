' ----------------------------------------
' --- FileCompareClass.vb - 07/23/2018 ---
' ----------------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/23/2018 - SBakker
'            - Make TrimBlanks also trim off trailing " _" for vb line continuation, not needed.
' 03/14/2014 - SBakker
'            - Use GetFileEncoding() to properly identify Binary files, and to get the right encoding
'              so File.ReadAllLines() will work correctly.
' 10/23/2013 - SBakker
'            - Switch to Arena versions of ConfigInfo, DataConn, and Utilities.
' 10/17/2013 - SBakker
'            - Added check for the two files being the same, except one has additional lines that the
'              other doesn't. Somehow this check had been missed!
'            - Fixed IgnoreLine() to have the proper squishing and trimming logic. It too was missed.
' 08/23/2011 - SBakker
'            - Added IgnoreSpaces for WordMode or TokenMode. This allows for comparison of
'              documents with many missing spaces between words, without cluttering up the
'              results with mostly those missing spaces.
' 06/28/2011 - SBakker
'            - Added new flavor of BinaryChar(integer), to go along with BinaryChar(char).
'              Made them both Public Shared.
' 06/24/2011 - SBakker
'            - Check for HTML codes &#8212; and &#8211; for em-dash and en-dash, as well as
'              &#8230; for elipsis, while ignoring.
' 02/07/2011 - SBakker
'            - Allow all control characters (1-31) be valid ASCII. Only CHR(0) and unused
'              characters are assumed to indicate a binary file.
'            - Added European left and right double quotes to those converted to (").
'            - Added all control characters to be excluded from tokens.
'            - Added InternationalLetters to included in words.
' 01/13/2011 - SBakker
'            - Added FirstDiffOnly for just finding out if files are different.
'            - Added check for binary, non-printable ASCII characters. Throw an error that
'              says the files are binary.
'            - Fixed missing blank line when only one difference at the end of the files.
'            - Added check to see if CurrLine is Nothing. File.ReadAllLines() can return a
'              Nothing line in some cases, usually at the end of the file.
' 01/07/2011 - SBakker
'            - Fixed to read each file's encoding before loading the lines, so there are no
'              issues with character interpretations.
'            - Added logic for ShowCarets and ShowAllCarets.
'            - Added routines Clear and ResetFlags so the same object can be used for
'              another compare. Clear is NOT necessary before another DoCompare, however.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FileUtils
Imports System.IO
Imports System.Text

''' <summary>
''' A data class which compares two files, returning the results in a text format.
''' </summary>
Public Class FileCompareClass

#Region " Private Constants "

    ''' <summary>
    ''' Used to reset the MaxDiff value.
    ''' </summary>
    Private Const MaxDiffDefault As Integer = 100

#End Region

#Region " Protected Variables "

    ''' <summary>
    ''' This allows an inherited class to know the name of the file currently being processed.
    ''' </summary>
    Protected CurrentFilename As String = ""

#End Region

#Region " Private Variables "

    Private Result As StringBuilder = Nothing
    Private List1 As List(Of LineItem)
    Private List2 As List(Of LineItem)

#End Region

#Region " Public Properties "

    ''' <summary>
    ''' Ignores case while comparing. Results are returned in all uppercase.
    ''' </summary>
    Public Property IgnoreCase As Boolean = False

    ''' <summary>
    ''' Ignores dashes, hyphens, emdashes, and endashes while comparing.
    ''' </summary>
    Public Property IgnoreDashes As Boolean = False

    ''' <summary>
    ''' Ignores all ellipsis, with and without internal spaces, while comparing.
    ''' </summary>
    Public Property IgnoreEllipsis As Boolean = False

    ''' <summary>
    ''' Ignores all single and double quotes while comparing.
    ''' </summary>
    Public Property IgnoreQuotes As Boolean = False

    ''' <summary>
    ''' Ignores spaces between words, but only for WordMode or TokenMode.
    ''' </summary>
    Public Property IgnoreSpaces As Boolean = False

    ''' <summary>
    ''' Converts all tabs to single spaces before comparison.
    ''' </summary>
    Public Property TabsToSpaces As Boolean = False

    ''' <summary>
    ''' Squishes all strings of more than one space into a single space before comparison.
    ''' </summary>
    Public Property SquishSpaces As Boolean = False

    ''' <summary>
    ''' Removes leading and trailing spaces before comparisons.
    ''' </summary>
    Public Property TrimBlanks As Boolean = False

    ''' <summary>
    ''' Removes blank lines during comparison. Original line numbers will be maintained, however.
    ''' </summary>
    Public Property SquishLines As Boolean = False

    ''' <summary>
    ''' Does not return any result if the files are identical.
    ''' </summary>
    Public Property QuietMode As Boolean = False

    ''' <summary>
    ''' Compares files as words comprised of alphabetic characters.
    ''' </summary>
    Public Property WordMode As Boolean = False

    ''' <summary>
    ''' Compares files as tokens comprised of non-whitespace characters.
    ''' </summary>
    Public Property TokenMode As Boolean = False

    ''' <summary>
    ''' For each pair of almost-matching lines, shows a caret under the first mismatched character.
    ''' </summary>
    Public Property ShowCarets As Boolean = False

    ''' <summary>
    ''' For each pair of almost-matching lines, shows carets under every mismatched character.
    ''' </summary>
    Public Property ShowAllCarets As Boolean = False

    ''' <summary>
    ''' Stop comparing once the first difference has been found.
    ''' </summary>
    Public Property FirstDiffOnly As Boolean = False

    ''' <summary>
    ''' Returns the number of differences after a compare has completed.
    ''' </summary>
    Public Property DiffCount As Integer = 0

    ''' <summary>
    ''' The maximum number of differences allowed before the files are "too different". May be adjusted before DoCompare().
    ''' </summary>
    Public Property MaxDiffs As Integer = MaxDiffDefault

#End Region

#Region " Public Routines "

    ''' <summary>
    ''' Clears any results. Not needed before DoCompare().
    ''' </summary>
    Public Sub Clear()
        DiffCount = 0
        Result = Nothing
    End Sub

    ''' <summary>
    ''' Returns result after a compare has completed.
    ''' </summary>
    Public Function Results() As String
        If Result Is Nothing Then
            Return ""
        End If
        Return Result.ToString
    End Function

    ''' <summary>
    ''' Compares two files and prepares the Results() for later use.
    ''' </summary>
    Public Sub DoCompare(ByVal Filename1 As String, ByVal Filename2 As String)
        Dim Pos1 As Integer
        Dim Pos2 As Integer
        Dim Found1 As Boolean
        Dim Found2 As Boolean
        Dim NewPos1 As Integer
        Dim NewPos2 As Integer
        Dim TempPos1 As Integer
        Dim TempPos2 As Integer
        Dim Offset As Integer
        ' ---------------------
        DiffCount = 0
        Result = New StringBuilder
        If Not File.Exists(Filename1) Then
            Result.Append("*** File not found: ")
            Result.Append(Filename1)
            Result.AppendLine(" ***")
            Exit Sub
        End If
        If Not File.Exists(Filename2) Then
            Result.Append("*** File not found: ")
            Result.Append(Filename2)
            Result.AppendLine(" ***")
            Exit Sub
        End If
        If Not QuietMode Then
            Result.Append("Comparing """)
            Result.Append(Filename1)
            Result.Append(""" and """)
            Result.Append(Filename2)
            Result.AppendLine("""")
            Result.AppendLine()
        End If
        List1 = New List(Of LineItem)
        List2 = New List(Of LineItem)
        Try
            FillLineList(Filename1, List1)
            FillLineList(Filename2, List2)
        Catch ex As Exception
            Result.AppendLine("*** Error loading files ***")
            Result.AppendLine()
            Result.AppendLine(ex.Message)
            Exit Sub
        End Try
        Pos1 = 0
        Pos2 = 0
        Do While Pos1 < List1.Count AndAlso Pos2 < List2.Count
            If List1.Item(Pos1).LineText = List2.Item(Pos2).LineText Then
                Pos1 += 1
                Pos2 += 1
                Continue Do
            End If
            If FirstDiffOnly Then ' Don't care about the results, just if they are different
                DiffCount = 1
                Exit Sub
            End If
            Found1 = False
            Found2 = False
            NewPos1 = Pos1 + MaxDiffs
            NewPos2 = Pos2 + MaxDiffs
            ' --- First check one side...
            For Offset = 0 To MaxDiffs - 1
                TempPos1 = Pos1 + Offset
                Do While TempPos1 < List1.Count AndAlso TempPos1 < Pos1 + MaxDiffs
                    TempPos2 = Pos2
                    Do While TempPos2 < List2.Count AndAlso TempPos2 < Pos2 + MaxDiffs
                        If List1(TempPos1).LineText = List2(TempPos2).LineText Then
                            Found1 = True
                            If (TempPos1 - Pos1) + (TempPos2 - Pos2) < (NewPos1 - Pos1) + (NewPos2 - Pos2) Then
                                NewPos1 = TempPos1
                                NewPos2 = TempPos2
                            End If
                            Exit Do
                        End If
                        TempPos2 += 1
                        If (TempPos1 - Pos1) + (TempPos2 - Pos2) >= (NewPos1 - Pos1) + (NewPos2 - Pos2) Then Exit Do
                    Loop
                    TempPos1 += 1
                    If (TempPos1 - Pos1) + (TempPos2 - Pos2) >= (NewPos1 - Pos1) + (NewPos2 - Pos2) Then Exit Do
                Loop
            Next
            ' --- ...then t'other
            For Offset = 0 To MaxDiffs - 1
                TempPos2 = Pos2 + Offset
                Do While TempPos2 < List2.Count AndAlso TempPos2 < Pos2 + MaxDiffs
                    TempPos1 = Pos1
                    Do While TempPos1 < List1.Count AndAlso TempPos1 < Pos1 + MaxDiffs
                        If List1(TempPos1).LineText = List2(TempPos2).LineText Then
                            Found2 = True
                            If (TempPos1 - Pos1) + (TempPos2 - Pos2) < (NewPos1 - Pos1) + (NewPos2 - Pos2) Then
                                NewPos1 = TempPos1
                                NewPos2 = TempPos2
                            End If
                            Exit Do
                        End If
                        TempPos1 += 1
                        If (TempPos1 - Pos1) + (TempPos2 - Pos2) >= (NewPos1 - Pos1) + (NewPos2 - Pos2) Then Exit Do
                    Loop
                    TempPos2 += 1
                    If (TempPos1 - Pos1) + (TempPos2 - Pos2) >= (NewPos1 - Pos1) + (NewPos2 - Pos2) Then Exit Do
                Loop
            Next
            ' --- Check if matching lines not found ---
            If Not Found1 AndAlso Not Found2 Then
                If Pos1 + MaxDiffs >= List1.Count AndAlso Pos2 + MaxDiffs >= List2.Count Then
                    AddResults(List1, 1, Pos1, List1.Count - 1)
                    AddResults(List2, 2, Pos2, List2.Count - 1)
                    Result.AppendLine()
                    DiffCount += 1
                    Pos1 = List1.Count
                    Pos2 = List2.Count
                    Exit Do
                End If
                Result.AppendLine("*** Files are very different ***")
                Exit Do
            End If
            ' --- Add shortest result found ---
            If (ShowCarets OrElse ShowAllCarets) AndAlso (NewPos1 - Pos1 = NewPos2 - Pos2) Then
                ' --- Same number of lines, so display each pair of lines with caret(s) showing the difference(s) ---
                For Offset = 0 To NewPos1 - Pos1 - 1
                    AddResults(List1, 1, Pos1 + Offset, Pos1 + Offset)
                    AddResults(List2, 2, Pos2 + Offset, Pos2 + Offset)
                    Dim FirstCaret As Boolean = False
                    Result.Append(Space(10)) ' must match length of line header: "1> 12345: "
                    For CharNum As Integer = 0 To Math.Min(List1(Pos1 + Offset).LineText.Length, List2(Pos2 + Offset).LineText.Length) - 1
                        If List1(Pos1 + Offset).LineText(CharNum) <> List2(Pos2 + Offset).LineText(CharNum) Then
                            Result.Append("^"c)
                            FirstCaret = True
                            If Not ShowAllCarets Then Exit For
                        Else
                            Result.Append(" "c)
                        End If
                    Next
                    If List1(Pos1 + Offset).LineText.Length <> List2(Pos2 + Offset).LineText.Length Then
                        If ShowAllCarets Then
                            Result.Append(StrDup(Math.Abs(List1(Pos1 + Offset).LineText.Length - List2(Pos2 + Offset).LineText.Length), "^"c))
                        ElseIf Not FirstCaret Then
                            Result.Append("^"c) ' lines match up to this point
                        End If
                    End If
                    Result.AppendLine() ' finish line of carets
                    Result.AppendLine()
                    DiffCount += 1
                Next
            ElseIf (WordMode OrElse TokenMode) AndAlso IgnoreSpaces Then
                ' --- Check if lines are truly different after removing spaces ---
                If GetResults(List1, 1, Pos1, NewPos1 - 1).Replace(" ", "") <> GetResults(List2, 2, Pos2, NewPos2 - 1).Replace(" ", "") Then
                    AddResults(List1, 1, Pos1, NewPos1 - 1)
                    AddResults(List2, 2, Pos2, NewPos2 - 1)
                    Result.AppendLine()
                    DiffCount += 1
                End If
            Else
                AddResults(List1, 1, Pos1, NewPos1 - 1)
                AddResults(List2, 2, Pos2, NewPos2 - 1)
                Result.AppendLine()
                DiffCount += 1
            End If
            Pos1 = NewPos1
            Pos2 = NewPos2
        Loop
        ' --- See if files are different lengths, even if they may compare up to this point ---
        If Pos1 < List1.Count OrElse Pos2 < List2.Count Then
            DiffCount += 1
        End If
        If DiffCount = 1 Then
            Result.AppendLine("*** 1 difference found ***")
        ElseIf DiffCount > 0 Then
            Result.Append("*** ")
            Result.Append(DiffCount.ToString)
            Result.AppendLine(" differences found ***")
        ElseIf Not QuietMode Then
            Result.AppendLine("*** No differences found ***")
        End If
    End Sub

#End Region

#Region " Public Overridable Routines "

    ''' <summary>
    ''' Resets all compare flags to their initial values.
    ''' </summary>
    Public Overridable Sub ResetFlags()
        IgnoreCase = False
        IgnoreDashes = False
        IgnoreEllipsis = False
        IgnoreQuotes = False
        IgnoreSpaces = False
        TabsToSpaces = False
        SquishSpaces = False
        TrimBlanks = False
        SquishLines = False
        QuietMode = False
        WordMode = False
        TokenMode = False
        ShowCarets = False
        ShowAllCarets = False
        FirstDiffOnly = False
        MaxDiffs = MaxDiffDefault
    End Sub

#End Region

#Region " Protected Overridable Routines "

    Protected Overridable Function FixLine(ByVal CurrLine As String) As String
        If IgnoreCase Then
            CurrLine = CurrLine.ToUpper
        End If
        If IgnoreDashes Then
            Do While CurrLine.Contains("—") ' em dash, chr(151), #8212
                CurrLine = CurrLine.Replace("—", "-")
            Loop
            Do While CurrLine.Contains("–") ' en dash, chr(150), #8211
                CurrLine = CurrLine.Replace("–", "-")
            Loop
            Do While CurrLine.Contains(" -")
                CurrLine = CurrLine.Replace(" -", "-")
            Loop
            Do While CurrLine.Contains("- ")
                CurrLine = CurrLine.Replace("- ", "-")
            Loop
            Do While CurrLine.Contains("--")
                CurrLine = CurrLine.Replace("--", "-")
            Loop
            CurrLine = CurrLine.Replace("&#8212;", " ")
            CurrLine = CurrLine.Replace("&#8211;", " ")
            CurrLine = CurrLine.Replace("-", " ")
        End If
        If IgnoreEllipsis Then
            Do While CurrLine.Contains("…") ' single-char ellipsis, chr(133), #8230
                CurrLine = CurrLine.Replace("…", "...")
            Loop
            Do While CurrLine.Contains(".  .")
                CurrLine = CurrLine.Replace(".  .", "...")
            Loop
            Do While CurrLine.Contains(". . .")
                CurrLine = CurrLine.Replace(". . .", "...")
            Loop
            Do While CurrLine.Contains(". .")
                CurrLine = CurrLine.Replace(". .", "...")
            Loop
            Do While CurrLine.Contains(" ...")
                CurrLine = CurrLine.Replace(" ...", "...")
            Loop
            Do While CurrLine.Contains("... ")
                CurrLine = CurrLine.Replace("... ", "...")
            Loop
            Do While CurrLine.Contains("....")
                CurrLine = CurrLine.Replace("....", "...")
            Loop
            CurrLine = CurrLine.Replace("...", " ")
            CurrLine = CurrLine.Replace("&#8230;", " ")
        End If
        If IgnoreQuotes Then
            Do While CurrLine.Contains("`") ' tic mark, chr(96)
                CurrLine = CurrLine.Replace("`", "'")
            Loop
            Do While CurrLine.Contains(Chr(145)) ' left single quote
                CurrLine = CurrLine.Replace(Chr(145), "'")
            Loop
            Do While CurrLine.Contains(Chr(146)) ' right single quote
                CurrLine = CurrLine.Replace(Chr(146), "'")
            Loop
            Do While CurrLine.Contains(Chr(147)) ' left double quote
                CurrLine = CurrLine.Replace(Chr(147), """")
            Loop
            Do While CurrLine.Contains(Chr(148)) ' right double quote
                CurrLine = CurrLine.Replace(Chr(148), """")
            Loop
            Do While CurrLine.Contains(Chr(171)) ' left european quote
                CurrLine = CurrLine.Replace(Chr(171), """")
            Loop
            Do While CurrLine.Contains(Chr(187)) ' right european quote
                CurrLine = CurrLine.Replace(Chr(187), """")
            Loop
        End If
        If (TabsToSpaces OrElse ShowCarets OrElse ShowAllCarets) AndAlso CurrLine.Contains(vbTab) Then
            CurrLine = CurrLine.Replace(vbTab, " "c)
        End If
        If SquishSpaces Then
            Do While CurrLine.Contains("  ")
                CurrLine = CurrLine.Replace("  ", " ")
            Loop
        End If
        If TrimBlanks Then
            If (CurrLine.EndsWith(" _")) Then ' VB.NET line continuation characters, not needed
                CurrLine = CurrLine.Substring(0, CurrLine.Length - 1).Trim
            ElseIf (CurrLine.StartsWith(" ") OrElse CurrLine.EndsWith(" ")) Then
                CurrLine = CurrLine.Trim
            End If
        End If
        Return CurrLine
    End Function

    Protected Overridable Function IgnoreLine(ByVal CurrLine As String) As Boolean
        If SquishLines Then
            If CurrLine = "" Then Return True
            If TabsToSpaces AndAlso CurrLine.Contains(vbTab) Then
                CurrLine = CurrLine.Replace(vbTab, " "c)
            End If
            If SquishSpaces AndAlso CurrLine.Trim = "" Then Return True
        End If
        Return False
    End Function

#End Region

#Region " Private Routines "

    Private Sub FillLineList(ByVal Filename As String, ByRef LineList As List(Of LineItem))
        Dim CurrLineItem As LineItem
        Dim CurrLineNum As Integer
        Dim KeepChar As Boolean
        Dim CurrLine As String
        Dim FileEncoding As Encoding
        ' --------------------------
        ' --- Save name for inherited classes ---
        CurrentFilename = Filename
        ' --- Process the lines from the file ---
        CurrLineNum = 0
        FileEncoding = GetFileEncoding(Filename)
        If FileEncoding Is Nothing Then
            Throw New SystemException("Binary file - Cannot compare")
        End If
        For Each CurrLine In File.ReadAllLines(Filename, FileEncoding)
            If CurrLine Is Nothing Then Continue For
            ' --- Process line ---
            CurrLineNum += 1
            CurrLineItem = Nothing
            CurrLine = FixLine(CurrLine)
            If WordMode OrElse TokenMode Then
                For Each CurrChar As Char In CurrLine
                    KeepChar = False
                    If WordMode Then
                        If CurrChar >= "A"c AndAlso CurrChar <= "Z"c Then KeepChar = True
                        If CurrChar >= "a"c AndAlso CurrChar <= "z"c Then KeepChar = True
                        If InternationalLetter(CurrChar) Then KeepChar = True
                    End If
                    If TokenMode Then
                        KeepChar = True
                        If CurrChar = " "c Then KeepChar = False
                        If ControlChar(CurrChar) Then KeepChar = False
                    End If
                    If KeepChar Then
                        If IgnoreCase AndAlso (CurrChar >= "a"c AndAlso CurrChar <= "z"c) Then
                            CurrChar = Char.ToUpper(CurrChar)
                        End If
                        If CurrLineItem Is Nothing Then
                            CurrLineItem = New LineItem
                            CurrLineItem.LineNum = CurrLineNum
                            CurrLineItem.LineText = CurrChar
                        Else
                            CurrLineItem.LineText += CurrChar
                        End If
                    ElseIf CurrLineItem IsNot Nothing Then
                        LineList.Add(CurrLineItem)
                        CurrLineItem = Nothing
                    End If
                Next
                If CurrLineItem IsNot Nothing Then
                    LineList.Add(CurrLineItem)
                End If
            ElseIf Not IgnoreLine(CurrLine) Then
                ' --- Add this line to the list ---
                CurrLineItem = New LineItem
                CurrLineItem.LineNum = CurrLineNum
                CurrLineItem.LineText = CurrLine
                LineList.Add(CurrLineItem)
            End If
        Next
    End Sub

    Private Sub AddResults(ByVal CurrList As List(Of LineItem), ByVal CurrListNum As Integer, ByVal PosStart As Integer, ByVal PosEnd As Integer)
        Dim CurrLineNum As Integer = 0
        ' ----------------------------
        If PosEnd < PosStart Then Exit Sub
        For CurrIndex As Integer = PosStart To PosEnd
            With Result
                If CurrLineNum < CurrList(CurrIndex).LineNum Then
                    If CurrLineNum > 0 Then .AppendLine()
                    .Append(CurrListNum.ToString)
                    .Append("> ")
                    .Append(CurrList(CurrIndex).LineNum.ToString.PadLeft(5))
                    .Append(": ")
                    CurrLineNum = CurrList(CurrIndex).LineNum
                Else
                    .Append(" ")
                End If
                .Append(CurrList(CurrIndex).LineText)
            End With
        Next
        Result.AppendLine()
    End Sub

    Private Function GetResults(ByVal CurrList As List(Of LineItem), ByVal CurrListNum As Integer, ByVal PosStart As Integer, ByVal PosEnd As Integer) As String
        Dim CurrLineNum As Integer = 0
        Dim TempResult As New StringBuilder
        ' ---------------------------------
        If PosEnd < PosStart Then Return ""
        For CurrIndex As Integer = PosStart To PosEnd
            With TempResult
                If CurrLineNum < CurrList(CurrIndex).LineNum Then
                    If CurrLineNum > 0 Then .AppendLine()
                    CurrLineNum = CurrList(CurrIndex).LineNum
                Else
                    .Append(" ")
                End If
                .Append(CurrList(CurrIndex).LineText)
            End With
        Next
        TempResult.AppendLine()
        Return TempResult.ToString
    End Function

    Public Shared Function BinaryChar(ByVal CurrChar As Char) As Boolean
        Return BinaryChar(Asc(CurrChar))
    End Function

    Public Shared Function BinaryChar(ByVal CurrChar As Integer) As Boolean
        If CurrChar = 0 Then Return True
        If CurrChar = 127 Then Return True
        If CurrChar = 129 Then Return True
        If CurrChar = 141 Then Return True
        If CurrChar = 143 Then Return True
        If CurrChar = 144 Then Return True
        If CurrChar = 157 Then Return True
        Return False
    End Function

    Private Function ControlChar(ByVal CurrChar As Char) As Boolean
        If Asc(CurrChar) >= 1 AndAlso Asc(CurrChar) <= 31 Then Return True
        Return False
    End Function

    Private Function InternationalLetter(ByVal CurrChar As Char) As Boolean
        If Asc(CurrChar) >= 192 AndAlso Asc(CurrChar) <= 255 Then
            If Asc(CurrChar) = 215 Then Return False ' multiplication sign
            If Asc(CurrChar) = 247 Then Return False ' division sign
            Return True
        End If
        If Asc(CurrChar) = 131 Then Return True
        If Asc(CurrChar) = 138 Then Return True
        If Asc(CurrChar) = 140 Then Return True
        If Asc(CurrChar) = 142 Then Return True
        If Asc(CurrChar) = 154 Then Return True
        If Asc(CurrChar) = 156 Then Return True
        If Asc(CurrChar) = 158 Then Return True
        If Asc(CurrChar) = 159 Then Return True
        If Asc(CurrChar) = 170 Then Return True
        If Asc(CurrChar) = 181 Then Return True
        If Asc(CurrChar) = 186 Then Return True
        Return False
    End Function

#End Region

End Class
