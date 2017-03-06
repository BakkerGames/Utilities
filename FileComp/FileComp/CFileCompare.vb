' ---------------------------------
' --- CFileCompare - 07/28/2010 ---
' ---------------------------------

Imports System.IO
Imports System.Text

Public Class CFileCompare

    ' --- internal constants ---

    Private Const MAX As Integer = 1024

    ' --- internal variables and storage ---

    Private MyFilename1 As String
    Private MyFilename2 As String
    Private File1 As StreamReader
    Private File2 As StreamReader
    Private Drive1 As String = ""
    Private Drive2 As String = ""

    Private BufferStart1 As Integer
    Private BufferStart2 As Integer
    Private BufferEnd1 As Integer
    Private BufferEnd2 As Integer

    Private Buffer1(MAX - 1) As String
    Private Buffer2(MAX - 1) As String

    Private Xref(MAX - 1, MAX - 1) As Byte

    Private CurrLine1 As Integer
    Private CurrLine2 As Integer
    Private LineNum1(MAX - 1) As Integer
    Private LineNum2(MAX - 1) As Integer

    Private RestOfLine1 As String = ""
    Private RestOfLine2 As String = ""

    ' --- public options. set to true if desired ---

    Public IgnoreCase As Boolean = False
    Public TabsToSpaces As Boolean = False
    Public SquishSpaces As Boolean = False
    Public TrimBlanks As Boolean = False
    Public SquishLines As Boolean = False
    Public QuietMode As Boolean = False
    Public WordMode As Boolean = False
    Public TokenMode As Boolean = False
    Public NormalizeNumbers As Boolean = False
    Public ShowCarets As Boolean = False
    Public ShowAllCarets As Boolean = False

    Public DiffCount As Integer = 0

    ' --- This holds the results until requested ---

    Private MyResults As StringBuilder = Nothing

    Public Sub Clear()
        IgnoreCase = False
        TabsToSpaces = False
        SquishSpaces = False
        TrimBlanks = False
        SquishLines = False
        QuietMode = False
        WordMode = False
        TokenMode = False
        NormalizeNumbers = False
        ShowCarets = False
        ShowAllCarets = False
        DiffCount = 0
        MyResults = Nothing
    End Sub

    Public Function Results() As String
        If MyResults Is Nothing Then
            Return ""
        End If
        Return MyResults.ToString
    End Function

    ' --- public methods ---

    Public Sub DoCompare(ByVal Filename1 As String, ByVal Filename2 As String)
        Dim Longest As Integer
        Dim LongestStart1 As Integer
        Dim LongestStart2 As Integer
        ' --- prepare all files ---
        MyResults = Nothing
        If Not System.IO.File.Exists(Filename1) Then
            MyResults.AppendLine("Cannot find file: " + Filename1)
            Exit Sub
        End If
        If Not System.IO.File.Exists(Filename2) Then
            MyResults.AppendLine("Cannot find file: " + Filename2)
            Exit Sub
        End If
        Try
            File1 = New StreamReader(Filename1, System.Text.Encoding.UTF8)
        Catch
            MyResults.AppendLine("Unable to open file: " + Filename1)
            Exit Sub
        End Try
        Try
            File2 = New StreamReader(Filename2, System.Text.Encoding.UTF8)
        Catch
            MyResults.AppendLine("Unable to open file: " + Filename2)
            Exit Sub
        End Try
        ' --- Get Drives for the two files ---
        Drive1 = Left(Filename1.ToUpper, 3)
        Drive2 = Left(Filename2.ToUpper, 3)
        ' --- init ---
        MyResults = New StringBuilder
        DiffCount = 0
        BufferStart1 = 0
        BufferStart2 = 0
        BufferEnd1 = -1
        BufferEnd2 = -1
        CurrLine1 = 0
        CurrLine2 = 0
        MyFilename1 = Filename1
        MyFilename2 = Filename2
        ' --- identify files to compare ---
        If Not QuietMode Then
            MyResults.AppendLine("Comparing files """ + MyFilename1 + """ and """ + MyFilename2 + """")
            MyResults.AppendLine("")
        End If
        ' --- load buffers first time ---
        FillBuffers()
        FillGrid()
        ' --- loop for each chunk of buffered lines ---
        Do While (BufferEnd1 >= 0) Or (BufferEnd2 >= 0)
            ' --- find the longest line in the defined section ---
            Longest = FindLongest(BufferStart1, BufferEnd1, BufferStart2, BufferEnd2, LongestStart1, LongestStart2)
            If Longest = 0 Then
                ' --- all buffered lines are different ---
                OutputChanged(BufferStart1, BufferEnd1, BufferStart2, BufferEnd2)
                Exit Do
            End If
            ' --- process section before longest ---
            ProcessLines(BufferStart1, LongestStart1 - 1, BufferStart2, LongestStart2 - 1)
            ' --- skip over longest section ---
            BufferStart1 = LongestStart1 + Longest
            BufferStart2 = LongestStart2 + Longest
            ' --- refill buffers for next round ---
            FillBuffers()
            FillGrid()
        Loop
        ' --- done ---
        If (Not QuietMode) And DiffCount = 0 Then
            MyResults.AppendLine("*** No differences found ***")
            MyResults.AppendLine("")
        End If
        If DiffCount = 1 Then
            MyResults.AppendLine("*** " + DiffCount.ToString + " difference found ***")
            MyResults.AppendLine("")
        End If
        If DiffCount > 1 Then
            MyResults.AppendLine("*** " + DiffCount.ToString + " differences found ***")
            MyResults.AppendLine("")
        End If
        ' --- close all files ---
        File1.Close()
        File2.Close()
    End Sub

    ' --- private methods and functions ---

    Private Sub FillBuffers()
        Dim LoopNum As Integer
        Dim CurrToken As String
        Dim TempLine As String
        ' --- remove processed lines, shuffle rest to top ---
        If BufferStart1 > 0 Then
            For LoopNum = BufferStart1 To BufferEnd1
                Buffer1(LoopNum - BufferStart1) = Buffer1(LoopNum)
                LineNum1(LoopNum - BufferStart1) = LineNum1(LoopNum)
            Next
            BufferEnd1 -= BufferStart1
            BufferStart1 = 0
        End If
        If BufferStart2 > 0 Then
            For LoopNum = BufferStart2 To BufferEnd2
                Buffer2(LoopNum - BufferStart2) = Buffer2(LoopNum)
                LineNum2(LoopNum - BufferStart2) = LineNum2(LoopNum)
            Next
            BufferEnd2 -= BufferStart2
            BufferStart2 = 0
        End If
        ' --- fill Buffer1 ---
        Do While ((File1.Peek >= 0) Or (RestOfLine1 <> "")) And (BufferEnd1 < MAX - 1)
            If WordMode Or TokenMode Then
                Do
                    If (File1.Peek >= 0) And (RestOfLine1 = "") Then
                        RestOfLine1 = FixLine(File1.ReadLine, Drive1)
                        CurrLine1 += 1
                    End If
                    CurrToken = GetNextToken(RestOfLine1)
                Loop Until (CurrToken <> "") Or (File1.Peek < 0)
                If CurrToken = "" Then Exit Do
                BufferEnd1 += 1
                Buffer1(BufferEnd1) = CurrToken
            Else
                TempLine = FixLine(File1.ReadLine, Drive1)
                If (Not SquishLines) Or (TempLine <> "") Then
                    BufferEnd1 += 1
                    Buffer1(BufferEnd1) = TempLine
                End If
                CurrLine1 += 1
            End If
            If BufferEnd1 >= 0 Then
                LineNum1(BufferEnd1) = CurrLine1
            End If
        Loop
        ' --- fill Buffer2 ---
        Do While ((File2.Peek >= 0) Or (RestOfLine2 <> "")) And (BufferEnd2 < MAX - 1)
            If WordMode Or TokenMode Then
                Do
                    If (File2.Peek >= 0) And (RestOfLine2 = "") Then
                        RestOfLine2 = FixLine(File2.ReadLine, Drive2)
                        CurrLine2 += 1
                    End If
                    CurrToken = GetNextToken(RestOfLine2)
                Loop Until (CurrToken <> "") Or (File2.Peek < 0)
                If CurrToken = "" Then Exit Do
                BufferEnd2 += 1
                Buffer2(BufferEnd2) = CurrToken
            Else
                TempLine = FixLine(File2.ReadLine, Drive2)
                If (Not SquishLines) Or (TempLine <> "") Then
                    BufferEnd2 += 1
                    Buffer2(BufferEnd2) = TempLine
                End If
                CurrLine2 += 1
            End If
            If BufferEnd2 >= 0 Then
                LineNum2(BufferEnd2) = CurrLine2
            End If
        Loop
    End Sub

    Private Function GetNextToken(ByRef Line As String) As String
        Dim Result As New StringBuilder
        Dim Chars() As Char = Line.ToCharArray
        Dim CharNum As Integer = 0
        Do While CharNum <= Chars.GetUpperBound(0)
            If (Chars(CharNum) >= "A"c And Chars(CharNum) <= "Z"c) Or _
               (Chars(CharNum) >= "a"c And Chars(CharNum) <= "z"c) Or _
               (Chars(CharNum) >= "0"c And Chars(CharNum) <= "9"c) Or _
               (Chars(CharNum) = "_") Or (Asc(Chars(CharNum)) >= 128) Then
                Result.Append(Chars(CharNum))
            ElseIf TokenMode And Asc(Chars(CharNum)) > Asc(" ") Then
                Result.Append(Chars(CharNum))
            ElseIf Result.Length > 0 Then
                Exit Do
            End If
            CharNum += 1
        Loop
        If CharNum + 1 > Line.Length Then
            Line = ""
        Else
            Line = Line.Substring(CharNum + 1)
        End If
        Return Result.ToString
    End Function

    Private Function FixLine(ByVal Line As String, ByVal FileDrive As String) As String
        ' --- modify input line based on specified flags ---
        If IgnoreCase Then
            Line = Line.ToUpper
        End If
        If TabsToSpaces Or WordMode Or TokenMode Then
            Line = Line.Replace(vbTab, " ")
        End If
        If SquishSpaces Or WordMode Or TokenMode Then
            Do While Line.IndexOf("  ") >= 0
                Line = Line.Replace("  ", " ")
            Loop
        End If
        If TrimBlanks Or WordMode Or TokenMode Then
            Line = Line.Trim
        End If
        ' --- done ---
        Return Line
    End Function

    Private Sub FillGrid()
        Dim Loop1 As Integer
        Dim Loop2 As Integer
        ' --- set all flags for all matching lines ---
        For Loop1 = BufferStart1 To BufferEnd1
            For Loop2 = BufferStart2 To BufferEnd2
                If Buffer1(Loop1) = Buffer2(Loop2) Then
                    Xref(Loop1, Loop2) = 1 ' matches
                Else
                    Xref(Loop1, Loop2) = 0 ' doesn't match
                End If
            Next
        Next
    End Sub

    Private Sub ProcessLines(ByRef Start1 As Integer, ByVal End1 As Integer, ByRef Start2 As Integer, ByVal End2 As Integer)
        Dim Longest As Integer
        Dim LongestStart1 As Integer
        Dim LongestStart2 As Integer
        ' --- check if both sections are empty ---
        If (End1 < Start1) And (End2 < Start2) Then
            Exit Sub ' nothing to do
        End If
        ' --- find the longest line in the defined section ---
        Longest = FindLongest(Start1, End1, Start2, End2, LongestStart1, LongestStart2)
        If Longest = 0 Then
            ' --- all lines different ---
            OutputChanged(Start1, End1, Start2, End2)
        Else
            ' --- process section before longest ---
            ProcessLines(Start1, LongestStart1 - 1, Start2, LongestStart2 - 1)
            ' --- skip over longest section ---
            Start1 = LongestStart1 + Longest
            Start2 = LongestStart2 + Longest
            ' --- process section after longest ---
            ProcessLines(Start1, End1, Start2, End2)
        End If
        ' --- bump forward start pointers over processed section ---
        Start1 = End1 + 1
        Start2 = End2 + 1
    End Sub

    Private Function FindLongest(ByVal Start1 As Integer, ByVal End1 As Integer, ByVal Start2 As Integer, ByVal End2 As Integer, ByRef LongestStart1 As Integer, ByRef LongestStart2 As Integer) As Integer
        Dim Loop1 As Integer
        Dim Loop2 As Integer
        Dim Loop3 As Integer
        Dim MaxDiag As Integer
        Dim DoCheck As Boolean
        Dim Longest As Integer
        ' --- init ---
        Longest = 0
        LongestStart1 = -1
        LongestStart2 = -1
        ' --- loop through all diagonals ---
        For Loop1 = Start1 To End1
            For Loop2 = Start2 To End2
                ' --- look for start of a diagonal ---
                If Xref(Loop1, Loop2) = 1 Then
                    ' --- find out if this diagonal has already been checked ---
                    DoCheck = True
                    If (Loop1 > Start1) And (Loop2 > Start2) Then
                        If Xref(Loop1 - 1, Loop2 - 1) = 1 Then
                            ' --- we are inside of a bigger diagonal ---
                            DoCheck = False
                        End If
                    End If
                    ' --- check this diagonal ---
                    If DoCheck Then
                        ' --- find maximum diagonal size ---
                        MaxDiag = End1 - Loop1
                        If (MaxDiag > End2 - Loop2) Then
                            MaxDiag = End2 - Loop2
                        End If
                        ' --- only check if it could be longer ---
                        If MaxDiag >= Longest Then
                            For Loop3 = 0 To MaxDiag
                                ' --- check if found end of diagonal ---
                                If Xref(Loop1 + Loop3, Loop2 + Loop3) = 0 Then
                                    Exit For
                                End If
                                ' --- update longest info ---
                                If Longest < Loop3 + 1 Then
                                    Longest = Loop3 + 1
                                    LongestStart1 = Loop1
                                    LongestStart2 = Loop2
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        Next
        Return Longest
    End Function

    Private Sub OutputChanged(ByVal Start1 As Integer, ByVal End1 As Integer, _
                              ByVal Start2 As Integer, ByVal End2 As Integer)
        Dim LoopNum As Integer
        Dim TempLine As String
        Dim OutLine As StringBuilder
        Dim LastLineNum As Integer
        ' --- output heading if not done yet ---
        If QuietMode And DiffCount = 0 Then
            MyResults.AppendLine("Comparing files """ + MyFilename1 + """ and """ + MyFilename2 + """")
            MyResults.AppendLine("")
        End If
        ' --- update difference counter, used for ending messages ---
        DiffCount += 1
        ' --- check if all buffered lines are different ---
        If ((Start1 = 0) And (End1 = MAX - 1)) Or ((Start2 = 0) And (End2 = MAX - 1)) Then
            ' --- all buffered lines are different ---
            MyResults.AppendLine("*** Files are very different ***")
            DiffCount = -1 ' don't print ending message
            Exit Sub
        End If
        ' --- output diffs for File1 ---
        If Start1 <= End1 Then
            OutLine = New StringBuilder
            LastLineNum = -1
            For LoopNum = Start1 To End1
                If LineNum1(LoopNum) <> LastLineNum Then
                    If OutLine.Length > 0 Then OutLine.Append(vbCrLf)
                    OutLine.Append("1> ")
                    TempLine = LineNum1(LoopNum).ToString.Trim
                    If TempLine.Length < 5 Then
                        TempLine = Right("     " + TempLine, 5)
                    End If
                    OutLine.Append(TempLine)
                    OutLine.Append(":")
                    LastLineNum = LineNum1(LoopNum)
                End If
                OutLine.Append(" ")
                OutLine.Append(Buffer1(LoopNum))
            Next
            MyResults.AppendLine(OutLine.ToString)
        End If
        ' --- output diffs for File2 ---
        If Start2 <= End2 Then
            OutLine = New StringBuilder
            LastLineNum = -1
            For LoopNum = Start2 To End2
                If LineNum2(LoopNum) <> LastLineNum Then
                    If OutLine.Length > 0 Then OutLine.Append(vbCrLf)
                    OutLine.Append("2> ")
                    TempLine = LineNum2(LoopNum).ToString.Trim
                    If TempLine.Length < 5 Then
                        TempLine = Right("     " + TempLine, 5)
                    End If
                    OutLine.Append(TempLine)
                    OutLine.Append(":")
                    LastLineNum = LineNum2(LoopNum)
                End If
                OutLine.Append(" ")
                OutLine.Append(Buffer2(LoopNum))
            Next
            MyResults.AppendLine(OutLine.ToString)
        End If
        ' --- add a blank line ---
        MyResults.AppendLine("")
    End Sub

End Class
