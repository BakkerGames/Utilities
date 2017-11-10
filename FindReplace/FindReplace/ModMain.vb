' -------------------------------
' --- ModMain.vb - 11/10/2017 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' 11/10/2017 - SBakker
'            - Write out UTF8 without BOM.
' 10/05/2016 - SBakker
'            - Added /TRIM mode to remove trailing whitespace.
' 01/13/2016 - SBakker
'            - Added /SQUISH mode to remove blank lines.
' 02/26/2014 - SBakker
'            - When encoding to ASCII, check if any chars exceed ASCW(127). If so, must save as UTF-8
'              instead. NOTE: Should already have UTF-8 BOM or it won't work.
' 11/07/2011 - SBakker
'            - Added /UTF8 and /ASCII options to force output encodings.
' 01/13/2011 - SBakker
'            - Added check to see if CurrLine is Nothing. File.ReadAllLines() can return a
'              Nothing line in some cases, usually at the end of the file.
' 04/17/2009 - SBakker
'            - Building a FindReplace program - tired of not having one!
'            - TODO: Add full RegEx support.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text

Module ModMain

    ' --- Command-line arguments ---
    Dim SearchPattern As String = Nothing
    Dim ReplacePattern As String = Nothing
    Dim FileArg As String = Nothing

    ' --- Command-line options ---
    Private DoSubdirs As Boolean = False
    Private IgnoreCase As Boolean = False
    Private UseRegEx As Boolean = False
    Private QuietMode As Boolean = False
    Private VerboseMode As Boolean = False
    Private OutputUTF8 As Boolean = False
    Private OutputASCII As Boolean = False
    Private SquishFlag As Boolean = False
    Private TrimFlag As Boolean = False

    Private Sub ShowSyntax()
        Console.WriteLine()
        Console.WriteLine("Syntax: " + My.Application.Info.ProductName + " [Options] <SearchPattern> <ReplacePattern> [ <Filename> ]")
        Console.WriteLine()
        Console.WriteLine("         /S      - Check files in specified directory and all subdirectories.")
        Console.WriteLine("         /I      - Case Insensitive")
        Console.WriteLine("         /R      - Use Regular Expressions (only ^ and $ supported now)")
        Console.WriteLine("         /Q      - Quiet Mode - no informational messages")
        Console.WriteLine("         /V      - Verbose - extra informational messages")
        Console.WriteLine("         /UTF8   - Force output to be UTF-8 encoded")
        Console.WriteLine("         /ASCII  - Force output to be ASCII")
        Console.WriteLine("         /SQUISH - Remove blank lines in output")
        Console.WriteLine("         /TRIM   - Remove whitespace at end of lines")
    End Sub

    Public Sub Main()

        Dim CurrArg As String

        Dim FileName As String
        Dim FilePath As String
        Dim FileSpec As String

        ' --- Check Command-line arguments ---
        For i As Integer = 0 To CmdLineArgs.Count - 1
            CurrArg = CmdLineArgs.Arg(i)
            If CurrArg.StartsWith("/") OrElse CurrArg.StartsWith("-") Then
                ' --- Options ---
                Select Case CurrArg.Substring(1).ToUpper
                    Case "S"
                        DoSubdirs = True
                    Case "I"
                        IgnoreCase = True
                    Case "R"
                        UseRegEx = True
                    Case "Q"
                        QuietMode = True
                    Case "V"
                        VerboseMode = True
                    Case "UTF8"
                        OutputUTF8 = True
                    Case "ASCII"
                        OutputASCII = True
                    Case "SQUISH"
                        SquishFlag = True
                    Case "TRIM"
                        TrimFlag = True
                    Case Else
                        ShowSyntax()
                        Exit Sub
                End Select
                Continue For
            End If
            If SearchPattern Is Nothing Then
                SearchPattern = CurrArg
            ElseIf ReplacePattern Is Nothing Then
                ReplacePattern = CurrArg
            ElseIf FileArg Is Nothing Then
                FileArg = CurrArg
            Else
                ShowSyntax()
                Exit Sub
            End If
        Next

        ' --- Make sure something was entered ---
        If SearchPattern Is Nothing OrElse ReplacePattern Is Nothing OrElse SearchPattern = "" Then
            ShowSyntax()
            Exit Sub
        End If

        ' --- check regular expression ---
        If UseRegEx Then
            If SearchPattern <> "^" AndAlso SearchPattern <> "$" Then
                ShowSyntax()
                Exit Sub
            End If
        End If

        ' --- check if using standard input/output ---
        If FileArg Is Nothing Then
            Dim CurrLine As String
            CurrLine = Console.ReadLine()
            Do While CurrLine IsNot Nothing
                Console.WriteLine(DoReplace(CurrLine))
                CurrLine = Console.ReadLine
            Loop
            Exit Sub
        End If

        ' --- Parse the filenames ---
        If FileArg.IndexOf("?"c) >= 0 OrElse FileArg.IndexOf("*"c) >= 0 Then
            Dim FileNames() As String
            If FileArg.StartsWith("\") OrElse FileArg.Substring(1, 1) = ":" Then
                FilePath = FileArg
            Else
                FilePath = My.Computer.FileSystem.CurrentDirectory + "\" + FileArg
            End If
            If FilePath.LastIndexOf("\") < 0 Then
                Console.Error.WriteLine("Unable to process file path: """ + FilePath + """")
                Exit Sub
            End If
            FileSpec = FilePath.Substring(FilePath.LastIndexOf("\") + 1)
            FilePath = FilePath.Substring(0, FilePath.LastIndexOf("\"))
            If VerboseMode Then
                Console.WriteLine("FilePath = """ + FilePath + """")
                Console.WriteLine("FileSpec = """ + FileSpec + """")
            End If
            If FilePath.IndexOf("?"c) >= 0 OrElse FilePath.IndexOf("*"c) >= 0 Then
                Console.Error.WriteLine("Unable to process file path: """ + FilePath + """")
                Exit Sub
            End If
            If DoSubdirs Then
                FileNames = Directory.GetFiles(FilePath, FileSpec, SearchOption.AllDirectories)
            Else
                FileNames = Directory.GetFiles(FilePath, FileSpec, SearchOption.TopDirectoryOnly)
            End If
            For Each FileName In FileNames
                If VerboseMode Then
                    Console.WriteLine(FileName + " - Checking...")
                End If
                DoChangeFile(FileName)
            Next
        Else
            FileName = FileArg
            If Not File.Exists(FileName) Then
                Console.Error.WriteLine("File not found: """ + FileName + """")
                Exit Sub
            End If
            DoChangeFile(FileName)
        End If

    End Sub

    Private Sub DoChangeFile(ByVal FileName As String)
        Dim Lines() As String
        Dim Result As StringBuilder
        Dim CurrLine As String
        Dim Changed As Boolean = False
        Dim CanUseASCII As Boolean
        Dim CurrChar As Integer
        Dim CurrEncoding As Encoding
        ' ----------------------------
        ' --- Check file encoding and whether ASCII can be used ---
        CanUseASCII = True
        Dim sr As New StreamReader(FileName)
        Do While Not sr.EndOfStream
            CurrChar = sr.Read
            If CurrChar = 0 OrElse CurrChar > 127 Then
                CanUseASCII = False
                Exit Do
            End If
        Loop
        CurrEncoding = sr.CurrentEncoding
        sr.Close()
        ' --- Read all the lines into memory ---
        Lines = File.ReadAllLines(FileName) ' , CurrEncoding)
        If VerboseMode Then
            Console.WriteLine(FileName + " - Line Count = " + (Lines.GetUpperBound(0) + 1).ToString)
        End If
        ' --- Make any changes necessary ---
        Changed = False
        Result = New StringBuilder
        For LineNum As Integer = 0 To Lines.GetUpperBound(0)
            CurrLine = Lines(LineNum)
            If CurrLine Is Nothing Then Continue For
            CurrLine = DoReplace(CurrLine)
            If TrimFlag Then
                CurrLine = CurrLine.TrimEnd
            End If
            If Lines(LineNum) <> CurrLine Then
                Changed = True
            End If
            If SquishFlag AndAlso String.IsNullOrWhiteSpace(CurrLine) Then
                Changed = True
            Else
                Result.AppendLine(CurrLine)
            End If
        Next
        If Changed Then
            Try
                If FileName.ToUpper.EndsWith(".BAT") Then
                    File.WriteAllText(FileName, Result.ToString, Encoding.ASCII)
                ElseIf OutputASCII AndAlso CanUseASCII Then
                    File.WriteAllText(FileName, Result.ToString, Encoding.ASCII)
                ElseIf OutputASCII AndAlso Not CanUseASCII Then
                    File.WriteAllText(FileName, Result.ToString, new UTF8Encoding(false, true))
                ElseIf OutputUTF8 Then
                    File.WriteAllText(FileName, Result.ToString, new UTF8Encoding(false, true))
                Else
                    File.WriteAllText(FileName, Result.ToString, CurrEncoding)
                End If
                If Not QuietMode Then Console.WriteLine(FileName + " - Changed")
            Catch ex As Exception
                Console.Error.WriteLine(FileName + " - Error saving file!")
            End Try
        End If
    End Sub

    Private Function DoReplace(ByVal CurrLine As String) As String
        Dim Result As New StringBuilder
        Dim CurrChar As Integer = 0
        Dim MatchPoint As Integer = 0
        Dim CompareType As StringComparison
        ' ---------------------------------
        If UseRegEx Then
            If SearchPattern = "^" Then
                Result.Append(ReplacePattern)
                Result.Append(CurrLine)
            ElseIf SearchPattern = "$" Then
                Result.Append(CurrLine)
                Result.Append(ReplacePattern)
            End If
        Else
            If IgnoreCase Then
                CompareType = StringComparison.CurrentCultureIgnoreCase
            Else
                CompareType = StringComparison.CurrentCulture
            End If
            Do While CurrChar < CurrLine.Length
                MatchPoint = CurrLine.IndexOf(SearchPattern, MatchPoint, CompareType)
                If MatchPoint >= 0 Then
                    If CurrChar < MatchPoint Then
                        Result.Append(CurrLine.Substring(CurrChar, MatchPoint - CurrChar))
                        CurrChar = MatchPoint
                    End If
                    Result.Append(ReplacePattern)
                    CurrChar += SearchPattern.Length
                    MatchPoint = CurrChar
                Else
                    Result.Append(CurrLine.Substring(CurrChar))
                    CurrChar = CurrLine.Length
                End If
            Loop
        End If
        Return Result.ToString
    End Function

End Module
