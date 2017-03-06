' -------------------------------
' --- ModMain.vb - 05/30/2012 ---
' -------------------------------

' ------------------------------------------------------------------------------------------
' 05/30/2012 - SBakker - Bug 
'            - Change all non-ASCII characters to the best equivalent, or "?".
' ------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text

Module ModMain

    Dim FileArg As String = Nothing
    Dim FileName As String
    Dim FilePath As String
    Dim FileSpec As String
    Dim FileChanged As Boolean = False

    Private Sub ShowSyntax()
        Console.WriteLine()
        Console.WriteLine("Syntax: " + My.Application.Info.ProductName + " <Filename>")
        Console.WriteLine()
        Console.WriteLine("This will take an Windows-1251 encoded file and change all")
        Console.WriteLine("non-ASCII chars to ASCII (32-126 + control chars).")
    End Sub

    Public Sub Main()

        Dim CurrArg As String
        Dim CurrChar As Integer
        Dim FileNameList As New List(Of String)

        ' --- Check Command-line arguments ---
        For i As Integer = 0 To CmdLineArgs.Count - 1
            CurrArg = CmdLineArgs.Arg(i)
            If CurrArg.StartsWith("/") OrElse CurrArg.StartsWith("-") Then
                ' --- Options ---
                Select Case CurrArg.Substring(1).ToUpper
                    Case Else
                        ShowSyntax()
                        Exit Sub
                End Select
                Continue For
            End If
            If FileArg Is Nothing Then
                FileArg = CurrArg
            Else
                ShowSyntax()
                Exit Sub
            End If
        Next

        ' --- check if using standard input/output ---
        If FileArg Is Nothing Then
            Try
                CurrChar = Console.Read
                Do While CurrChar >= 0
                    Console.Write(DoReplaceChar(CurrChar))
                    CurrChar = Console.Read
                Loop
            Catch ex As Exception
                ' --- ignore error ---
            End Try
            Exit Sub
        End If

        ' --- Parse the filenames ---
        If FileArg.IndexOf("?"c) >= 0 OrElse FileArg.IndexOf("*"c) >= 0 Then
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
            If FilePath.IndexOf("?"c) >= 0 OrElse FilePath.IndexOf("*"c) >= 0 Then
                Console.Error.WriteLine("Unable to process file path: """ + FilePath + """")
                Exit Sub
            End If
            Dim Filenames() As String = Directory.GetFiles(FilePath, FileSpec, SearchOption.TopDirectoryOnly)
            For Each FileName In Filenames
                FileNameList.Add(FileName)
            Next
        Else
            FileName = FileArg
            If Not File.Exists(FileName) Then
                Console.Error.WriteLine("File not found: """ + FileName + """")
                Exit Sub
            End If
            FileNameList.Add(FileName)
        End If

        For Each FileName In FileNameList
            Dim sr As New StreamReader(FileName, Encoding.Default, True)
            Dim sw As New StreamWriter(FileName + ".tmp", False, Encoding.ASCII)
            FileChanged = False
            Try
                CurrChar = sr.Read
                Do While CurrChar >= 0
                    sw.Write(Chr(DoReplaceChar(CurrChar)))
                    CurrChar = sr.Read
                Loop
            Catch ex As Exception
                ' --- ignore error ---
            End Try
            sr.Close()
            sw.Close()
            sr = Nothing
            sw = Nothing
            Try
                If FileChanged Then
                    File.Copy(FileName + ".tmp", FileName, True)
                End If
                File.Delete(FileName + ".tmp")
            Catch ex As Exception
                ' --- Can't do much in an automated program ---
            End Try
        Next

    End Sub

    Private Function DoReplaceChar(ByVal CurrChar As Integer) As Integer
        Static QuestionMark As Integer = Asc("?")
        Dim Result As Integer = QuestionMark
        ' ---------------------------------------
        ' --- Single character results ---
        Select Case CurrChar
            Case 9, 10, 13 ' Tab, LF, CR
                Result = CurrChar
            Case 0 To 31
                Result = 32 ' change all other cootrol chars to spaces
            Case 32 To 126
                Result = CurrChar
            Case 160 ' nbsp
                Result = 32
            Case 173 ' soft hyphen
                Result = Asc("-")
            Case Asc("À")
                Result = Asc("A")
            Case Asc("Á")
                Result = Asc("A")
            Case Asc("Â")
                Result = Asc("A")
            Case Asc("Ã")
                Result = Asc("A")
            Case Asc("Ä")
                Result = Asc("A")
            Case Asc("Å")
                Result = Asc("A")
            Case Asc("à")
                Result = Asc("a")
            Case Asc("á")
                Result = Asc("a")
            Case Asc("â")
                Result = Asc("a")
            Case Asc("ã")
                Result = Asc("a")
            Case Asc("ä")
                Result = Asc("a")
            Case Asc("å")
                Result = Asc("a")
            Case Asc("Ç")
                Result = Asc("C")
            Case Asc("ç")
                Result = Asc("c")
            Case Asc("È")
                Result = Asc("E")
            Case Asc("É")
                Result = Asc("E")
            Case Asc("Ê")
                Result = Asc("E")
            Case Asc("Ë")
                Result = Asc("E")
            Case Asc("è")
                Result = Asc("e")
            Case Asc("é")
                Result = Asc("e")
            Case Asc("ê")
                Result = Asc("e")
            Case Asc("ë")
                Result = Asc("e")
            Case Asc("Ì")
                Result = Asc("I")
            Case Asc("Í")
                Result = Asc("I")
            Case Asc("Î")
                Result = Asc("I")
            Case Asc("Ï")
                Result = Asc("I")
            Case Asc("ì")
                Result = Asc("i")
            Case Asc("í")
                Result = Asc("i")
            Case Asc("î")
                Result = Asc("i")
            Case Asc("ï")
                Result = Asc("i")
            Case Asc("Ñ")
                Result = Asc("N")
            Case Asc("ñ")
                Result = Asc("n")
            Case Asc("Ò")
                Result = Asc("O")
            Case Asc("Ó")
                Result = Asc("O")
            Case Asc("Ô")
                Result = Asc("O")
            Case Asc("Õ")
                Result = Asc("O")
            Case Asc("Ö")
                Result = Asc("O")
            Case Asc("Ø")
                Result = Asc("O")
            Case Asc("ò")
                Result = Asc("o")
            Case Asc("ó")
                Result = Asc("o")
            Case Asc("ô")
                Result = Asc("o")
            Case Asc("õ")
                Result = Asc("o")
            Case Asc("ö")
                Result = Asc("o")
            Case Asc("ø")
                Result = Asc("o")
            Case Asc("Š")
                Result = Asc("S")
            Case Asc("š")
                Result = Asc("s")
            Case Asc("Ù")
                Result = Asc("U")
            Case Asc("Ú")
                Result = Asc("U")
            Case Asc("Û")
                Result = Asc("U")
            Case Asc("Ü")
                Result = Asc("U")
            Case Asc("ù")
                Result = Asc("u")
            Case Asc("ú")
                Result = Asc("u")
            Case Asc("û")
                Result = Asc("u")
            Case Asc("ü")
                Result = Asc("u")
            Case Asc("Ÿ")
                Result = Asc("Y")
            Case Asc("Ý")
                Result = Asc("Y")
            Case Asc("ÿ")
                Result = Asc("y")
            Case Asc("ý")
                Result = Asc("y")
            Case Asc("Ž")
                Result = Asc("Z")
            Case Asc("ž")
                Result = Asc("z")
            Case 130, 139, 145, 146, 155, 180 ' curved single quotes, acute accent, single angle quotes
                Result = Asc("'")
            Case &H201A, &H2039, &H2018, &H2019, &H203A ' curved single quotes, acute accent, single angle quotes
                Result = Asc("'")
            Case 132, 147, 148, 171, 187 ' curved and angle double quotes
                Result = Asc("""")
            Case &H201E, &H201C, &H201D ' curved and angle double quotes
                Result = Asc("""")
            Case 136, &H2C6 ' circumflex
                Result = Asc("^")
            Case 152, &H2DC ' tilde
                Result = Asc("~")
            Case 166 ' broken bar
                Result = Asc("|")
            Case 161 ' inverted exclamation mark
                Result = Asc("!")
            Case 191 ' inverted question mark
                Result = Asc("?")
            Case 170 ' superscript "a"
                Result = Asc("a")
            Case 186 ' superscript "0"
                Result = Asc("0")
            Case 185 ' superscript "1"
                Result = Asc("1")
            Case 178 ' superscript "2"
                Result = Asc("2")
            Case 179 ' superscript "3"
                Result = Asc("3")
            Case Else
                Result = QuestionMark
        End Select
        If Result <> CurrChar Then
            FileChanged = True
        End If
        Return Result
    End Function

End Module
