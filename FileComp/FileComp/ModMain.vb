' -------------------------------
' --- ModMain.vb - 11/10/2017 ---
' -------------------------------

' ------------------------------------------------------------------------------------------
' 11/10/2017 - SBakker
'            - Write out UTF8 without BOM.
' 01/07/2011 - Moved FileCompareClass into its own data class, so it can be used everywhere.
' ------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports FileCompareDataClass

Module ModMain

    Private MyFileCompare As New FileCompareClass
    Private Filename1 As String = ""
    Private Filename2 As String = ""
    Private Filename3 As String = ""
    Private File3 As StreamWriter

    Sub Main()

        Dim FoundError As Boolean = False

        ' --- Read through command line arguments ---
        For Each Arg As String In My.Application.CommandLineArgs
            If Arg.StartsWith("/") Then
                ' --- Handle options ---
                Select Case Arg.ToUpper
                    Case "/I"
                        MyFileCompare.IgnoreCase = True
                    Case "/ID"
                        MyFileCompare.IgnoreDashes = True
                    Case "/IE"
                        MyFileCompare.IgnoreEllipsis = True
                    Case "/IQ"
                        MyFileCompare.IgnoreQuotes = True
                    Case "/C"
                        MyFileCompare.ShowCarets = True
                    Case "/CA"
                        MyFileCompare.ShowCarets = True
                        MyFileCompare.ShowAllCarets = True
                    Case "/Q"
                        MyFileCompare.QuietMode = True
                    Case "/S"
                        MyFileCompare.SquishSpaces = True
                        MyFileCompare.TabsToSpaces = True
                        MyFileCompare.TrimBlanks = True
                        MyFileCompare.SquishLines = True
                    Case "/W"
                        MyFileCompare.WordMode = True
                    Case "/T"
                        MyFileCompare.TokenMode = True
                    Case Else
                        FoundError = True
                        Exit For
                End Select
            ElseIf Filename1 = "" Then
                Filename1 = Arg
            ElseIf Filename2 = "" Then
                Filename2 = Arg
            ElseIf Filename3 = "" Then
                Filename3 = Arg
            Else
                FoundError = True
                Exit For
            End If
        Next

        If Not FoundError Then
            If Filename1 = "" OrElse Filename2 = "" Then
                FoundError = True
            ElseIf MyFileCompare.TokenMode AndAlso MyFileCompare.WordMode Then ' mutually exclusive
                FoundError = True
            ElseIf MyFileCompare.ShowCarets AndAlso MyFileCompare.WordMode Then ' no carets with words
                FoundError = True
            ElseIf MyFileCompare.ShowCarets AndAlso MyFileCompare.TokenMode Then ' no carets with tokens
                FoundError = True
            ElseIf Filename1.IndexOf("?") >= 0 OrElse Filename1.IndexOf("*") >= 0 Then
                If Not Directory.Exists(Filename2) Then
                    FoundError = True
                End If
            End If
        End If

        If FoundError Then
            ShowSyntax()
            Exit Sub
        End If

        ' --- Fix filename2 if only a path is specified ---
        If Filename2.EndsWith(":") OrElse Filename2.EndsWith("\") Then
            Filename2 = Filename2 + Filename1.Substring(Math.Max(Filename1.LastIndexOf(":"), Filename1.LastIndexOf("\")) + 1)
        Else
            Try
                If (GetAttr(Filename2) And FileAttribute.Directory) = FileAttribute.Directory Then
                    Filename2 += "\" + Filename1.Substring(Math.Max(Filename1.LastIndexOf(":"), Filename1.LastIndexOf("\")) + 1)
                End If
            Catch
            End Try
        End If

        ' --- open outputfile if specified ---
        If Filename3 <> "" Then
            File3 = New StreamWriter(Filename3, False, New UTF8Encoding(False, True))
        End If

        ' --- compare the files ---
        MyFileCompare.DoCompare(Filename1, Filename2)

        ' --- Write out the results ---
        If Filename3 <> "" Then
            File3.Write(MyFileCompare.Results)
        Else
            Console.Write(MyFileCompare.Results)
        End If

        ' --- this must be done last, to finish writing to the output file ---
        If Filename3 <> "" Then
            File3.Close()
        End If

    End Sub

    Private Sub ShowSyntax()
        Console.WriteLine("FileComp - Version " + My.Application.Info.Version.ToString)
        Console.WriteLine("")
        Console.WriteLine("Syntax: " + My.Application.Info.AssemblyName + " [options] File1 (File2|Dir2) [OutputFile]")
        Console.WriteLine("")
        Console.WriteLine("        /i    - Ignore case (compared as upper case)")
        Console.WriteLine("        /id   - Ignore dashes (-)")
        Console.WriteLine("        /ie   - Ignore ellipsis (...)")
        Console.WriteLine("        /iq   - Ignore quotes (change all Unicode single and double quotes to ASCII)")
        Console.WriteLine("        /s    - Squish out all spaces/tabs")
        Console.WriteLine("        /q    - Quiet mode - no messages if files match")
        Console.WriteLine("        /c    - Show caret (^) at first mismatch char if one-line diff")
        Console.WriteLine("        /ca   - Show carets under all mismatched chars if one-line diff")
        Console.WriteLine("        /w    - Word compare")
        Console.WriteLine("        /t    - Token compare")
        ''Console.WriteLine("        /n    - Normalize numeric and date values")
    End Sub

End Module
