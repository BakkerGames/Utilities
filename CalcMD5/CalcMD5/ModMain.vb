' -------------------------------
' --- ModMain.vb - 09/19/2013 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/19/2013 - SBakker
'            - Added "/F" option to add two spaces and the filename after the MD5 hashcode. This is the
'              standard for MD5 files to be opened by an MD5SUM program (such as QuckPar).
' 09/16/2013 - SBakker
'            - Added new option "/E" to preserve the original extention and add ".MD5" to the end of
'              the new file.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Module ModMain

    Private QuietMode As Boolean = False
    Private AddMode As Boolean = False
    Private UpdateMode As Boolean = False
    Private HideMD5 As Boolean = False
    Private ChangeCount As Integer = 0
    Private ErrorCount As Integer = 0
    Private AddExtensionToEnd As Boolean = False
    Private IncludeFilename As Boolean = False

    Sub Main()
        Dim TempPos As Integer
        Dim FromFile As String
        Dim FromPath As String
        Dim FileName As String
        Dim DoSubdirs As Boolean = False
        ' ------------------------------
        FromFile = ""
        FromPath = My.Computer.FileSystem.CurrentDirectory + "\"
        For ArgNum As Integer = 0 To My.Application.CommandLineArgs.Count - 1
            If My.Application.CommandLineArgs.Item(ArgNum).StartsWith("/") Then
                Select Case My.Application.CommandLineArgs.Item(ArgNum).ToUpper
                    Case "/?"
                        Console.WriteLine(My.Application.Info.AssemblyName + " - " + My.Application.Info.Version.ToString)
                        Console.WriteLine()
                        Console.WriteLine(My.Application.Info.AssemblyName + " <options> <filename>")
                        Console.WriteLine()
                        Console.WriteLine("   /R        - Recursively process all subdirectories")
                        Console.WriteLine("   /Q        - Quiet Mode - Only display errors")
                        Console.WriteLine("   /A        - Only add missing MD5 files")
                        Console.WriteLine("   /U        - Update Mode - Show errors and update MD5 hash")
                        Console.WriteLine("   /H        - Hide MD5 files")
                        Console.WriteLine("   /E        - Keep file extension and add "".MD5"" to the end")
                        Console.WriteLine("   /F        - Include Filename after MD5 value")
                        Exit Sub
                    Case "/R"
                        DoSubdirs = True
                    Case "/Q"
                        QuietMode = True
                    Case "/A"
                        AddMode = True
                    Case "/U"
                        UpdateMode = True
                    Case "/H"
                        HideMD5 = True
                    Case "/E"
                        AddExtensionToEnd = True
                    Case "/F"
                        IncludeFilename = True
                    Case Else
                        Console.WriteLine("Unknown option specified: " + My.Application.CommandLineArgs.Item(ArgNum))
                        Exit Sub
                End Select
            Else
                FromFile = My.Application.CommandLineArgs.Item(ArgNum)
            End If
        Next
        If Not DoSubdirs Then
            If FromFile = "" Then
                FromFile = "*.*"
            End If
            TempPos = FromFile.LastIndexOf("\")
            If TempPos >= 0 Then
                FromPath = FromFile.Substring(0, TempPos)
                FromFile = FromFile.Substring(TempPos + 1)
            End If
            If (Not FromPath.StartsWith("\")) AndAlso (Not FromPath.StartsWith(".")) AndAlso (Not FromPath.Substring(1, 1) = ":") Then
                FromPath = My.Computer.FileSystem.CurrentDirectory + "\" + FromPath
            End If
            If FromPath.IndexOf("*"c) >= 0 OrElse FromPath.IndexOf("?"c) >= 0 Then
                Console.WriteLine("Path may not contain wildcards.")
                Exit Sub
            End If
            For Each FileName In Directory.GetFiles(FromPath, FromFile)
                CalcMD5Hash(FileName)
            Next
        ElseIf FromFile = "" Then
            DoAllFiles(My.Computer.FileSystem.CurrentDirectory, "*.*")
        Else
            Dim SplitPos As Integer = FromFile.LastIndexOf("\")
            If SplitPos < 0 Then
                DoAllFiles(My.Computer.FileSystem.CurrentDirectory, FromFile)
            ElseIf Directory.Exists(FromFile) Then
                DoAllFiles(FromFile, "*.*")
            Else
                FromPath = FromFile.Substring(0, SplitPos)
                FromFile = FromFile.Substring(SplitPos + 1)
                DoAllFiles(FromPath, FromFile)
            End If
        End If
        ' --- Output results ---
        If ErrorCount = 1 Then
            Console.WriteLine("1 error found.")
        ElseIf ErrorCount > 0 Then
            Console.WriteLine(ErrorCount.ToString + " errors found.")
        End If
        If ChangeCount = 0 Then
            Console.WriteLine("No files changed.")
        ElseIf ChangeCount = 1 Then
            Console.WriteLine("1 file changed.")
        Else
            Console.WriteLine(ChangeCount.ToString + " files changed.")
        End If
        Console.WriteLine("*** Done ***")
    End Sub

    Private Sub DoAllFiles(ByVal FromPath As String, ByVal FileSpec As String)
        For Each FileName As String In Directory.GetFiles(FromPath, FileSpec)
            CalcMD5Hash(FileName)
        Next
        For Each DirName As String In Directory.GetDirectories(FromPath)
            DoAllFiles(DirName, FileSpec)
        Next
    End Sub

    Public Sub CalcMD5Hash(ByVal Filename As String)
        If Filename.ToLower.EndsWith(".md5") Then Exit Sub
        If Not File.Exists(Filename) Then
            Console.WriteLine("File not found: " + Filename)
            ErrorCount += 1
            Exit Sub
        End If
        Dim MD5Result As String
        Dim ExtPos As Integer = Filename.LastIndexOf("."c)
        If ExtPos < 0 Then ExtPos = Filename.Length
        Dim MD5Filename As String
        If Not AddExtensionToEnd Then
            MD5Filename = Filename.Substring(0, ExtPos) + ".md5"
        Else
            MD5Filename = Filename + ".md5"
        End If
        If AddMode AndAlso File.Exists(MD5Filename) Then
            Exit Sub
        End If
        Try
            Dim MD5Hasher As MD5 = MD5.Create
            Dim fs As FileStream = File.OpenRead(Filename)
            Dim Result() As Byte = MD5Hasher.ComputeHash(fs)
            fs.Close()
            Dim HexResult As New StringBuilder
            For Each b As Byte In Result
                HexResult.Append(b.ToString("x2"))
            Next
            MD5Result = HexResult.ToString
        Catch ex As Exception
            Console.WriteLine("Error processing file: " + Filename)
            Console.WriteLine(ex.Message)
            ErrorCount += 1
            Exit Sub
        End Try
        If File.Exists(MD5Filename) Then
            Dim OldMD5 As String = File.ReadAllText(MD5Filename)
            If OldMD5.Length > 32 Then
                OldMD5 = Left(OldMD5, 32)
            End If
            If OldMD5 = MD5Result Then
                If Not QuietMode Then
                    Console.WriteLine(Filename + " - Match")
                End If
            Else
                Console.WriteLine(Filename + " - Error!")
                Console.WriteLine("   Old MD5 = " + OldMD5 + ", New MD5 = " + MD5Result)
                ErrorCount += 1
                If UpdateMode Then
                    Try
                        File.SetAttributes(MD5Filename, FileAttributes.Normal)
                        File.Delete(MD5Filename)
                        File.WriteAllText(MD5Filename, MD5Result)
                        If HideMD5 Then
                            File.SetAttributes(MD5Filename, FileAttributes.Hidden)
                        End If
                        ChangeCount += 1
                    Catch ex As Exception
                        Console.WriteLine("Error updating file: " + MD5Filename)
                        Console.WriteLine(ex.Message)
                        ErrorCount += 1
                        Exit Sub
                    End Try
                End If
            End If
        Else
            If IncludeFilename Then
                Dim BaseFilename As String = Filename.Substring(Filename.LastIndexOf("\"c) + 1)
                File.WriteAllText(MD5Filename, MD5Result + "  " + BaseFilename)
            Else
                File.WriteAllText(MD5Filename, MD5Result)
            End If
            If HideMD5 Then
                File.SetAttributes(MD5Filename, FileAttributes.Hidden)
            End If
            ChangeCount += 1
            If Not QuietMode Then
                Console.WriteLine(Filename + " - Done")
            End If
        End If
    End Sub

End Module
