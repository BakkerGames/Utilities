' -------------------------------
' --- ModMain.vb - 05/15/2015 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' 05/15/2015 - SBakker
'            - Copied CalcMD5 to CalcSHA1 program.
' 09/19/2013 - SBakker
'            - Added "/F" option to add two spaces and the filename after the SHA1 hashcode. This is the
'              standard for SHA1 files to be opened by an SHA1SUM program (such as QuckPar).
' 09/16/2013 - SBakker
'            - Added new option "/E" to preserve the original extention and add ".sha1" to the end of
'              the new file.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Module ModMain

    Private QuietMode As Boolean = False
    Private AddMode As Boolean = False
    Private UpdateMode As Boolean = False
    Private HideSHA1 As Boolean = False
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
                        Console.WriteLine("   /A        - Only add missing SHA1 files")
                        Console.WriteLine("   /U        - Update Mode - Show errors and update SHA1 hash")
                        Console.WriteLine("   /H        - Hide SHA1 files")
                        Console.WriteLine("   /E        - Keep file extension and add "".sha1"" to the end")
                        Console.WriteLine("   /F        - Include Filename after SHA1 value")
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
                        HideSHA1 = True
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
                CalcSHA1Hash(FileName)
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
            CalcSHA1Hash(FileName)
        Next
        For Each DirName As String In Directory.GetDirectories(FromPath)
            DoAllFiles(DirName, FileSpec)
        Next
    End Sub

    Public Sub CalcSHA1Hash(ByVal Filename As String)
        If Filename.ToLower.EndsWith(".sha1") Then Exit Sub
        If Not File.Exists(Filename) Then
            Console.WriteLine("File not found: " + Filename)
            ErrorCount += 1
            Exit Sub
        End If
        Dim SHA1Result As String
        Dim ExtPos As Integer = Filename.LastIndexOf("."c)
        If ExtPos < 0 Then ExtPos = Filename.Length
        Dim SHA1Filename As String
        If Not AddExtensionToEnd Then
            SHA1Filename = Filename.Substring(0, ExtPos) + ".sha1"
        Else
            SHA1Filename = Filename + ".sha1"
        End If
        If AddMode AndAlso File.Exists(SHA1Filename) Then
            Exit Sub
        End If
        Try
            Dim SHA1Hasher As SHA1 = SHA1.Create
            Dim fs As FileStream = File.OpenRead(Filename)
            Dim Result() As Byte = SHA1Hasher.ComputeHash(fs)
            fs.Close()
            Dim HexResult As New StringBuilder
            For Each b As Byte In Result
                HexResult.Append(b.ToString("x2"))
            Next
            SHA1Result = HexResult.ToString
        Catch ex As Exception
            Console.WriteLine("Error processing file: " + Filename)
            Console.WriteLine(ex.Message)
            ErrorCount += 1
            Exit Sub
        End Try
        If File.Exists(SHA1Filename) Then
            Dim OldSHA1 As String = File.ReadAllText(SHA1Filename)
            If OldSHA1.Length > 40 Then
                OldSHA1 = Left(OldSHA1, 40)
            End If
            If OldSHA1 = SHA1Result Then
                If Not QuietMode Then
                    Console.WriteLine(Filename + " - Match")
                End If
            Else
                Console.WriteLine(Filename + " - Error!")
                Console.WriteLine("   Old SHA1 = " + OldSHA1 + ", New SHA1 = " + SHA1Result)
                ErrorCount += 1
                If UpdateMode Then
                    Try
                        File.SetAttributes(SHA1Filename, FileAttributes.Normal)
                        File.Delete(SHA1Filename)
                        File.WriteAllText(SHA1Filename, SHA1Result)
                        If HideSHA1 Then
                            File.SetAttributes(SHA1Filename, FileAttributes.Hidden)
                        End If
                        ChangeCount += 1
                    Catch ex As Exception
                        Console.WriteLine("Error updating file: " + SHA1Filename)
                        Console.WriteLine(ex.Message)
                        ErrorCount += 1
                        Exit Sub
                    End Try
                End If
            End If
        Else
            If IncludeFilename Then
                Dim BaseFilename As String = Filename.Substring(Filename.LastIndexOf("\"c) + 1)
                File.WriteAllText(SHA1Filename, SHA1Result + "  " + BaseFilename)
            Else
                File.WriteAllText(SHA1Filename, SHA1Result)
            End If
            If HideSHA1 Then
                File.SetAttributes(SHA1Filename, FileAttributes.Hidden)
            End If
            ChangeCount += 1
            If Not QuietMode Then
                Console.WriteLine(Filename + " - Done")
            End If
        End If
    End Sub

End Module
