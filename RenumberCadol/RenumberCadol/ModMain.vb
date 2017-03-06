' ----------------------------
' --- ModMain - 04/16/2014 ---
' ----------------------------

' ----------------------------------------------------------------------------------------------------
' 04/16/2014 - SBakker
'            - Find the proper Encoding when reading in Cadol programs, but always write as ASCII.
' 04/09/2014 - SBakker
'            - Catch badly formatted Cadol programs which cause negative indents. Usually will mean a
'              missing "THEN" or an extra block terminator.
'            - Switch internal modules to be in CadolSourceDataClass as Public Classes.
' 10/19/2010 - SBakker
'            - Added "/D <date>", "/U <user>", "/R", "/V", "/N", and "/?" command options.
' 09/15/2010 - SBakker
'            - Added graceful handling of errors while renumbering.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FileUtils
Imports CadolSourceDataClass
Imports System.IO
Imports System.Text

Module ModMain

    Private CurrDate As String = ""
    Private CurrUser As String = ""
    Private Reformat As Boolean = False
    Private Verbose As Boolean = False
    Private ChangeCount As Integer = 0
    Private ErrorCount As Integer = 0
    Private ChangeInfo As Boolean = True

    Sub Main()
        Dim TempPos As Integer
        Dim FromFile As String
        Dim FromPath As String
        Dim FileName As String
        Dim DoSubdirs As Boolean = False
        ' ------------------------------
        CurrDate = Format(Today, "MM/dd/yyyy")
        CurrUser = Environment.UserName.Substring(Environment.UserName.LastIndexOf("\"c) + 1).ToLower
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
                        Console.WriteLine("   /F        - Reformat source after renumbering")
                        Console.WriteLine("   /R        - Renumber/reformat source in all subdirectories")
                        Console.WriteLine("   /V        - Verbose - print each filename checked")
                        Console.WriteLine("   /D <date> - Specify date to use for change information (rather than TODAY)")
                        Console.WriteLine("   /U <user> - Specify username to use for change information (rather than current user)")
                        Console.WriteLine("   /N        - Don't update date/user information")
                        Exit Sub
                    Case "/F"
                        Reformat = True
                    Case "/R"
                        DoSubdirs = True
                    Case "/V"
                        Verbose = True
                    Case "/D"
                        ArgNum += 1
                        Try
                            CurrDate = My.Application.CommandLineArgs.Item(ArgNum).ToUpper
                        Catch ex As Exception
                            Console.WriteLine("Date not specified after /D parameter")
                            Exit Sub
                        End Try
                        Try
                            CurrDate = Format(CDate(CurrDate), "MM/dd/yyyy")
                        Catch ex As Exception
                            Console.WriteLine("Invalid date specified: " + My.Application.CommandLineArgs.Item(ArgNum))
                            Exit Sub
                        End Try
                    Case "/U"
                        ArgNum += 1
                        Try
                            CurrUser = My.Application.CommandLineArgs.Item(ArgNum).ToUpper.ToLower
                        Catch ex As Exception
                            Console.WriteLine("Username not specified after /U parameter")
                            Exit Sub
                        End Try
                    Case "/N"
                        ChangeInfo = False
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
                FromFile = "*.k"
            End If
            If FromFile.EndsWith(".*") OrElse FromFile.EndsWith(".?") Then
                FromFile = FromFile.Substring(0, FromFile.Length - 2) + ".k"
            End If
            If Not FromFile.ToLower.EndsWith(".k") Then
                Console.WriteLine("Only Cadol Program files (*.k) may be renumbered.")
                Exit Sub
            End If
            TempPos = FromFile.LastIndexOf("\")
            If TempPos >= 0 Then
                FromPath = FromFile.Substring(0, TempPos + 1)
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
                RenumberReformatFile(FromPath, FileName)
            Next
        Else
            DoAllFiles(My.Computer.FileSystem.CurrentDirectory)
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

    Private Sub DoAllFiles(ByVal FromPath As String)
        For Each FileName As String In Directory.GetFiles(FromPath, "*.k")
            RenumberReformatFile(FromPath, FileName)
        Next
        For Each DirName As String In Directory.GetDirectories(FromPath)
            DoAllFiles(DirName)
        Next
    End Sub

    Private Sub RenumberReformatFile(ByVal FromPath As String, ByVal FileName As String)
        Dim TempPos As Integer
        Dim TempPos2 As Integer
        Dim ProgBefore As String
        Dim ProgAfter As String
        Dim BaseFileName As String
        ' ------------------------
        If Verbose Then
            Console.WriteLine(FileName)
        End If
        BaseFileName = FileName.Substring(FileName.LastIndexOf("\") + 1)
        Dim CurrEncoding As Encoding = GetFileEncoding(FileName)
        ProgBefore = File.ReadAllText(FileName, CurrEncoding)
        Try
            ProgAfter = CadolRenumber.RenumberCadolProgram(ProgBefore)
        Catch ex As Exception
            Console.WriteLine("Error: " + FileName)
            Console.WriteLine("       " + ex.Message)
            ErrorCount += 1
            Exit Sub
        End Try
        ' --- check if also reformatting the program ---
        If Reformat Then
            Try
                ProgAfter = CadolReformat.ReformatCadolProgram(ProgAfter)
            Catch ex As Exception
                Console.WriteLine("Error: " + FileName)
                Console.WriteLine("       " + ex.Message)
                ErrorCount += 1
                Exit Sub
            End Try
        End If
        ' --- check if program changed ---
        If ProgAfter <> ProgBefore OrElse CurrEncoding IsNot Encoding.ASCII Then
            If (File.GetAttributes(FileName) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                Console.WriteLine("File can be changed but is read-only: " + FileName)
                ErrorCount += 1
                Exit Sub
            End If
            If ChangeInfo Then
                ' --- fix first line comment ---
                If ProgAfter.StartsWith("*") Then
                    TempPos = ProgAfter.IndexOf(vbCrLf)
                    ProgAfter = "* " + BaseFileName.ToLower + " - " + _
                                CurrDate + ProgAfter.Substring(TempPos)
                End If
                ' --- fix embedded datestamp ---
                TempPos = ProgAfter.IndexOf("if 1#1 display")
                If TempPos >= 0 Then
                    TempPos2 = ProgAfter.IndexOf(vbCrLf, TempPos + 12)
                    ProgAfter = ProgAfter.Substring(0, TempPos) + _
                                "if 1#1 display "" " + CurrDate + " - " + _
                                CurrUser + " """ + _
                                ProgAfter.Substring(TempPos2)
                End If
            End If
            ' --- output renumbered program ---
            File.WriteAllText(FileName, ProgAfter, Encoding.ASCII)
            Console.WriteLine(FileName + " - Done")
            ChangeCount += 1
        End If
    End Sub

End Module
