' -------------------------------
' --- ModMain.vb - 11/27/2017 ---
' -------------------------------

' ----------------------------------------------------------------------------------------------------
' 11/27/2017 - SBakker
'            - Display all errors on Console. Don't throw any.
' 09/28/2017 - SBakker
'            - Updated to Arena.Common.JSON.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports VVLibrary
Imports Arena.Common.JSON

Module ModMain

    Private WithEvents MyVault As New Vault

    Private StartSourcePath As String
    Private StartVaultPath As String

    Private Const ConfigFileName As String = ".vvconfig"
    Private vvConfig As JObject

    Private BackupCount As Integer

    Private SourceNewerCount As Integer
    Private VaultNewerCount As Integer
    Private SourceOnlyCount As Integer
    Private VaultOnlyCount As Integer

    Sub Main()

        Dim CompareFlag As Boolean = False
        Dim ShowTimeFlag As Boolean = False
        Dim UseConfigFile As String = ""

        For ArgNum As Integer = 0 To My.Application.CommandLineArgs.Count - 1
            If My.Application.CommandLineArgs.Item(ArgNum).StartsWith("/") Then
                If My.Application.CommandLineArgs.Item(ArgNum).ToUpper = "/?" Then
                    ShowSyntax()
                    Exit Sub
                ElseIf My.Application.CommandLineArgs.Item(ArgNum).ToUpper = "/C" OrElse
                       My.Application.CommandLineArgs.Item(ArgNum).ToUpper = "/COMPARE" Then
                    CompareFlag = True
                    ''ElseIf My.Application.CommandLineArgs.Item(ArgNum).ToUpper = "/T" OrElse
                    ''       My.Application.CommandLineArgs.Item(ArgNum).ToUpper = "/TIME" Then
                    ''    ShowTimeFlag = True
                ElseIf My.Application.CommandLineArgs.Item(ArgNum).ToUpper.StartsWith("/CONFIG:") Then
                    Dim TempPath As String = My.Application.CommandLineArgs.Item(ArgNum).Substring(3)
                    If TempPath.StartsWith("""") AndAlso TempPath.EndsWith("""") Then
                        TempPath = TempPath.Substring(1, TempPath.Length - 2)
                    End If
                    If Not File.Exists(TempPath) Then
                        Console.WriteLine($"Error: Config file not found: {TempPath}")
                        Exit Sub
                    End If
                    UseConfigFile = TempPath
                Else
                    Console.WriteLine($"Unknown option specified: {My.Application.CommandLineArgs.Item(ArgNum)}")
                    Exit Sub
                End If
            ElseIf String.IsNullOrEmpty(StartSourcePath) Then
                StartSourcePath = My.Application.CommandLineArgs.Item(ArgNum)
            ElseIf String.IsNullOrEmpty(StartVaultPath) Then
                StartVaultPath = My.Application.CommandLineArgs.Item(ArgNum)
            Else
                Console.WriteLine($"Unknown option specified: {My.Application.CommandLineArgs.Item(ArgNum)}")
                Exit Sub
            End If
        Next

        If String.IsNullOrEmpty(StartSourcePath) Then
            ShowSyntax()
            Exit Sub
        End If

        If String.IsNullOrEmpty(UseConfigFile) Then
            UseConfigFile = $"{StartSourcePath}\{ConfigFileName}"
        End If

        If Not File.Exists(UseConfigFile) Then
            Console.WriteLine($"{UseConfigFile} not found")
            Exit Sub
        End If
        vvConfig = JObject.Parse(File.ReadAllText($"{UseConfigFile}"))

        If String.IsNullOrEmpty(StartVaultPath) Then
            StartVaultPath = vvConfig.GetValueOrNull("VVPath").ToString
        End If

        Try
            If Not Directory.Exists(StartVaultPath) Then
                Directory.CreateDirectory(StartVaultPath)
            End If
        Catch ex As Exception
            Console.WriteLine($"Error: Cannot create path: {StartVaultPath}")
            Exit Sub
        End Try

        Try
            MyVault.SourcePath = StartSourcePath
        Catch ex As Exception
            Console.WriteLine($"Error assigning path: {StartSourcePath}")
            Exit Sub
        End Try

        Try
            MyVault.VaultPath = StartVaultPath
        Catch ex As Exception
            Console.WriteLine($"Error assigning path: {StartVaultPath}")
            Exit Sub
        End Try

        Console.WriteLine($"From : {MyVault.SourcePath}")
        Console.WriteLine($"Vault: {MyVault.VaultPath}")
        Console.WriteLine()

        If ShowTimeFlag Then
            Console.WriteLine(Now.ToString)
        End If

        Try
            If CompareFlag Then
                SourceNewerCount = 0
                VaultNewerCount = 0
                SourceOnlyCount = 0
                VaultOnlyCount = 0
                MyVault.CompareAll()
                If SourceNewerCount + SourceOnlyCount + VaultNewerCount + VaultOnlyCount > 0 Then
                    Console.WriteLine()
                End If
                Console.Write("Source newer: ")
                Console.WriteLine(ShowFileCount(SourceNewerCount))
                Console.Write("Source only : ")
                Console.WriteLine(ShowFileCount(SourceOnlyCount))
                Console.Write("Vault newer : ")
                Console.WriteLine(ShowFileCount(VaultNewerCount))
                Console.Write("Vault only  : ")
                Console.WriteLine(ShowFileCount(VaultOnlyCount))
            Else
                BackupCount = 0
                MyVault.BackupAll()
                If BackupCount > 0 Then
                    Console.WriteLine()
                End If
                Console.Write("Stored: ")
                Console.WriteLine(ShowFileCount(BackupCount))
            End If
        Catch ex As Exception
            Console.WriteLine($"Error: {ex.Message}")
        End Try

        If ShowTimeFlag Then
            Console.WriteLine(Now.ToString)
        End If

#If DEBUG Then
        Console.ReadLine()
#End If

    End Sub

    Private Sub ShowSyntax()
        Console.WriteLine($"{My.Application.Info.AssemblyName} - {My.Application.Info.Version.ToString}")
        Console.WriteLine()
        Console.WriteLine("Syntax:")
        Console.WriteLine($"{My.Application.Info.AssemblyName} [options] <sourcedir> [<vaultdir>]")
        Console.WriteLine($"{My.Application.Info.AssemblyName} /C{{OMPARE}} [options] <sourcedir> [<vaultdir>]")
        Console.WriteLine()
        Console.WriteLine("Options:")
        Console.WriteLine("    /CONFIG:<vvconfigpath>")
    End Sub

    Private Sub Handle_FileBackedUp(ByVal Filename As String) Handles MyVault.FileVaulted
        BackupCount += 1
        Console.WriteLine(Filename)
    End Sub

    Private Sub Handle_SourceFileIsNewer(ByVal Filename As String) Handles MyVault.SourceFileIsNewer
        SourceNewerCount += 1
        Console.WriteLine($"Source Newer: {Filename.Substring(StartSourcePath.Length + 1)}")
    End Sub
    Private Sub Handle_VaultFileIsNewer(ByVal Filename As String) Handles MyVault.VaultFileIsNewer
        VaultNewerCount += 1
        Console.WriteLine($"Vault Newer : {Filename.Substring(StartSourcePath.Length + 1)}")
    End Sub

    Private Sub Handle_FileNotInVault(ByVal Filename As String) Handles MyVault.FileNotInVault
        SourceOnlyCount += 1
        Console.WriteLine($"Source Only : {Filename.Substring(StartSourcePath.Length + 1)}")
    End Sub

    Private Sub Handle_FileOnlyInVault(ByVal Filename As String) Handles MyVault.FileOnlyInVault
        VaultOnlyCount += 1
        Console.WriteLine($"Vault Only  : {Filename.Substring(StartVaultPath.Length + 1)}")
    End Sub

    Private Sub Handle_AccessError(ByVal Message As String) Handles MyVault.AccessError
        Console.WriteLine($"ERROR! {Message}")
    End Sub

    Private Function ShowFileCount(ByVal Count As Integer) As String
        If Count = 1 Then
            Return ("1 file")
        Else
            Return ($"{Count} files")
        End If
    End Function

End Module
