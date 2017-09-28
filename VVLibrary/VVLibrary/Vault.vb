' -----------------------------
' --- Vault.vb - 09/28/2017 ---
' -----------------------------

Imports System.IO
Imports Arena.Common.JSON

' ----------------------------------------------------------------------------------------------------
' This class handles backing up a directory and its subdirectories to a Vault directory.
' A Vault directory is merely a directory tree matching the source directory tree, with an additional
'     level so that each source files becomes a directory with the same name.
' Within each file's directory are stored all of the history for the file.
' History files are named with their modification UTC date and time, to the second.
' The history files have the file's extension added, so they may be viewed directly using the proper
'     application.
' History files contain the full copy of the file, not a differential. This allows history files to be
'     opened and viewed directly without needing to use a "reconstruction" routine. It also avoid any
'     problems with encoding and line separators, and allows text and binary files to be handled the
'     same way. In addition, VersionVaults are "additive", so once a file is added it is never changed
'     or removed.
' To compare files, the Last Change UTCs are compared to see if they are exactly the same. If not,
'     they are different. No other information is checked.
' In the source directory, a file named ".vvconfig" must exist. It is a JSON Object file. It must
'     have a "VVPath" value for the default VersionVault location, an "IgnoreDir" array for directories
'     to be ignored, and an "IgnoreExt" array for file endings to be ignored.
' The "VVPath" setting in the .vvconfig can be overridden by specifying the VaultPath parameter below.
' All hidden files and directories, as well as files and directories starting with ".", are
'     automatically ignored and don't have to be included in the ".vvconfig" file.
' ----------------------------------------------------------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/28/2017 - SBakker
'            - Updated to Arena.Common.JSON.
' 02/22/2017 - SBakker
'            - Removed support for .vvignore file. Use .vvconfig file instead now.
' ----------------------------------------------------------------------------------------------------

Public Class Vault

    Private Const DateTimePattern As String = "yyyyMMdd_HHmmss"
    Private Const ConfigFileName As String = ".vvconfig"
    Private vvConfig As JObject

    Private IgnoreSpecificationList As List(Of String)

    Public Sub New()
    End Sub

    Public Sub New(ByVal SourcePath As String, ByVal VaultPath As String)
        Me.SourcePath = SourcePath
        Me.VaultPath = VaultPath
    End Sub

    Private _SourcePath As String = ""
    Public Property SourcePath As String
        Get
            Return _SourcePath
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(value)
            End If
            If value.EndsWith("\"c) Then
                value = value.Substring(0, value.Length - 1)
            End If
            If Not Directory.Exists(value) Then
                Throw New DirectoryNotFoundException(value)
            End If
            If Not File.Exists($"{value}\{ConfigFileName}") Then
                Throw New SystemException($"{value}\{ConfigFileName} not found")
            End If
            vvConfig = JObject.Parse(File.ReadAllText($"{value}\{ConfigFileName}"))
            _SourcePath = value
            ' --- Use path in config file for now, could be overridden ---
            VaultPath = vvConfig.GetValueOrNull("VVPath").ToString
        End Set
    End Property

    Private _VaultPath As String = ""
    Public Property VaultPath As String
        Get
            Return _VaultPath
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(value)
            End If
            If value.EndsWith("\"c) Then
                value = value.Substring(0, value.Length - 1)
            End If
            If Not Directory.Exists(value) Then
                Throw New DirectoryNotFoundException(value)
            End If
            _VaultPath = value
        End Set
    End Property

End Class
