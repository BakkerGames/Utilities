' -----------------------------------
' --- FileVersion.vb - 11/10/2010 ---
' -----------------------------------

' ------------------------------------------------------------------------------------------
' 11/10/2010 - SBakker
'            - Fixed IncNewVersion to properly use NewVersion instead of Version. Oops!
'            - Removed IncVersion, as it wasn't used anywhere and shouldn't be.
' 10/22/2010 - SBakker
'            - Added AssemblyDateTime property, for comparing against file dates if there is
'              no file version (such as an embedded resource file).
' 06/30/2010 - SBakker
'            - Added new functions New/VersionMajorMinorBuild, to get the part of
'              the version number before the revision.
'            - Added New/IncRevision, to easily increment the revision number if
'              an included file's datetime is later.
' --------------------------------------------------------------------------------

Public Class FileVersion

#Region " --- Properties --- "

    Private _DidIncNewVersion As Boolean = False

    Private _FullName As String = ""
    Public Property FullName() As String
        Get
            Return _FullName
        End Get
        Set(ByVal value As String)
            _FullName = value
        End Set
    End Property

    Private _DirectoryName As String = ""
    Public Property DirectoryName() As String
        Get
            Return _DirectoryName
        End Get
        Set(ByVal value As String)
            _DirectoryName = value
        End Set
    End Property

    Private _AssemblyFullName As String = ""
    Public Property AssemblyFullName() As String
        Get
            Return _AssemblyFullName
        End Get
        Set(ByVal value As String)
            _AssemblyFullName = value
        End Set
    End Property

    Private _FileName As String = ""
    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal value As String)
            _FileName = value
        End Set
    End Property

    Private _OutputType As String = ""
    Public Property OutputType() As String
        Get
            Return _OutputType
        End Get
        Set(ByVal value As String)
            _OutputType = value
        End Set
    End Property

    Private _Version As String = ""
    Public Property Version() As String
        Get
            Return _Version
        End Get
        Set(ByVal value As String)
            If _Version <> CompressedVersion(value) Then
                _Version = CompressedVersion(value)
            End If
        End Set
    End Property

    Private _NewVersion As String = ""
    Public Property NewVersion() As String
        Get
            Return _NewVersion
        End Get
        Set(ByVal value As String)
            If _NewVersion <> CompressedVersion(value) Then
                _NewVersion = CompressedVersion(value)
                _DidIncNewVersion = False
            End If
        End Set
    End Property

    Private _Updated As Boolean = False
    Public Property Updated() As Boolean
        Get
            Return _Updated
        End Get
        Set(ByVal value As Boolean)
            _Updated = value
        End Set
    End Property

    Private _Level As Integer = 0
    Public Property Level() As Integer
        Get
            Return _Level
        End Get
        Set(ByVal value As Integer)
            _Level = value
        End Set
    End Property

    Private _IsExecutable As Boolean = False
    Public Property IsExecutable() As Boolean
        Get
            Return _IsExecutable
        End Get
        Set(ByVal value As Boolean)
            _IsExecutable = value
        End Set
    End Property

    Private _PublishPath As String = ""
    Public Property PublishPath() As String
        Get
            Return _PublishPath
        End Get
        Set(ByVal value As String)
            _PublishPath = value
        End Set
    End Property

    Public Property IsAltProject As Boolean = False

    Public Property AssemblyDateTime As Date = Nothing

#End Region

#Region " --- Public Functions --- "

    Public Function VersionMajorMinorBuild() As String
        Dim Result As String = CompressedVersion(_Version)
        Result = Result.Substring(0, Result.LastIndexOf("."c))
        Return Result
    End Function

    Public Function NewVersionMajorMinorBuild() As String
        Dim Result As String = CompressedVersion(_NewVersion)
        Result = Result.Substring(0, Result.LastIndexOf("."c))
        Return Result
    End Function

    Public Function VersionRevision() As String
        Dim Result As String = CompressedVersion(_Version)
        Result = Result.Substring(Result.LastIndexOf("."c) + 1)
        Return Result
    End Function

    Public Function NewVersionRevision() As String
        Dim Result As String = CompressedVersion(_NewVersion)
        Result = Result.Substring(Result.LastIndexOf("."c) + 1)
        Return Result
    End Function

    Public Function NewVersionLessThan(ByVal CompareVersion As String) As Boolean
        Return (ExpandedVersion(_NewVersion) < ExpandedVersion(CompareVersion))
    End Function

#End Region

#Region " --- Public Subroutines --- "

    Public Sub IncNewRevision()
        If Not _DidIncNewVersion Then
            Dim Revision As String = CompressedVersion(_NewVersion)
            Revision = Revision.Substring(Revision.LastIndexOf("."c) + 1)
            Revision = (CInt(Revision) + 1).ToString
            Me.NewVersion = NewVersionMajorMinorBuild() + "." + Revision
            _DidIncNewVersion = True
        End If
    End Sub

#End Region

#Region " --- Internal Routines --- "

    Public Shared Function ExpandedVersion(ByVal Version As String) As String
        Dim Result As String = ""
        Dim Parts() As String = Version.Split("."c)
        For Each Part As String In Parts
            If Result <> "" Then
                Result += "."
            End If
            If Part.Length <> 4 Then Part = Right("0000" + Part, 4)
            Result += Part
        Next
        Return Result
    End Function

    Public Shared Function ExpandedMajorMinorBuild(ByVal Version As String) As String
        Dim Result As String = ""
        Dim Parts() As String = Version.Split("."c)
        For CurrIndex As Integer = 0 To 2
            Dim Part As String = Parts(CurrIndex)
            If Part.Length <> 4 Then Part = Right("0000" + Part, 4)
            If Result <> "" Then
                Result += "."
            End If
            Result += Part
        Next
        Return Result
    End Function

    Public Shared Function CompressedVersion(ByVal Version As String) As String
        Dim Result As String = ""
        Dim Parts() As String = Version.Split("."c)
        For Each Part As String In Parts
            If Result <> "" Then
                Result += "."
            End If
            If Part = "%2a" Then Part = "0"
            Do While Part.Length > 1 AndAlso Part.StartsWith("0")
                Part = Right(Part, Part.Length - 1)
            Loop
            Result += Part
        Next
        Return Result
    End Function

#End Region

End Class
