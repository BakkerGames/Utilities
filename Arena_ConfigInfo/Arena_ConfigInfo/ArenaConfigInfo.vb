' ---------------------------------------
' --- ArenaConfigInfo.vb - 10/27/2016 ---
' ---------------------------------------

' ----------------------------------------------------------------------------------------------------
' 10/27/2016 - Add TimeoutSeconds and TimeoutRetry settings, so that environments can be individually
'              configured without recompiling.
' 07/21/2016 - SBakker - URD 12917
'            - Changed path to Arena.xml to be relative, not absolute. Add in application path if
'              necessary.
'            - Added LetterStationaryDir property so paths don't need to be hardcoded in LTDLetters or
'              STDLetters.
'            - Added AltProgramPath property, to help with Bootstrapping.
' 06/20/2016 - SBakker
'            - Only add trailing "\" to ProgramPath and LaunchPath if they are not blank.
' 01/05/2016 - MJeyadarmar
'            - Added code to read LTDCM_EmailFrom, LTDCM_EmailTo parameters in Arena.xml(LoadSettings)
' 07/23/2014 - SBakker
'            - Moved BootStrap, LaunchPath, and ProgramPath up to be top-level items, from inside the
'              MainMenu level.
'            - Removed all MainMenu nodes and sub-nodes from Arena.xml
'            - Removed all ApplicationInfo handling, since it has no MainMenu info to read from.
' 03/07/2014 - SBakker
'            - Added AltConfigFilename so that both mapped drives and network paths can be specified.
' 03/05/2014 - Added ProgramPath to use for the source location of programs. The Bootstrap routines
'              will use it if a double-bounce is needed because a program is started locally instead
'              of from the ProgramPath.
' 11/04/2013 - SBakker
'            - Make sure to LoadSettings() if Bootstrap is the first property accessed.
' 10/24/2013 - SBakker
'            - Added handling of "%USERPROFILE%" in the LaunchPath. This allows programs to be run
'              from the user's personal area, instead of using the Click-Once area. It is recommended
'              to use the directory "%USERPROFILE%\Applications\<myappname_myenvironment>".
'            - Make sure LaunchPath ends with a "\", as other programs are expecting it.
' 10/03/2013 - MMiesburger
'            - Added 3 new Nodes and properties to support UserRequest application.
' 01/04/2013 - SBakker
'            - Added UseActiveDirectory parameter and Arena.xml option, for use on systems
'              which don't have Active Directory access.
' 11/22/2011 - SBakker
'            - Added Reporting/ReportServerURL property, for storing the path to the Report
'              Server.
' 08/02/2011 - SBakker
'            - Added Payment/PostResultsPath property, for storing text files containing the
'              results from the Claim Payment post.
' 07/29/2011 - MMiesburger
'            - Added ImageKey to handled Arena.xml ImageKey
' 05/20/2011 - SBakker
'            - Added ability for an entire subgroup to be hidden. This is very useful for a
'              group of applications which don't appear on the menu, but can be called using
'              Click-Once.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 11/09/2010 - SBakker
'            - Added a Hidden property, to handle ClickOnce items which are never called
'              directly from the Main Menu, only from inside another program.
'            - Only add non-Hidden items to GetTreeList. Only add a group if it contains
'              some non-Hidden items, so a group of only Hidden items can be created.
' 06/23/2010 - SBakker
'            - If configuration file is missing, look for it on the C: drive.
' 11/09/2009 - SBakker
'            - Added name (if found) to APPGROUP records.
' 09/03/2009 - SBakker
'            - Built new class for getting configuration information from the
'              Arena.xml file.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Environment
Imports System.Xml
Imports System.Windows.Forms

Public Class ArenaConfigInfo

#Region " Private Constants and Variables "

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private Shared _InfoLoaded As Boolean = False

    Private Shared _ConnList As New List(Of ConnectionInfo)
    Private Shared _TreeList As New TreeNode

#End Region

#Region " Properties "

    Private Shared _Environment As String = ""
    Public Shared ReadOnly Property Environment() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _Environment
        End Get
    End Property

    Private Shared _UseActiveDirectory As Boolean = True
    Public Shared ReadOnly Property UseActiveDirectory() As Boolean
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _UseActiveDirectory
        End Get
    End Property

    Private Shared _LaunchPath As String = ""
    Public Shared ReadOnly Property LaunchPath() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _LaunchPath
        End Get
    End Property

    Private Shared _ProgramPath As String = ""
    Public Shared ReadOnly Property ProgramPath() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _ProgramPath
        End Get
    End Property

    Private Shared _AltProgramPath As String = ""
    Public Shared ReadOnly Property AltProgramPath() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _AltProgramPath
        End Get
    End Property

    Private Shared _PostResultsPath As String = ""
    Public Shared ReadOnly Property PostResultsPath As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _PostResultsPath
        End Get
    End Property

    Private Shared _ReportServerURL As String = ""
    Public Shared ReadOnly Property ReportServerURL As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _ReportServerURL
        End Get
    End Property

    Private Shared _Domain As String = ""
    Public Shared ReadOnly Property Domain() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _Domain
        End Get
    End Property

    Private Shared _AttachedFilesDirectory As String = ""
    Public Shared ReadOnly Property AttachedFilesDirectory() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _AttachedFilesDirectory
        End Get
    End Property

    Private Shared _RemovedFilesDirectory As String = ""
    Public Shared ReadOnly Property RemovedFilesDirectory() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _RemovedFilesDirectory
        End Get
    End Property

    Private Shared _Bootstrap As Boolean = False
    Public Shared ReadOnly Property Bootstrap As Boolean
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _Bootstrap
        End Get
    End Property


    Private Shared _LTDCM_EmailFrom As String = ""
    Public Shared ReadOnly Property LTDCM_EmailFrom() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _LTDCM_EmailFrom
        End Get
    End Property

    Private Shared _LTDCM_EmailTo As String = ""
    Public Shared ReadOnly Property LTDCM_EmailTo() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _LTDCM_EmailTo
        End Get
    End Property

    Private Shared _LetterStationaryDir As String = ""
    Public Shared ReadOnly Property LetterStationaryDir() As String
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _LetterStationaryDir
        End Get
    End Property

    Private Shared _TimeoutSeconds As Integer = 60
    Public Shared ReadOnly Property TimeoutSeconds As Integer
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _TimeoutSeconds
        End Get
    End Property

    Private Shared _TimeoutRetry As Integer = 0
    Public Shared ReadOnly Property TimeoutRetry As Integer
        Get
            If Not _InfoLoaded Then
                LoadSettings()
            End If
            Return _TimeoutRetry
        End Get
    End Property

#End Region

#Region " Public Routines "

    Public Shared Function GetConnectionInfo(ByVal ConnName As String) As ConnectionInfo
        If Not _InfoLoaded Then
            LoadSettings()
        End If
        For Each TempConn As ConnectionInfo In _ConnList
            If String.Equals(TempConn.Name, ConnName, StringComparison.CurrentCultureIgnoreCase) Then
                Return TempConn
            End If
        Next
        Return Nothing
    End Function

    Public Shared Function GetConnectionList() As List(Of ConnectionInfo)
        If Not _InfoLoaded Then
            LoadSettings()
        End If
        Return _ConnList
    End Function

    Public Shared Function GetTreeList() As TreeNode
        If Not _InfoLoaded Then
            LoadSettings()
        End If
        Return _TreeList
    End Function

#End Region

#Region " Private Routines "

    Private Shared Sub LoadSettings()
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim Filename As String = My.Settings.ConfigFilename
        If String.IsNullOrWhiteSpace(Filename) Then
            Throw New SystemException(FuncName + vbCrLf + "Arena configuration file not specified.")
        End If
        ' --- Check for an alternate location ---
        If Not File.Exists(Filename) AndAlso Filename.StartsWith(".\") Then
            Filename = $"{My.Application.Info.DirectoryPath}\{Filename}"
        End If
        If Not File.Exists(Filename) AndAlso Not String.IsNullOrWhiteSpace(My.Settings.AltConfigFilename) Then
            Filename = My.Settings.AltConfigFilename
            If Not File.Exists(Filename) AndAlso Filename.StartsWith(".\") Then
                Filename = $"{My.Application.Info.DirectoryPath}\{Filename}"
            End If
        End If
        ' --- Check for configuration document ---
        If Not File.Exists(Filename) Then
            Throw New SystemException(FuncName + vbCrLf + "Arena configuration file not found: " + Filename)
        End If
        ' --- Load the Arena.xml configuration document ---
        Dim ArenaConfig As New XmlDocument
        ArenaConfig.Load(Filename)
        ' --- Get the Root node ---
        Dim Root As XmlNode = ArenaConfig.DocumentElement
        If Root.Name.ToUpper <> "ARENA" Then
            Throw New SystemException(FuncName + vbCrLf + "Invalid Arena configuration file")
        End If
        ' --- Check through the child nodes for useful info ---
        For Each ConfigNode As XmlNode In Root
            Select Case ConfigNode.Name.ToUpper
                Case "ENVIRONMENT"
                    _Environment = ConfigNode.InnerText
                Case "DOMAIN"
                    _Domain = ConfigNode.InnerText
                Case "ATTACHEDFILESDIRECTORY"
                    _AttachedFilesDirectory = ConfigNode.InnerText
                Case "REMOVEDFILESDIRECTORY"
                    _RemovedFilesDirectory = ConfigNode.InnerText
                Case "USEACTIVEDIRECTORY"
                    _UseActiveDirectory = (ConfigNode.InnerText.ToUpper <> "FALSE")
                Case "DATACONN"
                    For Each DCNode As XmlNode In ConfigNode
                        Dim TempConn As ConnectionInfo = GetConnData(DCNode)
                        _ConnList.Add(TempConn)
                    Next
                Case "BOOTSTRAP"
                    _Bootstrap = (ConfigNode.InnerText.ToUpper = "TRUE")
                Case "LAUNCHPATH"
                    _LaunchPath = ConfigNode.InnerText
                    If _LaunchPath.ToUpper.StartsWith("%USERPROFILE%") Then
                        _LaunchPath = GetEnvironmentVariable("USERPROFILE") + _LaunchPath.Substring(Len("%USERPROFILE%"))
                    End If
                    If Not String.IsNullOrEmpty(_LaunchPath) AndAlso Not _LaunchPath.EndsWith("\") Then
                        _LaunchPath += "\"
                    End If
                Case "PROGRAMPATH"
                    _ProgramPath = ConfigNode.InnerText
                    If Not String.IsNullOrEmpty(_ProgramPath) AndAlso Not _ProgramPath.EndsWith("\") Then
                        _ProgramPath += "\"
                    End If
                Case "ALTPROGRAMPATH"
                    _AltProgramPath = ConfigNode.InnerText
                    If Not String.IsNullOrEmpty(_AltProgramPath) AndAlso Not _AltProgramPath.EndsWith("\") Then
                        _AltProgramPath += "\"
                    End If
                Case "PAYMENT"
                    For Each PmtNode As XmlNode In ConfigNode
                        Select Case PmtNode.Name.ToUpper
                            Case "POSTRESULTSPATH"
                                _PostResultsPath = PmtNode.InnerText
                        End Select
                    Next
                Case "REPORTING"
                    For Each PmtNode As XmlNode In ConfigNode
                        Select Case PmtNode.Name.ToUpper
                            Case "REPORTSERVERURL"
                                _ReportServerURL = PmtNode.InnerText
                        End Select
                    Next
                Case "LTDCM_EMAILFROM"
                    _LTDCM_EmailFrom = ConfigNode.InnerText
                Case "LTDCM_EMAILTO"
                    _LTDCM_EmailTo = ConfigNode.InnerText
                Case "LETTERSTATIONARYDIR"
                    _LetterStationaryDir = ConfigNode.InnerText
                Case "TIMEOUTSECONDS"
                    _TimeoutSeconds = CInt(ConfigNode.InnerText)
                Case "TIMEOUTRETRY"
                    _TimeoutRetry = CInt(ConfigNode.InnerText)
            End Select
        Next
        _InfoLoaded = True
    End Sub

    Private Shared Function GetConnData(ByVal XNode As XmlNode) As ConnectionInfo
        Dim TempConn As New ConnectionInfo
        TempConn.Name = XNode.Name
        For Each SubNode As XmlNode In XNode
            Select Case SubNode.Name.ToUpper
                Case "SERVER"
                    TempConn.Server = SubNode.InnerText
                Case "DATABASE"
                    TempConn.Database = SubNode.InnerText
            End Select
        Next
        Return TempConn
    End Function

#End Region

End Class
