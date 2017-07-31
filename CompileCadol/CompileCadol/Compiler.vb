' --------------------------------
' --- Compiler.vb - 07/27/2017 ---
' --------------------------------

' ------------------------------------------------------------------------------------------
' 07/27/2017 - SBakker
'            - Enhanced "RENAME" logic. It should prevent the original name from being used
'              and prevent duplicate aliases. This is the original CADOL 3 compiler logic.
'            - Fill Renames dictionary with all possible rename variables. That prevents any
'              invalid "RENAME" statements and allows duplicate alias checking.
' 07/26/2016 - SBakker
'            - Convert "ENTER" and "EDIT" to specific "ENTERALPHA", "ENTERNUM", "EDITALPHA",
'              "EDITNUM" in third pass.
' 01/25/2016 - SBakker
'            - Changed all "And" to "AndAlso", "Or" to "OrElse".
' 12/07/2012 - SBakker
'            - Added new commands SAVEFILEINFO and RESTOREFILEINFO.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 09/15/2010 - SBakker
'            - Added error checking around having duplicate line numbers in a source file.
' 09/09/2009 - SBakker
'            - Turned on Option Strict, and fixed all type conversion issues.
'            - Changed MajorVer, MinorVer, RevisionVer to be Year, Month, Day.
' 10/08/2008 - SBAKKER - URD 11164
'            - Finally switched "%" to "_". Tired of having SourceSafe issues.
' 10/02/2008 - SBAKKER - URD 11164
'            - Added support for running Claims for specified clients only.
'            - Added support for read-only clients.
' 09/15/2008 - SBAKKER
'            - Check for existence of include file before trying to access it.
'            - Added an error for references to non-existent line numbers.
' 07/11/2008 - SBAKKER
'            - Made Compiler class a COM Interop assembly so that it may be
'              called from VB6 programs.
' 06/11/2008 - SBAKKER
'            - "Return True" from CompileFourthPass instead of "Exit Function".
'              Was always returning False, which caused failures.
' 08/24/2007 - SBAKKER
'            - Added error checking that will show the proper error when an
'              expression can't be evaluated.
' 05/03/2007 - SBAKKER - URD 10950
'            - Fixed problem with changing VBP files.
' 04/27/2007 - sbakker - Change unnecessary "Public"s to "Private"s.
' 01/18/2007 - sbakker - Leave "%" in EXE and VBP names, while replacing with
'              "_" for directory names.
' 01/10/2007 - sbakker - Added logic to handle both %SYSVOL and _SYSVOL,
'              %IDRISYS and _IDRISYS.
'            - Strip trailing "\" from FromPath/ToPath properties. Also make
'              uppercase.
' 10/16/2006 - Added checking for output files (.cvp and .bas) being read-only.
' 10/02/2006 - Added new function "IsNumericByteItem". Used for Enter/Edit nums,
'              to determine if the target variable is a byte register. Creates
'              statements using new ENTERBYTE and EDITBYTE commands. This will
'              prevent numeric overflows, as the value gets rejected internally
'              before it ever reaches the target variable.
' ------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Collections.Generic

<ComClass(Compiler.ClassId, Compiler.InterfaceId, Compiler.EventsId)>
Public Class Compiler

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

#Region " --- COM GUIDs --- "
    ' These  GUIDs provide the COM identity for this class
    ' and its COM interfaces. If you change them, existing
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4bc1fb88-9573-492f-8d46-8ffc8c2c8077"
    Public Const InterfaceId As String = "f8f9aea2-ae55-46c6-b3df-8af762cebc91"
    Public Const EventsId As String = "232c1dc2-ebe8-4462-8af6-3ef74c2b00dd"
#End Region

#Region " --- Structures --- "

    Private Structure IFEntry
        Public IfType As String
        Public IfLabel As String
        Public IfHadElse As Boolean
        Public IfAndLevel As Integer
    End Structure

    Private Structure ForEntry
        Public ForLabel As String
        Public ForVar As String
        Public ForFrom As String
        Public ForTo As String
        Public ForStep As String
    End Structure

    Private Structure ParseError
        Public LineNum As Integer
        Public SourceLine As String
        Public ErrorDesc As String
    End Structure

#End Region

#Region " --- Internal Variables --- "

    Private m_FromPath As String = ""
    Private m_ToPath As String = ""
    Private m_KeepComments As Boolean = True
    Private m_LatestDateTime As Date = Nothing

    Private Renames As New Dictionary(Of String, String)
    Private Equates As New Dictionary(Of String, String)
    Private Renumbers As New Dictionary(Of String, String)
    Private IfStack As New List(Of IFEntry)
    Private ForStack As New List(Of ForEntry)
    Private IFLabelNum As Integer = 0
    Private ForLabelNum As Integer = 0
    Private JumpPointList As New List(Of Integer)
    Private GotoLineList As New List(Of Integer)
    Private ObjProgNum As Integer
    Private ParseErrors As New List(Of ParseError)

#End Region

#Region " --- Constructors and Properties --- "

    ' A creatable COM class must have a Public Sub New()
    ' with no parameters, otherwise, the class will not be
    ' registered in the COM registry and cannot be created
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal FromPath As String, ByVal ToPath As String)
        Me.FromPath = FromPath
        Me.ToPath = ToPath
    End Sub

    Public Property FromPath() As String
        Get
            Return m_FromPath
        End Get
        Set(ByVal value As String)
            Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
            If value.EndsWith("\") Then
                value = value.Substring(0, value.Length - 1)
            End If
            value = value.ToUpper
            If Not Directory.Exists(value) Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Directory doesn't exist: " + value)
            End If
            m_FromPath = value
        End Set
    End Property

    Public Property ToPath() As String
        Get
            Return m_ToPath
        End Get
        Set(ByVal value As String)
            Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
            If value.EndsWith("\") Then
                value = value.Substring(0, value.Length - 1)
            End If
            value = value.ToUpper
            If Not Directory.Exists(value) Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Directory doesn't exist: " + value)
            End If
            m_ToPath = value
        End Set
    End Property

    Public Property KeepComments() As Boolean
        Get
            Return m_KeepComments
        End Get
        Set(ByVal value As Boolean)
            m_KeepComments = value
        End Set
    End Property

    Public Function ErrorCount() As Integer
        Return ParseErrors.Count
    End Function

    Public Function ErrorList() As String
        Dim ErrorInfo As ParseError
        Dim Results As New StringBuilder
        ' ------------------------------
        For Each ErrorInfo In ParseErrors
            Results.Append(ErrorInfo.LineNum.ToString)
            Results.Append(" - ")
            Results.Append(ErrorInfo.SourceLine.Replace(vbTab, " "))
            Results.Append(vbCrLf)
            Results.Append("      ")
            Results.Append(ErrorInfo.ErrorDesc.Replace(vbTab, " "))
            Results.Append(vbCrLf)
        Next
        Return Results.ToString
    End Function

    Public Function LatestDateTime() As Date
        Return m_LatestDateTime
    End Function

#End Region

#Region " --- Main Public Routines --- "

    Public Function CompileToIL(ByVal FileName As String,
                                ByVal ChangedOnly As Boolean,
                                ByRef DidCompile As Boolean) As Integer
        ' --- Compiles Cadol program to an output file ---
        Dim CurrLine As String
        Dim ObjBaseName As String
        Dim ObjFileNameCVP As String
        Dim ObjFileNameVB6 As String
        Dim sw As StreamWriter
        Dim ErrorInfo As ParseError
        ' --------------------------
        DidCompile = False
        ParseErrors.Clear()
        If FileName = "" Then
            ErrorInfo = New ParseError
            With ErrorInfo
                .LineNum = -1
                .SourceLine = ""
                .ErrorDesc = "FileName not specified"
            End With
            ParseErrors.Add(ErrorInfo)
            ErrorInfo = Nothing
            Return -1
        End If
        If m_FromPath = "" OrElse m_ToPath = "" Then
            ErrorInfo = New ParseError
            With ErrorInfo
                .LineNum = -1
                .SourceLine = ""
                .ErrorDesc = "Paths not specified"
            End With
            ParseErrors.Add(ErrorInfo)
            ErrorInfo = Nothing
            Return -1
        End If
        If Not File.Exists(m_FromPath + "\" + FileName) Then
            ErrorInfo = New ParseError
            With ErrorInfo
                .LineNum = -1
                .SourceLine = ""
                .ErrorDesc = "File not found: " + m_FromPath + "\" + FileName
            End With
            ParseErrors.Add(ErrorInfo)
            ErrorInfo = Nothing
            Return -1
        End If
        ' --- Initialize variables ---
        ObjProgNum = -1
        ObjBaseName = CollapseFilename(FileName.ToUpper)
        If ObjBaseName.EndsWith("_K") Then
            ObjBaseName = ObjBaseName.Substring(0, ObjBaseName.Length - 2)
        End If
        ' --- load the program into a list ---
        m_LatestDateTime = File.GetLastWriteTimeUtc(m_FromPath + "\" + FileName)
        Dim sr As New StreamReader(m_FromPath + "\" + FileName)
        Dim Lines As New List(Of String)
        Do While Not sr.EndOfStream
            CurrLine = sr.ReadLine()
            Lines.Add(CurrLine)
        Loop
        sr.Close()
        ' --- perform first-pass translation on program ---
        Dim Lines1 As New List(Of String)
        If Not PerformFirstPass(Lines, Lines1, ObjBaseName, ObjProgNum) Then
            Return -1 ' error found
        End If
        ' --- get output filenames ---
        ObjFileNameCVP = ObjProgNum.ToString.PadLeft(3, "0"c) + "_" + ObjBaseName + ".CVP"
        If Not m_ToPath.EndsWith("\_IDRISYS") Then
            ObjFileNameVB6 = "modProg" + ObjProgNum.ToString.PadLeft(3, "0"c) + ".bas"
        Else
            ObjFileNameVB6 = "modSysProg" + ObjProgNum.ToString.PadLeft(3, "0"c) + ".bas"
        End If
        ' --- check if program has been renamed ---
        For Each TempFile As String In Directory.GetFiles(m_ToPath, ObjProgNum.ToString.PadLeft(3, "0"c) + "*.CVP")
            TempFile = TempFile.Substring(TempFile.LastIndexOf("\") + 1) ' remove path
            If TempFile <> ObjFileNameCVP Then
                Try
                    If (File.GetAttributes(m_ToPath + "\" + TempFile) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                        File.SetAttributes(m_ToPath + "\" + TempFile, FileAttributes.Normal)
                    End If
                    File.Delete(m_ToPath + "\" + TempFile)
                Catch ex As Exception
                    ' --- ignore errors ---
                End Try
            End If
        Next
        ' --- check if program needs to be compiled ---
        If Not ChangedOnly Then GoTo StartPass2
        If Not File.Exists(m_ToPath + "\" + ObjFileNameCVP) Then GoTo StartPass2
        If Not File.Exists(m_ToPath + "\" + ObjFileNameVB6) Then GoTo StartPass2
        If m_LatestDateTime > File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameCVP) Then GoTo StartPass2
        If m_LatestDateTime > File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameVB6) Then GoTo StartPass2
        GoTo Done
        ' --- perform second-pass translation on program ---
StartPass2:
        Dim Lines2 As New List(Of String)
        If Not PerformSecondPass(Lines1, Lines2) Then
            Return -1 ' error found
        End If
        ' --- perform third-pass renumbering and comment attachment ---
        Dim Lines3 As New List(Of String)
        If Not PerformThirdPass(Lines2, Lines3) Then
            Return -1 ' error found
        End If
        ' --- perform fourth-pass unique function handling ---
        Dim Lines4 As New List(Of String)
        If Not PerformFourthPass(Lines3, Lines4) Then
            Return -1 ' error found
        End If
        ' --- perform fifth-pass adding line numbers ---
        Dim Lines5 As New List(Of String)
        PerformFifthPass(Lines4, Lines5)
        ' --- check if file is read-only ---
        If File.Exists(m_ToPath + "\" + ObjFileNameCVP) Then
            If (File.GetAttributes(m_ToPath + "\" + ObjFileNameCVP) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = -1
                    .SourceLine = ""
                    .ErrorDesc = ObjFileNameCVP + " is Read-Only"
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Return -1
            End If
        End If
        ' --- output original IL results ---
        sw = New StreamWriter(m_ToPath + "\" + ObjFileNameCVP)
        For Each CurrLine In Lines3
            sw.WriteLine(CurrLine)
        Next
        sw.Close()
        ' --- check if file is read-only ---
        If File.Exists(m_ToPath + "\" + ObjFileNameVB6) Then
            If (File.GetAttributes(m_ToPath + "\" + ObjFileNameVB6) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = -1
                    .SourceLine = ""
                    .ErrorDesc = ObjFileNameCVP + " is Read-Only"
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Return -1
            End If
        End If
        ' --- output VB6 code results ---
        sw = New StreamWriter(m_ToPath + "\" + ObjFileNameVB6)
        For Each CurrLine In Lines5
            sw.WriteLine(CurrLine)
        Next
        sw.Close()
        DidCompile = True
        ' --- done ---
Done:
        If m_LatestDateTime < File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameCVP) Then
            m_LatestDateTime = File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameCVP)
        End If
        If m_LatestDateTime < File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameVB6) Then
            m_LatestDateTime = File.GetLastWriteTimeUtc(m_ToPath + "\" + ObjFileNameVB6)
        End If
        Return ObjProgNum
    End Function

    ''Public Function CompileToIL(ByVal Lines As List(Of String)) As List(Of String)
    ''    ' --- Compile Cadol program to a list of strings ---
    ''    Dim ObjBaseName As String
    ''    Dim ObjProgNum As Integer
    ''    ' -----------------------
    ''    ' --- Initialize variables ---
    ''    ObjProgNum = -1
    ''    ObjBaseName = ""
    ''    ' --- perform first-pass translation on program ---
    ''    Dim Lines1 As New List(Of String)
    ''    If Not PerformFirstPass(Lines, Lines1, ObjBaseName, ObjProgNum) Then
    ''        Return Nothing
    ''    End If
    ''    ' --- perform second-pass translation on program ---
    ''    Dim Lines2 As New List(Of String)
    ''    If Not PerformSecondPass(Lines1, Lines2) Then
    ''        Return Lines2
    ''    End If
    ''    ' --- perform third-pass renumbering and comment attachment ---
    ''    Dim Lines3 As New List(Of String)
    ''    PerformThirdPass(Lines2, Lines3)
    ''    ' --- perform fourth-pass unique function handling ---
    ''    Dim Lines4 As New List(Of String)
    ''    PerformFourthPass(Lines3, Lines4)
    ''    ' --- perform fifth-pass adding line numbers ---
    ''    Dim Lines5 As New List(Of String)
    ''    PerformFifthPass(Lines4, Lines5)
    ''    ' --- return result ---
    ''    Return Lines3
    ''End Function

#End Region

#Region " --- First Pass Routines --- "

    Private Function PerformFirstPass(ByRef Lines As List(Of String),
                                      ByRef Lines1 As List(Of String),
                                      ByRef ObjBaseName As String,
                                      ByRef ObjProgNum As Integer) As Boolean
        Dim Pos As Integer
        Dim Tokens() As String
        Dim IncludeFilename As String
        Dim CurrLine As String
        Dim SourceLine2 As String
        Dim ErrorInfo As ParseError
        Dim SourceLinenum As Integer
        Dim SourceLinenum2 As Integer
        ' ---------------------------
        FillRenamesWithDefaults()
        Equates.Clear()
        For SourceLinenum = 0 To Lines.Count - 1
            CurrLine = Lines(SourceLinenum)
            CurrLine = CurrLine.Trim
            If CurrLine.ToUpper.StartsWith("* .IDRIS ") Then
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = SourceLinenum
                    .SourceLine = Lines(SourceLinenum)
                    .ErrorDesc = """.IDRIS"" directive is no longer supported."
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Return False
            End If
            CurrLine = TokenizeLine(CurrLine)
            If CurrLine = "" Then Continue For
            CurrLine = CombineTokens(CurrLine)
            ' --- Check for trailing comment, might be empty comment ---
            Pos = CurrLine.IndexOf(vbTab + "!")
            If Pos > 0 Then
                ' --- Store trailing comment first ---
                Lines1.Add("!" + CurrLine.Substring(Pos + 1)) ' change to "!! ..."
                ' --- Get beginning of line for further processing ---
                CurrLine = CurrLine.Substring(0, Pos)
            End If
            ' --- Check for .OBJ line ---
            If CurrLine.StartsWith(".OBJ" + vbTab) Then
                ObjBaseName = CollapseFilename(CurrLine.Substring(5).ToUpper)
                ' --- remove trailing "_K" from object name ---
                If ObjBaseName.EndsWith("_K") Then
                    ObjBaseName = ObjBaseName.Substring(0, ObjBaseName.Length - 2)
                End If
                Continue For
            End If
            ' --- Check for END line ---
            If CurrLine.StartsWith("END" + vbTab) Then
                ObjProgNum = Integer.Parse(CurrLine.Substring(4))
                Continue For
            End If
            ' --- Check for RENAME line ---
            If CurrLine.StartsWith("RENAME" + vbTab) Then
                Tokens = CurrLine.Split(CChar(vbTab))
                If Tokens(2) <> "AS" Then
                    ErrorInfo = New ParseError
                    With ErrorInfo
                        .LineNum = SourceLinenum
                        .SourceLine = Lines(SourceLinenum)
                        .ErrorDesc = "Invalid RENAME statement"
                    End With
                    ParseErrors.Add(ErrorInfo)
                    ErrorInfo = Nothing
                    Return False
                End If
                ' --- Handle renames ---
                Try
                    If Not Renames.ContainsKey(Tokens(1)) Then
                        Throw New SystemException($"Invalid rename variable: {Tokens(1)}")
                    End If
                    If Renames.ContainsValue(Tokens(3)) Then
                        Throw New SystemException($"Rename alias already exists: {Tokens(3)}")
                    End If
                    Renames(Tokens(1)) = Tokens(3) ' rename n1 as price
                Catch
                    ErrorInfo = New ParseError
                    With ErrorInfo
                        .LineNum = SourceLinenum
                        .SourceLine = Lines(SourceLinenum)
                        .ErrorDesc = "Error parsing RENAME statement: " + Tokens(1) + " AS " + Tokens(3)
                    End With
                    ParseErrors.Add(ErrorInfo)
                    ErrorInfo = Nothing
                    Return False
                End Try
                Continue For
            End If
            ' --- Check for EQUATE line ---
            If CurrLine.StartsWith("EQUATE" + vbTab) Then
                Tokens = CurrLine.Split(CChar(vbTab))
                If Tokens(2) <> "TO" Then
                    ErrorInfo = New ParseError
                    With ErrorInfo
                        .LineNum = SourceLinenum
                        .SourceLine = Lines(SourceLinenum)
                        .ErrorDesc = "Invalid EQUATE statement"
                    End With
                    ParseErrors.Add(ErrorInfo)
                    ErrorInfo = Nothing
                    Return False
                End If
                Try
                    Do While Tokens(3).Length > 1 AndAlso Tokens(3).StartsWith("0")
                        Tokens(3) = Tokens(3).Substring(1)
                    Loop
                    Equates.Add(Tokens(1), Tokens(3)) ' equate value to 1
                Catch ex As Exception
                    ErrorInfo = New ParseError
                    With ErrorInfo
                        .LineNum = SourceLinenum
                        .SourceLine = Lines(SourceLinenum)
                        .ErrorDesc = "Duplicate EQUATE statements for " + Tokens(1)
                    End With
                    ParseErrors.Add(ErrorInfo)
                    ErrorInfo = Nothing
                    Return False
                End Try
                Continue For
            End If
            ' --- Check for INCLUDE line ---
            If CurrLine.StartsWith("INCLUDE" + vbTab) Then
                IncludeFilename = CurrLine.Substring(8)
                If IncludeFilename.IndexOf(vbTab) >= 0 Then
                    IncludeFilename = IncludeFilename.Substring(0, IncludeFilename.IndexOf(vbTab))
                End If
                If IncludeFilename.IndexOf(" "c) >= 0 Then
                    IncludeFilename = IncludeFilename.Substring(0, IncludeFilename.IndexOf(" "c))
                End If
                Try
                    If Not File.Exists(m_FromPath + "\" + IncludeFilename) Then
                        ErrorInfo = New ParseError
                        With ErrorInfo
                            .LineNum = SourceLinenum
                            .SourceLine = Lines(SourceLinenum)
                            .ErrorDesc = "Include File Not Found"
                        End With
                        ParseErrors.Add(ErrorInfo)
                        ErrorInfo = Nothing
                        Return False
                    End If
                    If m_LatestDateTime < File.GetLastWriteTimeUtc(m_FromPath + "\" + IncludeFilename) Then
                        m_LatestDateTime = File.GetLastWriteTimeUtc(m_FromPath + "\" + IncludeFilename)
                    End If
                    Dim sr_inc As New StreamReader(m_FromPath + "\" + IncludeFilename)
                    SourceLinenum2 = -1
                    Do While Not sr_inc.EndOfStream
                        SourceLine2 = sr_inc.ReadLine()
                        CurrLine = TokenizeLine(SourceLine2)
                        SourceLinenum2 += 1
                        If CurrLine = "" Then Continue Do
                        CurrLine = CombineTokens(CurrLine)
                        ' --- Check for trailing comment ---
                        Pos = CurrLine.IndexOf(vbTab + "!" + vbTab)
                        If Pos > 0 Then
                            ' --- Get beginning of line only ---
                            CurrLine = CurrLine.Substring(0, Pos)
                        End If
                        ' --- Check for full-line comment ---
                        If CurrLine.StartsWith("!") Then
                            Continue Do
                        End If
                        ' --- Check for RENAME line ---
                        If CurrLine.StartsWith("RENAME" + vbTab) Then
                            Tokens = CurrLine.Split(CChar(vbTab))
                            If Tokens(2) <> "AS" Then
                                ErrorInfo = New ParseError
                                With ErrorInfo
                                    .LineNum = SourceLinenum2
                                    .SourceLine = IncludeFilename + ": " + SourceLine2
                                    .ErrorDesc = "Invalid RENAME statement"
                                End With
                                ParseErrors.Add(ErrorInfo)
                                ErrorInfo = Nothing
                                Return False
                            End If
                            Try
                                If Not Renames.ContainsKey(Tokens(1)) Then
                                    Throw New SystemException($"Invalid rename variable: {Tokens(1)}")
                                End If
                                If Renames.ContainsValue(Tokens(3)) Then
                                    Throw New SystemException($"Rename alias already exists: {Tokens(3)}")
                                End If
                                Renames(Tokens(1)) = Tokens(3) ' rename n1 as price
                            Catch
                                ErrorInfo = New ParseError
                                With ErrorInfo
                                    .LineNum = SourceLinenum
                                    .SourceLine = Lines(SourceLinenum)
                                    .ErrorDesc = "Error parsing RENAME statement: " + Tokens(1) + " AS " + Tokens(3)
                                End With
                                ParseErrors.Add(ErrorInfo)
                                ErrorInfo = Nothing
                                Return False
                            End Try
                            Continue Do
                        End If
                        ' --- Check for EQUATE line ---
                        If CurrLine.StartsWith("EQUATE" + vbTab) Then
                            Tokens = CurrLine.Split(CChar(vbTab))
                            If Tokens(2) <> "TO" Then
                                ErrorInfo = New ParseError
                                With ErrorInfo
                                    .LineNum = SourceLinenum2
                                    .SourceLine = IncludeFilename + ": " + SourceLine2
                                    .ErrorDesc = "Invalid EQUATE statement"
                                End With
                                ParseErrors.Add(ErrorInfo)
                                ErrorInfo = Nothing
                                Return False
                            End If
                            Try
                                Do While Tokens(3).Length > 1 AndAlso Tokens(3).StartsWith("0")
                                    Tokens(3) = Tokens(3).Substring(1)
                                Loop
                                Equates.Add(Tokens(1), Tokens(3)) ' equate value to 1
                            Catch ex As Exception
                                ErrorInfo = New ParseError
                                With ErrorInfo
                                    .LineNum = SourceLinenum2
                                    .SourceLine = IncludeFilename + ": " + SourceLine2
                                    .ErrorDesc = "Duplicate EQUATE statements for " + Tokens(1)
                                End With
                                ParseErrors.Add(ErrorInfo)
                                ErrorInfo = Nothing
                                Return False
                            End Try
                            Continue Do
                        End If
                        ' --- Remove END command ---
                        If CurrLine = "END" Then
                            Continue Do
                        End If
                        ' --- Replace Renames and Equates ---
                        Try
                            CurrLine = ReplaceRenamesEquates(CurrLine)
                        Catch ex As Exception
                            ErrorInfo = New ParseError
                            With ErrorInfo
                                .LineNum = SourceLinenum2
                                .SourceLine = IncludeFilename + ": " + SourceLine2
                                .ErrorDesc = ex.Message
                            End With
                            ParseErrors.Add(ErrorInfo)
                            ErrorInfo = Nothing
                            Return False
                        End Try
                        ' --- Store command line ---
                        Lines1.Add(CurrLine)
                    Loop
                    sr_inc.Close()
                Catch ex As Exception
                    Throw ex
                End Try
                Continue For
            End If
            ' --- Replace Renames and Equates ---
            Try
                CurrLine = ReplaceRenamesEquates(CurrLine)
            Catch ex As Exception
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = SourceLinenum
                    .SourceLine = Lines(SourceLinenum)
                    .ErrorDesc = ex.Message
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Return False
            End Try
            ' --- Store command line ---
            Lines1.Add(CurrLine)
        Next
        If ObjProgNum < 0 Then
            Return False
        End If
        Return True
    End Function

    Private Function TokenizeLine(ByVal CurrLine As String) As String
        ' --- This pass will tokenize all the line items. ---
        Dim Result As New StringBuilder
        Dim InComment As Boolean = False
        Dim KeepNextSpace As Boolean = True
        Dim InQuote As Boolean = False
        Dim InToken As Boolean = False
        Dim QuoteChar As Char
        Dim LastWasAlpha As Boolean = False
        Dim CharPos As Integer = 0
        Dim CurrChar As Char
        ' ---------------------------------
        ' --- remove trailing box chars and spaces ---
        If CurrLine.StartsWith("* ") AndAlso CurrLine.EndsWith(" *") Then
            CurrLine = CurrLine.Substring(0, CurrLine.Length - 2)
        End If
        CurrLine = CurrLine.Trim ' remove any leading and trailing spaces
        CurrLine = CurrLine.Replace(vbTab, " "c) ' change tabs to spaces
        ' --- Check if line is special ---
        If CurrLine.StartsWith(".OBJ", StringComparison.CurrentCultureIgnoreCase) Then
            CharPos = CurrLine.IndexOf(" "c)
            If CharPos < 0 Then
                Result.Append(CurrLine.ToUpper)
            Else
                Result.Append(CurrLine.Substring(0, CharPos).ToUpper)
                Result.Append(vbTab)
                Result.Append(CurrLine.Substring(CharPos + 1).ToUpper)
            End If
            GoTo Done
        End If
        ' --- Ignore all other Cadol compiler directives ---
        If CurrLine.StartsWith("."c) Then
            GoTo Done
        End If
        ' --- Ignore datestamp line ---
        If CurrLine.StartsWith("IF 1#1 DISPLAY", StringComparison.CurrentCultureIgnoreCase) Then
            GoTo Done
        End If
        ' --- Check for Include lines ---
        If CurrLine.StartsWith("INCLUDE ", StringComparison.CurrentCultureIgnoreCase) Then
            CharPos = CurrLine.IndexOf(" "c)
            If CharPos >= 0 Then
                Result.Append(CurrLine.Substring(0, CharPos).ToUpper)
                Result.Append(vbTab)
                Result.Append(CurrLine.Substring(CharPos + 1).ToUpper)
            End If
            GoTo Done
        End If
        ' --- check out each character in the line ---
        For CharNum As Integer = 0 To CurrLine.Length - 1
            CurrChar = CurrLine(CharNum)
            If CharNum = 0 AndAlso CurrChar = "*"c Then
                ' --- alternate form of comment ---
                LastWasAlpha = False
                InComment = True
                Result.Append("!")
                Result.Append(vbTab)
                KeepNextSpace = False
            ElseIf InQuote Then
                ' --- check for ending quote ---
                LastWasAlpha = False
                If CurrChar = QuoteChar Then
                    Result.Append(QuoteChar)
                    InToken = True
                    InQuote = False
                ElseIf CurrChar = vbTab Then
                    Result.Append(" "c)
                ElseIf CurrChar = """"c Then
                    Result.Append(""""c)
                Else
                    Result.Append(CurrChar)
                End If
            ElseIf InComment Then
                ' --- for comments, just append the chars ---
                LastWasAlpha = False
                ' --- ignore first spaces ---
                If CurrChar <> " "c OrElse KeepNextSpace Then
                    Result.Append(CurrChar)
                    KeepNextSpace = True
                End If
            ElseIf (CurrChar = "!"c) OrElse
                   (CurrChar = "'"c AndAlso Not InQuote AndAlso CurrLine.LastIndexOf("'"c) = CharNum) Then
                ' --- normal form of comment ---
                LastWasAlpha = False
                InComment = True
                If InToken Then
                    Result.Append(vbTab)
                    InToken = False
                End If
                Result.Append("!")
                Result.Append(vbTab)
                KeepNextSpace = False
            ElseIf CurrChar = """"c OrElse
                   CurrChar = "'"c OrElse
                   CurrChar = "$"c OrElse
                   CurrChar = "%"c Then
                ' --- these are all Cadol quote characters ---
                LastWasAlpha = False
                If InToken Then
                    Result.Append(vbTab)
                    InToken = False
                End If
                InQuote = True
                QuoteChar = CurrChar ' save actual quote char
                Result.Append(QuoteChar) ' ### Result.Append("""")
                InToken = True
            ElseIf Char.IsWhiteSpace(CurrChar) Then
                ' --- spaces or tabs separate tokens ---
                LastWasAlpha = False
                If InToken Then
                    Result.Append(vbTab)
                    InToken = False
                End If
            ElseIf Char.IsLetterOrDigit(CurrChar) OrElse CurrChar = "_"c OrElse CurrChar = "."c Then
                If InToken AndAlso Not LastWasAlpha Then
                    Result.Append(vbTab)
                End If
                InToken = True
                Result.Append(Char.ToUpper(CurrChar))
                LastWasAlpha = True
            Else ' must be a symbol of some kind 
                LastWasAlpha = False
                If InToken Then
                    Result.Append(vbTab)
                    InToken = False
                End If
                Result.Append(CurrChar)
                InToken = True
            End If
        Next
Done:
        Return Result.ToString
    End Function

    Private Function CombineTokens(ByVal CurrLine As String) As String
        Dim Result As String = CurrLine
        ' --- check for operators ---
        Result = DoCombine(Result, "<", ">")
        Result = DoCombine(Result, "<", "=")
        Result = DoCombine(Result, ">", "=")
        ' --- combine multi-word tokens into single word ---
        Result = DoCombine(Result, "GO", "TO")
        Result = DoCombine(Result, "END", "IF")
        Result = DoReplace(Result, "TEXT" + vbTab + "AREA", "TFA")
        ' --- combine multi-word commands into single command ---
        Result = DoCombine(Result, "ASSIGN", "DEVICE")
        Result = DoCombine(Result, "ASSIGN", "PRINTER")
        Result = DoCombine(Result, "BACKSPACE", "CHANNEL")
        Result = DoCombine(Result, "BREAK", "POINT") ' IDRIS debug command
        Result = DoCombine(Result, "CHAR", "DELETE")
        Result = DoCombine(Result, "CHAR", "INSERT")
        Result = DoCombine(Result, "CLOSE", "CHANNEL")
        Result = DoCombine(Result, "CLOSE", "DEVICE")
        Result = DoCombine(Result, "CLOSE", "TFA")
        Result = DoCombine(Result, "CLOSE", "VOLUME")
        Result = DoCombine(Result, "CONTROL", "CHANNEL")
        Result = DoCombine(Result, "CREATE", "CHANNEL")
        Result = DoCombine(Result, "CREATE", "DEVICE")
        Result = DoCombine(Result, "CREATE", "TFA")
        Result = DoCombine(Result, "CREATE", "VOLUME")
        Result = DoCombine(Result, "CURSOR", "AT")
        Result = DoCombine(Result, "DELETE", "CHANNEL")
        Result = DoCombine(Result, "DEVICE", "OFF")
        Result = DoCombine(Result, "DEVICE", "ON")
        Result = DoCombine(Result, "EOF", "CHANNEL")
        Result = DoCombine(Result, "GRAPH", "OFF")
        Result = DoCombine(Result, "GRAPH", "ON")
        Result = DoCombine(Result, "INIT", "FETCH")
        Result = DoCombine(Result, "INIT", "SORT")
        Result = DoCombine(Result, "LINE", "DELETE")
        Result = DoCombine(Result, "LINE", "INSERT")
        Result = DoCombine(Result, "OPEN", "CHANNEL")
        Result = DoCombine(Result, "OPEN", "DATA")
        Result = DoCombine(Result, "OPEN", "DEVICE")
        Result = DoCombine(Result, "OPEN", "DIRECTORY")
        Result = DoCombine(Result, "OPEN", "TFA")
        Result = DoCombine(Result, "OPEN", "VOLUME")
        Result = DoCombine(Result, "OPEN", "DIR")
        Result = DoCombine(Result, "OPENDIR", "DEVICEON")
        Result = DoCombine(Result, "OPENDIR", "LIB")
        Result = DoReplace(Result, "OPENDIR" + vbTab + "TEXT", "OPENDIRTFA")
        Result = DoCombine(Result, "OPENDIR", "TFA")
        Result = DoCombine(Result, "OPENDIR", "VOLUME")
        Result = DoCombine(Result, "PRINT", "FF")
        Result = DoCombine(Result, "PRINT", "NL")
        Result = DoCombine(Result, "PRINT", "OFF")
        Result = DoCombine(Result, "PRINT", "ON")
        Result = DoCombine(Result, "PRINTER", "OFF")
        Result = DoCombine(Result, "PRINTER", "ON")
        Result = DoCombine(Result, "READ", "CHANNEL")
        Result = DoCombine(Result, "READ", "KEY")
        Result = DoCombine(Result, "READ", "REC")
        Result = DoCombine(Result, "RELEASE", "DEVICE")
        Result = DoCombine(Result, "RELEASE", "PRINTER")
        Result = DoCombine(Result, "RELEASE", "TERMINAL")
        Result = DoCombine(Result, "RENAME", "CHANNEL")
        Result = DoCombine(Result, "REWIND", "CHANNEL")
        Result = DoCombine(Result, "SCROLL", "DOWN")
        Result = DoCombine(Result, "SCROLL", "UP")
        Result = DoCombine(Result, "TAB", "CANCEL")
        Result = DoCombine(Result, "TAB", "CLEAR")
        Result = DoCombine(Result, "TAB", "SET")
        Result = DoCombine(Result, "TRACE", "OFF") ' IDRIS debug command
        Result = DoCombine(Result, "TRACE", "ON") ' IDRIS debug command
        Result = DoCombine(Result, "WHEN", "ESCAPE")
        Result = DoCombine(Result, "WHEN", "CANCEL")
        Result = DoCombine(Result, "WHEN", "ERROR")
        Result = DoCombine(Result, "WIND", "CHANNEL")
        Result = DoCombine(Result, "WRITE", "BACK")
        Result = DoCombine(Result, "WRITE", "CHANNEL")
        Result = DoCombine(Result, "WRITE", "KEY")
        Result = DoCombine(Result, "WRITE", "REC")
        ' --- remove unneccessary "ON" ---
        Result = DoReplace(Result, "OPENDATA" + vbTab + "ON", "OPENDATA")
        Result = DoReplace(Result, "OPENDIRECTORY" + vbTab + "ON", "OPENDIRECTORY")
        Result = DoReplace(Result, "OPENDIRDEVICEON", "OPENDIRDEVICE")
        Result = DoReplace(Result, "OPENDIRDEVICE" + vbTab + "ON", "OPENDIRDEVICE")
        Result = DoReplace(Result, "OPENDIRLIB" + vbTab + "ON", "OPENDIRLIB")
        Result = DoReplace(Result, "OPENDIRTFA" + vbTab + "ON", "OPENDIRTFA")
        Result = DoReplace(Result, "OPENDIRVOLUME" + vbTab + "ON", "OPENDIRVOLUME")
        ' --- normalize different formats of same command ---
        Result = DoReplace(Result, "ASSIGNPRINTER", "ASSIGNDEVICE" + vbTab + "PRTNUM")
        Result = DoReplace(Result, "BREAKPOINT", "BREAK")
        Result = DoReplace(Result, "DEVICEOFF", "PRINTOFF")
        Result = DoReplace(Result, "DEVICEON", "PRINTON")
        Result = DoReplace(Result, "ESCAPE", "ESC")
        Result = DoReplace(Result, "OVERPRINT", "CR")
        Result = DoReplace(Result, "PRINTFF", "FF")
        Result = DoReplace(Result, "PRINTNL", "NL")
        Result = DoReplace(Result, "PRINTEROFF", "PRINTOFF")
        Result = DoReplace(Result, "PRINTERON", "PRINTON")
        Result = DoReplace(Result, "RELEASEPRINTER", "RELEASEDEVICE")
        Result = DoReplace(Result, "THEN" + vbTab + "GOTO", "GOTO")
        Result = DoReplace(Result, "TRACEOFF", "TROFF")
        Result = DoReplace(Result, "TRACEON", "TRON")
        ' --- remove optional items ---
        Result = DoRemove(Result, "ERR")
        Result = DoRemove(Result, "LET")
        ' --- fix ambiguous items ---
        If Result = "RESET" OrElse
           Result.EndsWith(vbTab + "RESET") OrElse
           Result.IndexOf("RESET" + vbTab + "!") >= 0 Then
            Result = DoReplace(Result, "RESET", "RESETSCREEN")
        End If
        ' --- Done ---
        Return Result
    End Function

    Private Function DoCombine(ByVal CurrLine As String, ByVal Part1 As String, ByVal Part2 As String) As String
        ' --- Add surrounding Tabs ---
        Dim Result As String = vbTab + CurrLine + vbTab
        ' --- Combine parts ---
        Do While Result.IndexOf(vbTab + Part1 + vbTab + Part2 + vbTab) >= 0
            Result = Result.Replace(vbTab + Part1 + vbTab + Part2 + vbTab, vbTab + Part1 + Part2 + vbTab)
        Loop
        ' --- Remove surrounding Tabs ---
        Return Result.Substring(1, Result.Length - 2)
    End Function

    Private Function DoReplace(ByVal CurrLine As String, ByVal FromValue As String, ByVal ToValue As String) As String
        ' --- Add surrounding Tabs ---
        Dim Result As String = vbTab + CurrLine + vbTab
        ' --- Replace specified string ---
        Do While Result.IndexOf(vbTab + FromValue + vbTab) >= 0
            Result = Result.Replace(vbTab + FromValue + vbTab, vbTab + ToValue + vbTab)
        Loop
        ' --- Remove surrounding Tabs ---
        Return Result.Substring(1, Result.Length - 2)
    End Function

    Private Function DoRemove(ByVal CurrLine As String, ByVal FromValue As String) As String
        ' --- Add surrounding Tabs ---
        Dim Result As String = vbTab + CurrLine + vbTab
        ' --- Replace specified string ---
        Do While Result.IndexOf(vbTab + FromValue + vbTab) >= 0
            Result = Result.Replace(vbTab + FromValue + vbTab, vbTab)
        Loop
        ' --- Remove surrounding Tabs ---
        Return Result.Substring(1, Result.Length - 2)
    End Function

    Private Function IsReservedWord(ByVal Value As String) As Boolean
        ' --- check for registers and system variables ---
        If IsRegister(Value) Then Return True
        If IsBufferPtrByValue(Value) Then Return True
        If IsSystemVarByValue(Value) Then Return True
        ' --- all numbers are reserved! ---
        If NumOnly(Value) Then Return True
        ' --- This is every reserved word in IDRIS. ---
        Select Case Value
            Case "AND"
            Case "ASSIGNDEVICE"
            Case "ATT"
            Case "BACK"
            Case "BACKSPACE"
            Case "BACKSPACECHANNEL"
            Case "BELL"
            Case "BREAK"
            Case "BY"
            Case "CAN"
            Case "CANCEL"
            Case "CANVAL"
            Case "CHARDELETE"
            Case "CHARINSERT"
            Case "CLEAR"
            Case "CLEARSCREEN"
            Case "CLOSE"
            Case "CLOSECHANNEL"
            Case "CLOSEDEVICE"
            Case "CLOSEFILE"
            Case "CLOSETFA"
            Case "CLOSEVOLUME"
            Case "CMD"
            Case "COMMENT"
            Case "CONTROLCHANNEL"
            Case "CONVERT"
            Case "CR"
            Case "CRD"
            Case "CREATECHANNEL"
            Case "CREATETFA"
            Case "CURSORAT"
            Case "DATE"
            Case "DATEVAL"
            Case "DCH"
            Case "DEBUG"
            Case "DELAY"
            Case "DELETE"
            Case "DELETECHANNEL"
            Case "DISPLAY"
            Case "DISPLAYNUM"
            Case "DISPLAYSPACE"
            Case "DISPLAYSTRING"
            Case "DIVREM"
            Case "DO"
            Case "DOWN"
            Case "EDIT"
            Case "EDITALPHA"
            Case "EDITBYTE"
            Case "EDITNUM"
            Case "ELSE"
            Case "END"
            Case "ENDIF"
            Case "ENTER"
            Case "ENTERALPHA"
            Case "ENTERBYTE"
            Case "ENTERNUM"
            Case "EOFCHANNEL"
            Case "EQUATE"
            Case "ERR"
            Case "ESC"
            Case "ESCAPE"
            Case "ESCVAL"
            Case "EXECSQL"
            Case "EXITRUNTIME"
            Case "FALSE"
            Case "FALSEVAL"
            Case "FATALERROR"
            Case "FETCH"
            Case "FF"
            Case "FLIP"
            Case "FLOP"
            Case "FOR"
            Case "FORMATNUM"
            Case "FROM"
            Case "GETCLIENTLIST"
            Case "GETDATESTR"
            Case "GETDATEVAL"
            Case "GETENVIRONMENT"
            Case "GETLOGINID"
            Case "GETREADONLY"
            Case "GETTIMEHUND"
            Case "GETTIMESTR"
            Case "GETTIMEVAL"
            Case "GOS"
            Case "GOSUB"
            Case "GOSUBPROG"
            Case "GOTO"
            Case "GRAPHOFF"
            Case "GRAPHON"
            Case "HEX"
            Case "HOME"
            Case "IF"
            Case "INCLUDE"
            Case "INIT"
            Case "INITFETCH"
            Case "INITSORT"
            Case "KEY"
            Case "KFREE"
            Case "KLOCK"
            Case "LEFT"
            Case "LET"
            Case "LINEDELETE"
            Case "LINEINSERT"
            Case "LOAD"
            Case "LOADPROG"
            Case "LOCK"
            Case "LOCKREC"
            Case "LOCKVAL"
            Case "MERGE"
            Case "MOVE"
            Case "MOVEDOWN"
            Case "MOVELEFT"
            Case "MOVERIGHT"
            Case "MOVEUP"
            Case "MUSTEXIT"
            Case "NEXT"
            Case "NL"
            Case "NOP"
            Case "NOSIGN"
            Case "NULL"
            Case "ON"
            Case "ONESTEP"
            Case "OPENCHANNEL"
            Case "OPENDATA"
            Case "OPENDEVICE"
            Case "OPENDIRDEVICE"
            Case "OPENDIRECTORY"
            Case "OPENDIRLIB"
            Case "OPENDIRTFA"
            Case "OPENDIRVOLUME"
            Case "OPENSORTFILE"
            Case "OPENTFA"
            Case "OPENVOLUME"
            Case "OR"
            Case "PACK"
            Case "PAD"
            Case "PRINT"
            Case "PRINTOFF"
            Case "PRINTON"
            Case "READ"
            Case "READCHANNEL"
            Case "READFILE"
            Case "READKEY"
            Case "READREC"
            Case "REC"
            Case "REJECT"
            Case "RELEASEDEVICE"
            Case "RELEASETERMINAL"
            Case "REM"
            Case "REMVAL"
            Case "RENAME"
            Case "RENAMECHANNEL"
            Case "REPEAT"
            Case "RESET"
            Case "RESETSCREEN"
            Case "RESTOREFILEINFO"
            Case "RETURN"
            Case "RETURNPROG"
            Case "REWINDCHANNEL"
            Case "RIGHT"
            Case "SAVEFILEINFO"
            Case "SCROLLDOWN"
            Case "SCROLLUP"
            Case "SET"
            Case "SETSUBQUERY"
            Case "SETSUBQUERYFILE"
            Case "SIGNSET"
            Case "SKIP"
            Case "SORT"
            Case "SORTALPHA"
            Case "SORTNUM"
            Case "SPACE"
            Case "SPOOL"
            Case "STAY"
            Case "STEP"
            Case "SYSVAR"
            Case "TAB"
            Case "TABCANCEL"
            Case "TABCLEAR"
            Case "TABCURSOR"
            Case "TABSET"
            Case "THEN"
            Case "THROWERROR"
            Case "TO"
            Case "TRACEOFF"
            Case "TRACEON"
            Case "TRAP"
            Case "TROFF"
            Case "TRON"
            Case "TRUE"
            Case "TRUEVAL"
            Case "UNLOCK"
            Case "UNLOCKREC"
            Case "UNTIL"
            Case "UP"
            Case "UPDATE"
            Case "WHENCANCEL"
            Case "WHENERROR"
            Case "WHENESCAPE"
            Case "WHILE"
            Case "WINDCHANNEL"
            Case "WINEXEC"
            Case "WINEXECALL"
            Case "WINGETATTR"
            Case "WININIT"
            Case "WINSAVE"
            Case "WINSAVEALL"
            Case "WINSETATTR"
            Case "WINSHOW"
            Case "WINSTATUS"
            Case "WINUNLOAD"
            Case "WINUNLOADALL"
            Case "WINVALIDATE"
            Case "WINVALIDATEALL"
            Case "WRITE"
            Case "WRITEBACK"
            Case "WRITECHANNEL"
            Case "WRITEFILE"
            Case "WRITEKEY"
            Case "WRITEREC"
            Case "ZERO"
            Case Else
                Return False
        End Select
        Return True
    End Function

    Private Function IsRegister(ByVal Value As String) As Boolean
        Dim TempVal As Integer
        ' --------------------
        ' --- Check if VarName is a standard variable. ---
        If "ABCDENFG".IndexOf(Value(0)) < 0 Then Return False
        ' --- single-letter variables ok ---
        If Value.Length = 1 Then Return True
        ' --- check if variable has a valid number ---
        If Not NumOnly(Value.Substring(1)) Then Return False
        ' --- leading zeros not allowed ---
        If Value(1) = "0"c Then Return False
        ' --- save value for more checking ---
        TempVal = Integer.Parse(Value.Substring(1))
        ' --- zero is not valid here ---
        If TempVal < 1 Then Return False
        If TempVal > 99 Then Return False
        ' --- alphas and globals only go up to 9, not 99 ---
        If TempVal > 9 Then
            If Value(0) = "N"c Then Return True
            If Value(0) = "F"c Then Return True
            Return False
        End If
Done:
        IsRegister = True
        Exit Function
ErrorFound:
        IsRegister = False
    End Function

    Private Function ReplaceRenamesEquates(ByVal Value As String) As String
        Dim Changed As Boolean = False
        Dim TokenNum As Integer
        Dim Tokens() As String = Value.Split(CChar(vbTab))
        ' -----------------------------------------
        For TokenNum = 0 To Tokens.GetUpperBound(0)
            If Tokens(TokenNum) = "!!" OrElse Tokens(TokenNum) = "!" Then
                Exit For
            End If
            If Renames.ContainsValue(Tokens(TokenNum)) Then
                For Each DictItem As KeyValuePair(Of String, String) In Renames
                    If DictItem.Value = Tokens(TokenNum) Then
                        If DictItem.Value <> DictItem.Key Then
                            Tokens(TokenNum) = DictItem.Key
                            Changed = True
                        End If
                        Exit For
                    End If
                Next
            ElseIf Renames.ContainsKey(Tokens(TokenNum)) Then
                Throw New SystemException($"Variable has been renamed and isn't available: {Tokens(TokenNum)}")
            End If
            If Equates.ContainsKey(Tokens(TokenNum)) Then
                Tokens(TokenNum) = Equates.Item(Tokens(TokenNum))
                Changed = True
            End If
        Next
        If Changed Then
            Dim Result As New StringBuilder
            For TokenNum = 0 To Tokens.GetUpperBound(0)
                If TokenNum > 0 Then
                    Result.Append(vbTab)
                End If
                Result.Append(Tokens(TokenNum))
            Next
            Return Result.ToString
        Else
            Return Value
        End If
    End Function

    Private Function CollapseFilename(ByVal Value As String) As String
        Dim Result As New StringBuilder
        Dim CharNum As Integer
        Dim FirstChar As Boolean = True
        Dim NeedUL As Boolean = False
        ' ---------------------------
        For CharNum = 0 To Value.Length - 1
            Select Case Value(CharNum)
                Case "A"c To "Z"c, "a"c To "z"c, "0"c To "9"c
                    If NeedUL Then
                        Result.Append("_"c)
                        NeedUL = False
                    End If
                    Result.Append(Value(CharNum))
                    FirstChar = False
                Case Else
                    If Not FirstChar Then
                        NeedUL = True
                    End If
            End Select
        Next
        Return Result.ToString
    End Function

#End Region

#Region " --- Second Pass Routines --- "

    Private Function PerformSecondPass(ByRef Lines1 As List(Of String), ByRef Lines2 As List(Of String)) As Boolean
        Dim CurrLine As String
        Dim CurrLine2 As String
        Dim LineNum As Integer
        Dim TempLines() As String
        Dim LastLineNop As Boolean = False
        Dim ErrorInfo As ParseError
        Dim SourceLinenum As Integer
        ' --------------------------------
        For SourceLinenum = 0 To Lines1.Count - 1
            CurrLine = Lines1(SourceLinenum)
            CurrLine = SecondPass(CurrLine)
            If CurrLine.StartsWith("***") Then
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = SourceLinenum
                    .SourceLine = Lines1(SourceLinenum)
                    .ErrorDesc = CurrLine
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Lines2.Add(CurrLine + " - " + Lines1(SourceLinenum).Replace(vbTab, " "))
                Continue For
            End If
            TempLines = CurrLine.Split(CChar(vbLf))
            For Each CurrLine2 In TempLines
                If CurrLine2 = "" Then Continue For
                If CurrLine2.StartsWith("!") Then
                    If LastLineNop Then
                        Lines2.Add(Lines2(Lines2.Count - 1))
                        Lines2(Lines2.Count - 2) = CurrLine2
                    Else
                        Lines2.Add(CurrLine2)
                    End If
                Else
                    If LastLineNop AndAlso Not CurrLine2.StartsWith(":") Then
                        ' --- join this line to the last line, stripping out the "NOP" ---
                        LineNum = Lines2.Count - 1
                        Lines2(LineNum) = Lines2(LineNum).Substring(0, Lines2(LineNum).Length - 4) + CurrLine2
                    ElseIf LastLineNop Then
                        ' --- need to keep the NOP ---
                        LineNum = Lines2.Count - 1
                        Lines2(LineNum) = Lines2(LineNum).Substring(0, Lines2(LineNum).Length - 4) + "NOP"
                        Lines2.Add(CurrLine2)
                    Else
                        ' --- just add the current line ---
                        Lines2.Add(CurrLine2)
                    End If
                    LastLineNop = CurrLine2.EndsWith(vbTab + "\NOP")
                End If
            Next
        Next
        If LastLineNop Then
            ' --- need to keep the NOP ---
            LineNum = Lines2.Count - 1
            Lines2(LineNum) = Lines2(LineNum).Substring(0, Lines2(LineNum).Length - 4) + "NOP"
        End If
        Return (ParseErrors.Count = 0)
    End Function

    Private Function SecondPass(ByVal Value As String) As String
        Dim Changed As Boolean = False
        Dim Result As StringBuilder
        Dim TokenNum As Integer = 0
        Dim SaveTokenNum As Integer = 0
        Dim TempSaveToken As Integer = 0
        Dim Tokens() As String
        Dim LastToken As Integer
        Dim CurrIFEntry As IFEntry
        Dim CurrForEntry As ForEntry
        Dim CompoundCommand As Boolean = False
        Dim AndFlag As Boolean = False
        Dim OrFlag As Boolean = False
        Dim ErrorMsg As String = "*** Invalid Command Syntax ***"
        ' -------------------------------------------------------
        ' --- Check for comment ---
        If Value.StartsWith("!") Then
            If m_KeepComments Then
                Return Value
            Else
                Return ""
            End If
        End If
        ' --- Initialize values ---
        Tokens = Value.Split(CChar(vbTab))
        LastToken = Tokens.GetUpperBound(0)
        ' --- Check for leading line number ---
        If TokenNum = 0 AndAlso NumOnly(Tokens(TokenNum)) Then
            Tokens(TokenNum) = ":" + Tokens(TokenNum)
            Changed = True
            TokenNum += 1
            ' --- Check if line number was only token ---
            If TokenNum > LastToken Then
                Tokens(TokenNum - 1) += vbTab + "NOP"
                Changed = True
                GoTo Done
            End If
        End If
NextCommand:
        ' --- Check for alpha and numeric assignments ---
        SaveTokenNum = TokenNum
        If IsAlphaTarget(Tokens, TokenNum) Then
            If TokenNum <= LastToken AndAlso Tokens(TokenNum) = "=" Then
                TokenNum += 1
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Assignment ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Alpha Assignment ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            End If
        End If
        TokenNum = SaveTokenNum
        If IsNumTarget(Tokens, TokenNum) Then
            If TokenNum <= LastToken AndAlso Tokens(TokenNum) = "=" Then
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Assignment ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Numeric Assignment ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            End If
        End If
        TokenNum = SaveTokenNum
        ' --- Check for simple one-word commands ---
        If TokenNum = LastToken Then
            Select Case Tokens(TokenNum)
                Case "BACK" : GoTo Done
                Case "CANCEL" : GoTo Done
                Case "CLEAR" : GoTo Done
                Case "CLOSETFA" : GoTo Done
                Case "CLOSEVOLUME" : GoTo Done
                Case "CR" : GoTo Done
                Case "ESC" : GoTo Done
                Case "GRAPHOFF" : GoTo Done
                Case "GRAPHON" : GoTo Done
                Case "HOME" : GoTo Done
                Case "INITFETCH" : GoTo Done
                Case "KLOCK" : GoTo Done
                Case "KFREE" : GoTo Done
                Case "LOCK" : GoTo Done
                Case "MERGE" : GoTo Done
                Case "NOP" : GoTo Done
                Case "PRINTOFF" : GoTo Done
                Case "PRINTON" : GoTo Done
                Case "REJECT" : GoTo Done
                Case "RESETSCREEN" : GoTo Done
                Case "RELEASEDEVICE" : GoTo Done
                Case "RETURN" : GoTo Done
                Case "STAY" : GoTo Done
                Case "UNLOCK" : GoTo Done
                Case "TABCANCEL" : GoTo Done
                Case "TABCLEAR" : GoTo Done
                Case "TABSET" : GoTo Done
                Case "WRITEBACK" : GoTo Done
                Case "ZERO" : GoTo Done
            End Select
        End If
        ' --- check for other commands ---
        Select Case Tokens(TokenNum)
            Case "EQUATE", "RENAME", "INCLUDE", "END"
                ' --- metacommands ---
                ErrorMsg = "*** Metacommand not processed ***"
                GoTo ErrorFound
            Case "GOTO", "GOS", "CREATETFA", "OPENTFA", "OPENVOLUME", "RELEASETERMINAL"
                ' --- token/label pairs (expressions not allowed) ---
                TokenNum += 1
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                Changed = True
                GoTo Done
            Case "ATT", "CLOSE", "CLOSECHANNEL", "CLOSEDEVICE",
                 "DELAY", "GOSUB", "INITSORT", "LOAD", "PAD"
                ' --- token/expr pairs, expr is required ---
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "BACKSPACE", "BELL", "CHARDELETE", "CHARINSERT", "CRD",
                 "DOWN", "FF", "LEFT", "LINEDELETE", "LINEINSERT", "NL",
                 "RIGHT", "SCROLLDOWN", "SCROLLUP", "SPACE", "TAB", "UP"
                ' --- token/expr pairs, expr is optional ---
                If TokenNum = LastToken Then GoTo Done
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "ASSIGNDEVICE", "CONTROLCHANNEL", "CREATECHANNEL", "DELETE",
                 "DELETECHANNEL", "EOFCHANNEL", "FETCH", "OPENCHANNEL", "OPENDATA",
                 "OPENDEVICE", "OPENDIRDEVICE", "OPENDIRECTORY", "OPENDIRLIB",
                 "OPENDIRTFA", "OPENDIRVOLUME", "OPENSORTFILE", "READ",
                 "READKEY", "READREC", "RENAMECHANNEL", "REWINDCHANNEL",
                 "WINDCHANNEL", "WRITE", "WRITEKEY", "WRITEREC"
                ' --- token/expr/label ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                ' --- parse line number ---
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                Changed = True
                GoTo Done
            Case "BACKSPACECHANNEL"
                ' --- token/expr/until/label ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "UNTIL" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Tokens(TokenNum).Length <> 1 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) < 0 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                End If
                ' --- parse line number ---
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                Changed = True
                GoTo Done
            Case "READCHANNEL"
                ' --- token/expr/to/until/label ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "TO" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Len(Tokens(TokenNum)) <> 1 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) < 0 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                End If
                If Tokens(TokenNum) = "UNTIL" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Len(Tokens(TokenNum)) <> 1 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) < 0 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                End If
                ' --- parse line number ---
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                Changed = True
                GoTo Done
            Case "WRITECHANNEL"
                ' --- token/expr/from/label ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "FROM" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Len(Tokens(TokenNum)) <> 1 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) < 0 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                End If
                ' --- parse line number ---
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                Changed = True
                GoTo Done
            Case "ON"
                ' --- ON expr GOTO label, label, label... ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "GOTO" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                Do
                    ' --- parse line number ---
                    If Not NumOnly(Tokens(TokenNum)) Then
                        ErrorMsg = "*** Invalid Line Number ***"
                        GoTo ErrorFound
                    End If
                    Tokens(TokenNum) = ":" + Tokens(TokenNum) ' change to label format
                    Changed = True
                    ' --- check for more line numbers ---
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo Done
                    If Tokens(TokenNum) <> "," Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                Loop Until TokenNum > LastToken
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "DISPLAY"
                ' --- display ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If IsAlphaTerm(Tokens, TokenNum) Then
                    If TokenNum <= LastToken Then
                        ErrorMsg = "*** End-of-line Expected ***"
                        GoTo ErrorFound
                    End If
                    GoTo Done
                End If
                If Tokens(TokenNum) = "(" Then ' numeric format
                    TokenNum += 1
                    SaveTokenNum = TokenNum
                    If Not IsDisplayFmt(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Display Format ***"
                        GoTo ErrorFound
                    End If
                    If Tokens(TokenNum - 1) <> ")" Then
                        ErrorMsg = "*** Invalid Display Format ***"
                        GoTo ErrorFound
                    End If
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                    If TokenNum <= LastToken Then
                        ErrorMsg = "*** End-Of-Line Expected ***"
                        GoTo ErrorFound
                    End If
                    GoTo Done
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "CURSORAT"
                ' --- cursor at ---
                If TokenNum = LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "," Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "ENTER", "EDIT"
                ' --- enter/edit ---
                Dim EETokenNum As Integer = TokenNum
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Enter/Edit Format ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                SaveTokenNum = TokenNum
                ' --- check for alpha enter/edit ---
                If IsNumExpr(Tokens, TokenNum) Then
                    If Tokens(TokenNum - 1) = ")" Then
                        If IsAlphaTarget(Tokens, TokenNum) Then
                            If TokenNum > LastToken Then
                                If Tokens(EETokenNum) = "ENTER" Then
                                    Tokens(EETokenNum) = "ENTERALPHA"
                                    Changed = True
                                ElseIf Tokens(EETokenNum) = "EDIT" Then
                                    Tokens(EETokenNum) = "EDITALPHA"
                                    Changed = True
                                End If
                                GoTo Done
                            End If
                        End If
                    End If
                End If
                ' --- check for numeric enter/edit ---
                TokenNum = SaveTokenNum
                If IsEnterFmt(Tokens, TokenNum) Then
                    If Tokens(TokenNum - 1) = ")" Then
                        If IsNumTarget(Tokens, TokenNum) Then
                            If TokenNum > LastToken Then
                                If Tokens(EETokenNum) = "ENTER" Then
                                    Tokens(EETokenNum) = "ENTERNUM"
                                    Changed = True
                                ElseIf Tokens(EETokenNum) = "EDIT" Then
                                    Tokens(EETokenNum) = "EDITNUM"
                                    Changed = True
                                End If
                                GoTo Done
                            End If
                        End If
                    End If
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "WHENCANCEL", "WHENESCAPE", "WHENERROR"
                ' --- when vectors ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum - 1) = "WHENERROR" AndAlso Tokens(TokenNum) = "TRAP" Then
                    If TokenNum < LastToken Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                    GoTo Done
                End If
                If Tokens(TokenNum) = "GOTO" Then
                    TokenNum += 1
                    ' --- parse line number ---
                    If TokenNum <> LastToken Then
                        ErrorMsg = "*** End-of-line Expected ***"
                        GoTo ErrorFound
                    End If
                    If Not NumOnly(Tokens(TokenNum)) Then
                        ErrorMsg = "*** Invalid Line Number ***"
                        GoTo ErrorFound
                    End If
                    Tokens(TokenNum) = ":" + Trim$(Str$(Val(Tokens(TokenNum)))) ' change to label format
                    Changed = True
                    GoTo Done
                End If
                If Tokens(TokenNum) = "LOAD" Then
                    TokenNum += 1
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                    If TokenNum <= LastToken Then
                        ErrorMsg = "*** End-of-line Expected ***"
                        GoTo ErrorFound
                    End If
                    GoTo Done
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "SPOOL"
                ' --- spool ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                ' --- check for alpha spool ---
                SaveTokenNum = TokenNum
                If IsNumExpr(Tokens, TokenNum) Then
                    If Tokens(TokenNum - 1) = ")" Then
                        If IsAlphaTerm(Tokens, TokenNum) Then
                            If TokenNum > LastToken Then ' default of "TO W"
                                LastToken = LastToken + 1
                                ReDim Preserve Tokens(LastToken)
                                Tokens(LastToken) = "TO" + vbTab + "W"
                                Changed = True
                                GoTo Done
                            End If
                            If Tokens(TokenNum) = "TO" Then
                                TokenNum += 1
                                If IsAlphaTarget(Tokens, TokenNum) Then
                                    If TokenNum > LastToken Then GoTo Done
                                End If
                            End If
                        End If
                    End If
                End If
                ' --- check for numeric spool ---
                TokenNum = SaveTokenNum
                If IsDisplayFmt(Tokens, TokenNum) Then
                    If Tokens(TokenNum - 1) = ")" Then
                        If IsNumExpr(Tokens, TokenNum) Then
                            If TokenNum > LastToken Then ' default of "TO W"
                                LastToken = LastToken + 1
                                ReDim Preserve Tokens(LastToken)
                                Tokens(LastToken) = "TO" + vbTab + "W"
                                Changed = True
                                GoTo Done
                            End If
                            If Tokens(TokenNum) = "TO" Then
                                TokenNum += 1
                                If IsAlphaTarget(Tokens, TokenNum) Then
                                    If TokenNum > LastToken Then GoTo Done
                                End If
                            End If
                        End If
                    End If
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "PACK"
                ' --- pack ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum - 1) <> ")" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "TO" Then ' insert default string of "R"
                    Tokens(TokenNum - 1) += vbTab + "R"
                    Changed = True
                Else
                    If Not IsAlphaTerm(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Alpha Expression ***"
                        GoTo ErrorFound
                    End If
                End If
                If Tokens(TokenNum) <> "TO" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If IsAlphaTarget(Tokens, TokenNum) Then
                    If TokenNum > LastToken Then GoTo Done
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "MOVE"
                ' --- move ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum - 1) <> ")" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTarget(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Target ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "TO" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If Not IsAlphaTarget(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Target ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then GoTo Done
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "CONVERT"
                ' --- convert ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                SaveTokenNum = TokenNum
                If Not IsEnterFmt(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Enter Format ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum - 1) <> ")" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Expression ***"
                    GoTo ErrorFound
                End If
                ' --- insert default "TO N" if needed ---
                If Tokens(TokenNum) <> "TO" Then
                    Tokens(TokenNum - 1) += vbTab + "TO" + vbTab + "N"
                    Changed = True
                Else
                    ' --- check target ---
                    TokenNum += 1
                    If Not IsNumTarget(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Target ***"
                        GoTo ErrorFound
                    End If
                End If
                ' --- parse line number ---
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If Not NumOnly(Tokens(TokenNum)) Then
                    ErrorMsg = "*** Invalid Line Number ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = ":" + Trim$(Str$(Val(Tokens(TokenNum)))) ' change to label format
                Changed = True
                GoTo Done
            Case "SORT"
                ' --- sort ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum - 1) <> ")" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- check for alpha sort tag ---
                SaveTokenNum = TokenNum
                If IsAlphaTarget(Tokens, TokenNum) Then
                    If TokenNum > LastToken Then GoTo Done
                End If
                ' --- check for numeric sort tag ---
                If IsNumExpr(Tokens, TokenNum) Then
                    If TokenNum > LastToken Then GoTo Done
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "UPDATE"
                ' --- update ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "(" Then ' insert default buffer of "R"
                    Tokens(TokenNum - 1) += vbTab + "R"
                    Changed = True
                Else
                    If Len(Tokens(TokenNum)) <> 1 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    If InStr("RZXYWSTUV", Tokens(TokenNum)) = 0 Then
                        ErrorMsg = "*** Invalid Buffer Name ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                End If
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Len(Tokens(TokenNum)) <> 1 Then
                    ErrorMsg = "*** Invalid Numeric Size ***"
                    GoTo ErrorFound
                End If
                If InStr("123456", Tokens(TokenNum)) = 0 Then
                    ErrorMsg = "*** Invalid Numeric Size ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> ")" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then GoTo Done
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "INIT"
                ' --- init buffer ---
                TokenNum += 1
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ' --- can't init RP2 pointers ---
                If Right$(Tokens(TokenNum), 1) = "2" Then
                    ErrorMsg = "*** Can't Init ?P2 Pointers ***"
                    GoTo ErrorFound
                End If
                ' --- check if buffer pointer ---
                If IsBufferPtrByValue(Tokens(TokenNum)) Then GoTo Done
                ' --- check if just buffer name ---
                If Len(Tokens(TokenNum)) = 1 Then
                    If InStr("RZXYWSTUV", Tokens(TokenNum)) > 0 Then GoTo Done
                End If
                ' --- check if "INIT IR" command ---
                If Len(Tokens(TokenNum)) = 2 Then
                    If Left$(Tokens(TokenNum), 1) = "I" Then
                        If InStr("RZXYWSTUV", Right$(Tokens(TokenNum), 1)) > 0 Then GoTo Done
                    End If
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "SET", "RESET"
                ' --- set/reset buffer pointers ---
                TokenNum += 1
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ' --- can't set/reset IRP pointers ---
                If Left$(Tokens(TokenNum), 1) = "I" Then
                    ErrorMsg = "*** Can't Set/Reset I? Pointers ***"
                    GoTo ErrorFound
                End If
                ' --- can't set/reset RP2 pointers ---
                If Right$(Tokens(TokenNum), 1) = "2" Then
                    ErrorMsg = "*** Can't Set/Reset ?P2 Pointers ***"
                    GoTo ErrorFound
                End If
                ' --- check if buffer pointer ---
                If IsBufferPtrByValue(Tokens(TokenNum)) Then GoTo Done
                ' --- check if just buffer name ---
                If Len(Tokens(TokenNum)) = 1 Then
                    If InStr("RZXYWSTUV", Tokens(TokenNum)) > 0 Then GoTo Done
                End If
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "SKIP"
                ' --- skip buffer ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- check for buffer name ---
                If Len(Tokens(TokenNum)) <> 1 Then
                    ErrorMsg = "*** Invalid Buffer Name ***"
                    GoTo ErrorFound
                End If
                If InStr("RZXYWSTUV", Tokens(TokenNum)) = 0 Then
                    ErrorMsg = "*** Invalid Buffer Name ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then GoTo Done ' SKIP R
                If Tokens(TokenNum) <> "(" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- check for skip alphas ---
                If Tokens(TokenNum) = "A" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Tokens(TokenNum) <> ")" Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo Done ' SKIP R(A)
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                Else ' --- check for skip numerics ---
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                    If Tokens(TokenNum - 1) <> ")" Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                End If
                If TokenNum > LastToken Then GoTo Done
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "DCH"
                ' --- dch ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- check for "DCH (xx) xxx" ---
                If Tokens(TokenNum) = "(" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Len(Tokens(TokenNum)) <> 2 Then
                        ErrorMsg = "*** Invalid Hex Value ***"
                        GoTo ErrorFound
                    End If
                    If InStr("0123456789ABCDEF", Left$(Tokens(TokenNum), 1)) = 0 Then
                        ErrorMsg = "*** Invalid Hex Value ***"
                        GoTo ErrorFound
                    End If
                    If InStr("0123456789ABCDEF", Right$(Tokens(TokenNum), 1)) = 0 Then
                        ErrorMsg = "*** Invalid Hex Value ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    If Tokens(TokenNum) <> ")" Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo Done
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                    If TokenNum > LastToken Then GoTo Done
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                ' --- check for "DCH char xxx" ---
                If Len(Tokens(TokenNum)) <> 1 Then
                    ErrorMsg = "*** Invalid Character Value ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then GoTo Done
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then GoTo Done
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "FLIP", "FLOP", "NOSIGN", "SIGNSET"
                ' --- flip/flop/nosign/signset all have same format ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If Tokens(TokenNum) <> "TO" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumTarget(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Target ***"
                    GoTo ErrorFound
                End If
                If TokenNum > LastToken Then GoTo Done
                ErrorMsg = "*** Invalid Command Syntax ***"
                GoTo ErrorFound
            Case "IF"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- create new IfStack entry ---
                CurrIFEntry = New IFEntry
                IFLabelNum += 1
                With CurrIFEntry
                    .IfType = "IF"
                    .IfLabel = IFLabelNum.ToString
                    .IfHadElse = False
                    .IfAndLevel = 0
                End With
                IfStack.Add(CurrIFEntry) ' add, but keep using
NextIfExpression:
                If Not IsCondition(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Conditional Expression ***"
                    GoTo ErrorFound
                End If
                ' --- check for "IF condition THEN" ---
                If TokenNum <= LastToken Then
                    If Tokens(TokenNum) = "THEN" Then
                        If TokenNum <> LastToken Then
                            ErrorMsg = "*** End-of-line Expected ***"
                            GoTo ErrorFound
                        End If
                        ' --- remove "THEN" and continue with next section ---
                        Tokens(LastToken) = ""
                        LastToken -= 1
                        Changed = True
                    End If
                End If
                ' --- check for "IF condition" ---
                If TokenNum > LastToken Then
                    ' --- rebuild line ---
                    Tokens(TokenNum - 1) += vbTab + "GOTO" + vbTab + ":T" + CurrIFEntry.IfLabel + vbLf
                    Tokens(TokenNum - 1) += "GOTO" + vbTab + ":E" + CurrIFEntry.IfLabel + vbLf
                    Tokens(TokenNum - 1) += ":T" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                    Changed = True
                    GoTo Done
                End If
                ' --- ancient syntax used "IF...IF" instead of "IF...AND" ---
                If Tokens(TokenNum) = "IF" Then
                    Tokens(TokenNum) = "AND"
                    Changed = True
                End If
                ' --- check for conjunctions ---
                If Tokens(TokenNum) = "AND" Then
                    If OrFlag Then ' can't mix AND and OR
                        ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                        GoTo ErrorFound
                    End If
                    AndFlag = True
                    CurrIFEntry.IfAndLevel += 1
                    Tokens(TokenNum) = "GOTO" + vbTab + ":A" + CurrIFEntry.IfLabel + "_" +
                                       CurrIFEntry.IfAndLevel.ToString + vbLf +
                                       "GOTO" + vbTab + ":E" + CurrIFEntry.IfLabel + vbLf +
                                       ":A" + CurrIFEntry.IfLabel + "_" +
                                       CurrIFEntry.IfAndLevel.ToString + vbTab + "IF"
                    Changed = True
                    TokenNum += 1
                    GoTo NextIfExpression
                End If
                If Tokens(TokenNum) = "OR" Then
                    If AndFlag Then ' can't mix AND and OR
                        ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                        GoTo ErrorFound
                    End If
                    OrFlag = True
                    Tokens(TokenNum) = "GOTO" + vbTab + ":T" + CurrIFEntry.IfLabel + vbLf +
                                       "IF"
                    Changed = True
                    TokenNum += 1
                    GoTo NextIfExpression
                End If
                ' --- check for simple "IF...GOTO..." ---
                If (Tokens(TokenNum) = "GOTO") AndAlso (Not AndFlag) AndAlso (Not OrFlag) Then
                    IfStack.RemoveAt(IfStack.Count - 1) ' get rid of unneeded entry
                    GoTo NextCommand
                End If
                ' --- must be "IF condition command" ---
                CompoundCommand = True
                Tokens(TokenNum - 1) += vbTab + "GOTO" + vbTab + ":T" + CurrIFEntry.IfLabel + vbLf +
                                        "GOTO" + vbTab + ":E" + CurrIFEntry.IfLabel + vbLf +
                                        ":T" + CurrIFEntry.IfLabel
                Changed = True
                GoTo NextCommand
            Case "THEN"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> 0 Then ' no line numbers allowed
                    ErrorMsg = "*** Line Number Not Allowed ***"
                    GoTo ErrorFound
                End If
                If IfStack.Count = 0 Then
                    ErrorMsg = "*** Mismatched IF/THEN statements ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry = IfStack.Item(IfStack.Count - 1)
                If CurrIFEntry.IfType <> "IF" Then
                    ErrorMsg = "*** Mismatched IF/THEN statements ***"
                    GoTo ErrorFound
                End If
                ' --- remove "THEN" ---
                Tokens(TokenNum) = ""
                Changed = True
                If TokenNum = LastToken Then GoTo Done
                ' --- has a command after "THEN" ---
                TokenNum += 1
                GoTo NextCommand
            Case "ELSE"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> 0 Then ' no line numbers allowed
                    ErrorMsg = "*** Line Numbers Not Allowed ***"
                    GoTo ErrorFound
                End If
                If IfStack.Count = 0 Then
                    ErrorMsg = "*** Mismatched IF/ELSE statements ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry = IfStack.Item(IfStack.Count - 1)
                If CurrIFEntry.IfType <> "IF" Then
                    ErrorMsg = "*** Mismatched IF/ELSE statements ***"
                    GoTo ErrorFound
                End If
                ' --- handle "ELSE" command ---
                If CurrIFEntry.IfHadElse Then
                    ErrorMsg = "*** Invalid 'IF' Structure ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry.IfHadElse = True
                IfStack.Item(IfStack.Count - 1) = CurrIFEntry ' update IfStack
                ' --- replace "ELSE" with line number handling ---
                Tokens(TokenNum) = "GOTO" + vbTab + ":I" + CurrIFEntry.IfLabel + vbLf +
                                   ":E" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Changed = True
                If TokenNum = LastToken Then GoTo Done
                ' --- has a command after "ELSE" ---
                Tokens(TokenNum) += vbLf
                Changed = True
                TokenNum += 1
                CompoundCommand = True
                GoTo NextCommand
            Case "ENDIF"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                If IfStack.Count = 0 Then
                    ErrorMsg = "*** Mismatched IF/ENDIF statements ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry = IfStack.Item(IfStack.Count - 1)
                If CurrIFEntry.IfType <> "IF" Then
                    ErrorMsg = "*** Mismatched IF/ENDIF statements ***"
                    GoTo ErrorFound
                End If
                ' --- finish up IF construct ---
                If CurrIFEntry.IfHadElse Then
                    Tokens(TokenNum) = ":I" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Else
                    Tokens(TokenNum) = ":E" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                End If
                Changed = True
                IfStack.RemoveAt(IfStack.Count - 1)
                GoTo Done
            Case "FOR"
                SaveTokenNum = TokenNum
                CurrForEntry = New ForEntry
                TokenNum += 1
                ' --- get loop variable ---
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                TempSaveToken = TokenNum
                If Not IsNumTarget(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Target ***"
                    GoTo ErrorFound
                End If
                CurrForEntry.ForVar = GetExpr(Tokens, TempSaveToken, TokenNum - 1) ' loop variable
                ' --- check for "=" ---
                If Tokens(TokenNum) <> "=" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                ' --- get from value ---
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                TempSaveToken = TokenNum
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                CurrForEntry.ForFrom = GetExpr(Tokens, TempSaveToken, TokenNum - 1) ' from value
                ' --- check for "TO" ---
                If Tokens(TokenNum) <> "TO" Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                TokenNum += 1
                ' --- get to value ---
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                TempSaveToken = TokenNum
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                CurrForEntry.ForTo = GetExpr(Tokens, TempSaveToken, TokenNum - 1) ' to value
                ' --- get step value ---
                CurrForEntry.ForStep = "1"
                If TokenNum < LastToken Then
                    If Tokens(TokenNum) <> "BY" AndAlso Tokens(TokenNum) <> "STEP" Then
                        ErrorMsg = "*** Invalid Command Syntax ***"
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        ErrorMsg = "*** Missing Parameters ***"
                        GoTo ErrorFound
                    End If
                    TempSaveToken = TokenNum
                    If Not IsNumExpr(Tokens, TokenNum) Then
                        ErrorMsg = "*** Invalid Numeric Expression ***"
                        GoTo ErrorFound
                    End If
                    CurrForEntry.ForStep = GetExpr(Tokens, TempSaveToken, TokenNum - 1) ' step value
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ForLabelNum += 1
                CurrForEntry.ForLabel = ForLabelNum.ToString
                ForStack.Add(CurrForEntry) ' keep using it below
                ' --- rebuild line into Assignment:If-Goto format ---
                Result = New StringBuilder
                Result.Append(GetExpr(Tokens, 0, SaveTokenNum - 1)) ' get anything before "FOR"
                If Result.Length > 0 Then Result.Append(vbTab)
                Result.Append(CurrForEntry.ForVar + vbTab + "=" + vbTab + CurrForEntry.ForFrom)
                Result.Append(vbLf) ' next line
                Result.Append(":F" + CurrForEntry.ForLabel + vbTab)
                Result.Append("IF" + vbTab + CurrForEntry.ForVar + vbTab)
                If CurrForEntry.ForStep.StartsWith("-" + vbTab) Then
                    Result.Append("<") ' reverse for
                Else
                    Result.Append(">") ' normal for
                End If
                Result.Append(vbTab + CurrForEntry.ForTo + vbTab + "GOTO" + vbTab)
                Result.Append(":N" + CurrForEntry.ForLabel)
                CurrForEntry = Nothing
                GoTo DoneReturn
            Case "NEXT"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If ForStack.Count <= 0 Then
                    ErrorMsg = "*** 'NEXT' Without Matching 'FOR' ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ' --- Get current For/Next entry from stack ---
                CurrForEntry = ForStack.Item(ForStack.Count - 1)
                ForStack.RemoveAt(ForStack.Count - 1)
                ' --- rebuild command into Increment:Goto format ---
                Result = New StringBuilder
                Result.Append(GetExpr(Tokens, 0, TokenNum - 1)) ' get anything before "NEXT"
                If Result.Length > 0 Then Result.Append(vbTab)
                Result.Append(CurrForEntry.ForVar + vbTab + "=" + vbTab + CurrForEntry.ForVar)
                Result.Append(vbTab + "+" + vbTab + CurrForEntry.ForStep + vbLf)
                Result.Append("GOTO" + vbTab + ":F" + CurrForEntry.ForLabel + vbLf)
                Result.Append(":N" + CurrForEntry.ForLabel + vbTab + "\NOP")
                GoTo DoneReturn
            Case "REPEAT"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ' --- create new IfStack entry ---
                CurrIFEntry = New IFEntry
                With CurrIFEntry
                    .IfType = "REPEAT"
                    IFLabelNum = IFLabelNum + 1
                    .IfLabel = Trim$(Str$(IFLabelNum))
                    .IfHadElse = False
                    .IfAndLevel = 0
                End With
                IfStack.Add(CurrIFEntry) ' add, but keep using CurrIfEntry
                ' --- rebuild command into Label format ---
                If TokenNum > 0 Then
                    Tokens(TokenNum - 1) += vbTab + "\NOP" + vbLf
                End If
                Tokens(TokenNum) = ":R" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Changed = True
                GoTo Done
            Case "UNTIL"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If IfStack.Count <= 0 Then
                    ErrorMsg = "*** Mismatched REPEAT/UNTIL statements ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry = IfStack.Item(IfStack.Count - 1)
                IfStack.RemoveAt(IfStack.Count - 1)
                If CurrIFEntry.IfType <> "REPEAT" Then
                    ErrorMsg = "*** Mismatched REPEAT/UNTIL statements ***"
                    GoTo ErrorFound
                End If
                Tokens(TokenNum) = "IF" ' change UNTIL to IF
                Changed = True
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
NextUntilExpression:
                If Not IsCondition(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Conditional Expression ***"
                    GoTo ErrorFound
                End If
                ' --- check for conjunctions ---
                If TokenNum <= LastToken Then
                    If Tokens(TokenNum) = "AND" Then
                        If OrFlag Then
                            ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                            GoTo ErrorFound
                        End If ' can't mix AND and OR
                        AndFlag = True
                        CurrIFEntry.IfAndLevel += 1
                        Tokens(TokenNum) = "GOTO" + vbTab + ":A" + CurrIFEntry.IfLabel + "_" +
                                           CurrIFEntry.IfAndLevel.ToString + vbLf +
                                           "GOTO" + vbTab + ":R" + CurrIFEntry.IfLabel + vbLf +
                                           ":A" + CurrIFEntry.IfLabel + "_" +
                                           CurrIFEntry.IfAndLevel.ToString + vbTab + "IF"
                        Changed = True
                        TokenNum += 1
                        GoTo NextUntilExpression
                    End If
                    If Tokens(TokenNum) = "OR" Then
                        If AndFlag Then ' can't mix AND and OR
                            ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                            GoTo ErrorFound
                        End If
                        OrFlag = True
                        Tokens(TokenNum) = "GOTO" + vbTab + ":U" + CurrIFEntry.IfLabel + vbLf +
                                           "IF"
                        Changed = True
                        TokenNum += 1
                        GoTo NextUntilExpression
                    End If
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                ' --- finish UNTIL line ---
                Tokens(LastToken) += vbTab + "GOTO" + vbTab + ":U" + CurrIFEntry.IfLabel + vbLf +
                                     "GOTO" + vbTab + ":R" + CurrIFEntry.IfLabel + vbLf +
                                     ":U" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Changed = True
                GoTo Done
            Case "WHILE"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                ' --- create new IfStack entry ---
                CurrIFEntry = New IFEntry
                With CurrIFEntry
                    .IfType = "WHILE"
                    IFLabelNum = IFLabelNum + 1
                    .IfLabel = Trim$(Str$(IFLabelNum))
                    .IfHadElse = False
                    .IfAndLevel = 0
                End With
                IfStack.Add(CurrIFEntry) ' add, but keep using CurrIfEntry
                ' --- change WHILE to IF ---
                If TokenNum > 0 Then
                    Tokens(TokenNum - 1) = Tokens(TokenNum - 1) + vbTab + "\NOP" + vbLf
                End If
                Tokens(TokenNum) = ":W" + CurrIFEntry.IfLabel + vbTab + "IF"
                Changed = True
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                ' --- check all conditions ---
NextWhileExpression:
                If Not IsCondition(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Conditional Expression ***"
                    GoTo ErrorFound
                End If
                ' --- check for conjunctions ---
                If TokenNum <= LastToken Then
                    If Tokens(TokenNum) = "AND" Then
                        If OrFlag Then ' can't mix AND and OR
                            ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                            GoTo ErrorFound
                        End If
                        AndFlag = True
                        CurrIFEntry.IfAndLevel = CurrIFEntry.IfAndLevel + 1
                        Tokens(TokenNum) = "GOTO" + vbTab + ":A" + CurrIFEntry.IfLabel + "_" +
                                           CurrIFEntry.IfAndLevel.ToString + vbLf +
                                           "GOTO" + vbTab + ":D" + CurrIFEntry.IfLabel + vbLf +
                                           ":A" + CurrIFEntry.IfLabel + "_" +
                                           CurrIFEntry.IfAndLevel.ToString + vbTab + "IF"
                        Changed = True
                        TokenNum += 1
                        GoTo NextWhileExpression
                    End If
                    If Tokens(TokenNum) = "OR" Then
                        If AndFlag Then ' can't mix AND and OR
                            ErrorMsg = "*** Can't Mix 'AND' And 'OR' Expressions ***"
                            GoTo ErrorFound
                        End If
                        OrFlag = True
                        Tokens(TokenNum) = "GOTO" + vbTab + ":V" + CurrIFEntry.IfLabel + vbLf +
                                           "IF"
                        Changed = True
                        TokenNum += 1
                        GoTo NextWhileExpression
                    End If
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                ' --- rebuild line ---
                Tokens(LastToken) += vbTab + "GOTO" + vbTab + ":V" + CurrIFEntry.IfLabel + vbLf +
                                     "GOTO" + vbTab + ":D" + CurrIFEntry.IfLabel + vbLf +
                                     ":V" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Changed = True
                GoTo Done
            Case "DO"
                If CompoundCommand Then
                    ErrorMsg = "*** Invalid Compound Command ***"
                    GoTo ErrorFound
                End If
                If IfStack.Count = 0 Then
                    ErrorMsg = "*** Mismatched WHILE/DO statements ***"
                    GoTo ErrorFound
                End If
                If TokenNum <> LastToken Then
                    ErrorMsg = "*** End-of-line Expected ***"
                    GoTo ErrorFound
                End If
                CurrIFEntry = IfStack.Item(IfStack.Count - 1)
                IfStack.RemoveAt(IfStack.Count - 1)
                If CurrIFEntry.IfType <> "WHILE" Then
                    ErrorMsg = "*** Mismatched WHILE/DO statements ***"
                    GoTo ErrorFound
                End If
                ' --- rebuild line ---
                Tokens(TokenNum) = "GOTO" + vbTab + ":W" + CurrIFEntry.IfLabel + vbLf +
                                   ":D" + CurrIFEntry.IfLabel + vbTab + "\NOP"
                Changed = True
                GoTo Done
            Case "HEX"
                ' --- hex commands (ignored while running) ---
                GoTo Done
            Case "FATALERROR"
                ' --- new IDRIS commands ---
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "EXITRUNTIME"
                TokenNum += 1
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "SETSUBQUERYFILE"
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "SETSUBQUERY"
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "EXECSQL"
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsAlphaTerm(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Alpha Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "SAVEFILEINFO"
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case "RESTOREFILEINFO"
                TokenNum += 1
                If TokenNum > LastToken Then
                    ErrorMsg = "*** Missing Parameters ***"
                    GoTo ErrorFound
                End If
                If Not IsNumExpr(Tokens, TokenNum) Then
                    ErrorMsg = "*** Invalid Numeric Expression ***"
                    GoTo ErrorFound
                End If
                If TokenNum <= LastToken Then
                    ErrorMsg = "*** Invalid Command Syntax ***"
                    GoTo ErrorFound
                End If
                GoTo Done
            Case Else
                ' --- nothing matches ---
                ErrorMsg = "*** Unknown command ***"
                GoTo ErrorFound
        End Select
        ErrorMsg = "*** Unknown command ***"
        GoTo ErrorFound
Done:
        ' --- check if this was a compound command ---
        If CompoundCommand Then
            CurrIFEntry = IfStack.Item(IfStack.Count - 1)
            If CurrIFEntry.IfHadElse Then
                Tokens(LastToken) += vbLf + ":I" + CurrIFEntry.IfLabel + vbTab + "\NOP"
            Else
                Tokens(LastToken) += vbLf + ":E" + CurrIFEntry.IfLabel + vbTab + "\NOP"
            End If
            Changed = True
            IfStack.RemoveAt(IfStack.Count - 1)
        End If
        ' --- check if changed ---
        If Not Changed Then
            Return Value
        End If
        ' --- Recombine tokens into a single line ---
        Result = New StringBuilder
        For TokenNum = 0 To LastToken
            If Tokens(TokenNum) <> "" Then
                If Result.Length > 0 Then
                    If Not Tokens(TokenNum - 1).EndsWith(vbLf) Then
                        Result.Append(vbTab)
                    End If
                End If
                Result.Append(Tokens(TokenNum))
            End If
        Next
DoneReturn:
        Return Result.ToString
ErrorFound:
        ' --- Mark any errors with "***" ---
        Return ErrorMsg
    End Function

    Private Function IsAlphaTarget(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim IsBuffer As Boolean = False
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        ' --- Check for simple registers ---
        If Tokens(TokenNum) = "KEY" OrElse
           Tokens(TokenNum) = "DATE" OrElse
           Tokens(TokenNum) = "A" OrElse
           Tokens(TokenNum) = "B" OrElse
           Tokens(TokenNum) = "C" OrElse
           Tokens(TokenNum) = "D" OrElse
           Tokens(TokenNum) = "E" Then
            TokenNum += 1
            GoTo OffsetCheck
        End If
        ' --- Check for A1 to E9 ---
        If "ABCDE".IndexOf(Tokens(TokenNum).Substring(0, 1)) >= 0 Then
            If Tokens(TokenNum).Length <> 2 Then GoTo ErrorFound
            If Not NumOnly(Tokens(TokenNum).Substring(1, 1)) Then GoTo ErrorFound
            TokenNum += 1
            GoTo OffsetCheck
        End If
        ' --- Check for buffers ---
        If Tokens(TokenNum).Length = 1 Then
            If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) >= 0 Then
                TokenNum += 1
                IsBuffer = True
                GoTo OffsetCheck
            End If
        End If
        ' --- Not an alpha target ---
        GoTo ErrorFound
OffsetCheck:
        If TokenNum > LastToken Then GoTo Done
        If Tokens(TokenNum) <> "[" Then GoTo BufferCheck
        TokenNum += 1
        If Not IsNumExpr(Tokens, TokenNum) Then GoTo ErrorFound
        If Tokens(TokenNum - 1) <> "]" Then GoTo ErrorFound
        GoTo BufferCheck
BufferCheck:
        If IsBuffer Then
            If TokenNum > LastToken - 2 Then GoTo Done
            If Tokens(TokenNum) <> "(" Then GoTo Done
            TokenNum += 1
            If Tokens(TokenNum) <> "A" Then GoTo ErrorFound
            TokenNum += 1
            If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
            TokenNum += 1
        End If
        GoTo Done
Done:
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsNumExpr(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim CurrOper As String
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        If Not IsNumTerm(Tokens, TokenNum) Then GoTo ErrorFound
NextOperator:
        If TokenNum > LastToken Then GoTo Done
        CurrOper = Tokens(TokenNum)
        ' --- check if need to gobble up but not use token ---
        If CurrOper = ")" OrElse CurrOper = "]" Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- check if valid operator ---
        If CurrOper <> "+" AndAlso
           CurrOper <> "-" AndAlso
           CurrOper <> "*" AndAlso
           CurrOper <> "/" Then
            GoTo Done
        End If
        ' --- use the token ---
        TokenNum += 1
        ' --- check for next numeric term ---
        If Not IsNumTerm(Tokens, TokenNum) Then GoTo ErrorFound
        ' --- get next operator ---
        GoTo NextOperator
Done:
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsNumTerm(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        ' --- check for unary minus ---
        If Tokens(TokenNum) = "-" Then
            TokenNum += 1
            If TokenNum > LastToken Then GoTo ErrorFound
        End If
        ' --- check for beginning of an expression ---
        If Tokens(TokenNum) = "(" Then
            TokenNum += 1
            If Not IsNumExpr(Tokens, TokenNum) Then GoTo ErrorFound
            If Tokens(TokenNum - 1) <> ")" Then GoTo ErrorFound
            GoTo Done
        End If
        ' --- check for a number ---
        If NumOnly(Tokens(TokenNum)) Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- check for internal constants and numeric functions ---
        If Tokens(TokenNum) = "TRUE" OrElse
           Tokens(TokenNum) = "FALSE" OrElse
           Tokens(TokenNum) = "GETDATEVAL" OrElse
           Tokens(TokenNum) = "GETTIMEVAL" OrElse
           Tokens(TokenNum) = "GETCOMPANYNAMELEN" OrElse
           Tokens(TokenNum) = "GETREADONLY" OrElse
           Tokens(TokenNum) = "GETTIMEHUND" Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- check for numeric target ---
        If IsNumTarget(Tokens, TokenNum) Then GoTo Done
        ' --- not a numeric term ---
        GoTo ErrorFound
Done:
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsNumTarget(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim TempValue As Integer
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        ' --- check for simple targets ---
        If Tokens(TokenNum) = "REC" OrElse
           Tokens(TokenNum) = "REM" Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- check for numerics that might have offsets - n[x], f[x], g[x] ---
        If Tokens(TokenNum) = "N" OrElse
           Tokens(TokenNum) = "F" OrElse
           Tokens(TokenNum) = "G" Then
            TokenNum += 1
            If TokenNum > LastToken Then GoTo Done
            If Tokens(TokenNum) <> "[" Then GoTo Done
            TokenNum += 1
            If Not IsNumExpr(Tokens, TokenNum) Then GoTo ErrorFound
            If Tokens(TokenNum - 1) <> "]" Then GoTo ErrorFound
            GoTo Done
        End If
        ' --- check for numerics up to N99, F99, G9 ---
        If Tokens(TokenNum).Length > 1 Then
            If "NFG".IndexOf(Tokens(TokenNum).Substring(0, 1)) >= 0 Then
                If Not NumOnly(Tokens(TokenNum).Substring(1)) Then GoTo ErrorFound
                TempValue = Integer.Parse(Tokens(TokenNum).Substring(1))
                If TempValue < 1 OrElse TempValue > 99 Then GoTo ErrorFound
                If Tokens(TokenNum).StartsWith("G") Then
                    If TempValue > 9 Then GoTo ErrorFound
                End If
                TokenNum += 1
                GoTo Done
            End If
        End If
        ' --- numeric buffer fields ---
        If Tokens(TokenNum).Length = 1 Then
            If "RZXYWSTUV".IndexOf(Tokens(TokenNum)) >= 0 Then
                TokenNum += 1
                If TokenNum > LastToken Then GoTo ErrorFound
                If Tokens(TokenNum) = "[" Then ' found offset
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo ErrorFound
                    If Not IsNumExpr(Tokens, TokenNum) Then GoTo ErrorFound
                    If Tokens(TokenNum - 1) <> "]" Then GoTo ErrorFound
                    If TokenNum > LastToken Then GoTo ErrorFound
                End If
                If Tokens(TokenNum) = "(" Then ' check for field size
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo ErrorFound
                    If Not NumOnly(Tokens(TokenNum)) Then GoTo ErrorFound
                    TempValue = Integer.Parse(Tokens(TokenNum))
                    If TempValue < 1 OrElse TempValue > 6 Then GoTo ErrorFound
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo ErrorFound
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    GoTo Done
                End If
                GoTo ErrorFound
            End If
        End If
        ' --- sysvar access function ---
        If Tokens(TokenNum) = "SYSVAR" Then
            TokenNum += 1
            If TokenNum > LastToken Then GoTo ErrorFound
            If Tokens(TokenNum) <> "(" Then GoTo ErrorFound ' check for byte number expression
            TokenNum += 1
            If TokenNum > LastToken Then GoTo ErrorFound
            If Not IsNumExpr(Tokens, TokenNum) Then GoTo ErrorFound
            If Tokens(TokenNum - 1) <> ")" Then GoTo ErrorFound
            GoTo Done
        End If
        ' --- numeric system variables ---
        If IsBufferPtrByValue(Tokens(TokenNum)) Then
            TokenNum += 1
            GoTo Done
        End If
        If IsSystemVarByValue(Tokens(TokenNum)) Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- not a numeric target ---
        GoTo ErrorFound
Done:
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsAlphaTerm(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
NextTerm:
        If TokenNum > LastToken Then GoTo ErrorFound
        ' --- string constant ---
        If """'%$".IndexOf(Tokens(TokenNum).Substring(0, 1)) >= 0 Then ' string
            If Not Tokens(TokenNum).EndsWith(Tokens(TokenNum).Substring(0, 1)) Then GoTo ErrorFound
            TokenNum += 1
            GoTo Done
        End If
        ' --- check for internal constants and functions ---
        If Tokens(TokenNum) = "NULL" OrElse
           Tokens(TokenNum) = "GETCLIENTLIST" OrElse
           Tokens(TokenNum) = "GETDATESTR" OrElse
           Tokens(TokenNum) = "GETENVIRONMENT" OrElse
           Tokens(TokenNum) = "GETTIMESTR" OrElse
           Tokens(TokenNum) = "GETLOGINID" OrElse
           Tokens(TokenNum) = "GETCOMPANYNAME" OrElse
           Tokens(TokenNum) = "GETCOMPANYINITIALS" Then
            TokenNum += 1
            GoTo Done
        End If
        ' --- all alpha targets are valid terms ---
        If IsAlphaTarget(Tokens, TokenNum) Then GoTo Done
        ' --- doesn't match any alpha pattern ---
        GoTo ErrorFound
Done:
        ' --- check for string concatenation ---
        If TokenNum < LastToken AndAlso Tokens(TokenNum) = "&" Then
            TokenNum += 1
            GoTo NextTerm
        End If
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsDisplayFmt(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim intPos As Integer
        Dim HasNeg As Boolean
        Dim HasDecimal As Boolean
        Dim HasAbove As Boolean
        Dim HasBelow As Boolean
        Dim HasComma As Boolean
        Dim HasZeroFill As Boolean
        Dim HasStarFill As Boolean
        Dim HasLeftParen As Boolean
        Dim HasRightParen As Boolean
        Dim intAboveVal As Integer
        Dim intBelowVal As Integer
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        HasNeg = False
        HasDecimal = False
        HasAbove = False
        HasBelow = False
        HasComma = False
        HasZeroFill = False
        HasStarFill = False
        HasLeftParen = False
        HasRightParen = False
NextItem:
        If Tokens(TokenNum) = "-" Then
            If HasNeg OrElse HasRightParen OrElse HasLeftParen Then GoTo ErrorFound
            HasNeg = True
            GoTo DoneItem
        End If
        If Tokens(TokenNum) = "(" Then
            If HasNeg OrElse HasLeftParen Then GoTo ErrorFound
            HasLeftParen = True
            GoTo DoneItem
        End If
        If Tokens(TokenNum) = ")" Then
            If HasLeftParen = HasRightParen Then
                If Not HasAbove Then GoTo ErrorFound
                TokenNum += 1
                GoTo Done ' this is the way out
            End If
            If Not HasLeftParen Then GoTo ErrorFound
            HasRightParen = True
            GoTo DoneItem
        End If
        If Tokens(TokenNum) = "*" Then
            If HasComma OrElse HasZeroFill OrElse HasStarFill Then GoTo ErrorFound
            HasStarFill = True
            GoTo DoneItem
        End If
        If Tokens(TokenNum) = "," Then
            If HasComma OrElse HasZeroFill OrElse HasStarFill Then GoTo ErrorFound
            HasComma = True
            GoTo DoneItem
        End If
        If InStr(Tokens(TokenNum), ".") > 0 Then
            intPos = InStr(Tokens(TokenNum), ".")
            If intPos > 1 Then
                If HasAbove Then GoTo ErrorFound
                If Left$(Tokens(TokenNum), 1) = "Z" Then
                    If HasComma OrElse HasZeroFill OrElse HasStarFill Then GoTo ErrorFound
                    HasZeroFill = True
                    If Not NumOnly(Mid$(Tokens(TokenNum), 2, intPos - 2)) Then GoTo ErrorFound
                    intAboveVal = CInt(Val(Mid$(Tokens(TokenNum), 2, intPos - 2)))
                Else
                    If Not NumOnly(Left$(Tokens(TokenNum), intPos - 1)) Then GoTo ErrorFound
                    intAboveVal = CInt(Val(Left$(Tokens(TokenNum), intPos - 1)))
                End If
                If intAboveVal > 14 Then GoTo ErrorFound
                HasAbove = True
            End If
            If HasDecimal Then GoTo ErrorFound
            HasDecimal = True
            If intPos < Len(Tokens(TokenNum)) Then
                If HasBelow Then GoTo ErrorFound
                If Not NumOnly(Mid$(Tokens(TokenNum), intPos + 1)) Then GoTo ErrorFound
                intBelowVal = CInt(Val(Mid$(Tokens(TokenNum), intPos + 1)))
                If intBelowVal > 7 Then GoTo ErrorFound
                If intAboveVal + intBelowVal > 14 Then GoTo ErrorFound
                HasBelow = True
            End If
            GoTo DoneItem
        End If
        If Left$(Tokens(TokenNum), 1) = "Z" Then
            If HasComma OrElse HasZeroFill OrElse HasStarFill Then GoTo ErrorFound
            HasZeroFill = True
            If Len(Tokens(TokenNum)) > 1 Then ' may combine Z with digits above
                If HasAbove Then GoTo ErrorFound
                If Not NumOnly(Mid$(Tokens(TokenNum), 2)) Then GoTo ErrorFound
                If Val(Mid$(Tokens(TokenNum), 2)) > 14 Then GoTo ErrorFound
                intAboveVal = CInt(Val(Mid$(Tokens(TokenNum), 2)))
                HasAbove = True
            End If
            GoTo DoneItem
        End If
        If NumOnly(Tokens(TokenNum)) Then
            If Not HasAbove Then
                intAboveVal = CInt(Val(Tokens(TokenNum)))
                If intAboveVal > 14 Then GoTo ErrorFound
                HasAbove = True
            Else
                If Not HasDecimal Then GoTo ErrorFound
                If HasBelow Then GoTo ErrorFound
                intBelowVal = CInt(Val(Tokens(TokenNum)))
                If intBelowVal > 7 Then GoTo ErrorFound
                If intAboveVal + intBelowVal > 14 Then GoTo ErrorFound
                HasBelow = True
            End If
            GoTo DoneItem
        End If
        GoTo ErrorFound
DoneItem:
        TokenNum += 1
        If TokenNum > LastToken Then GoTo ErrorFound
        GoTo NextItem
Done:
        IsDisplayFmt = True
        Exit Function
ErrorFound:
        TokenNum = SaveTokenNum
        IsDisplayFmt = False
        Exit Function
    End Function

    Private Function IsEnterFmt(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim HasNeg As Boolean
        Dim HasDecimal As Boolean
        Dim HasAbove As Boolean
        Dim HasBelow As Boolean
        Dim intPos As Integer
        Dim intAboveVal As Integer
        Dim intBelowVal As Integer
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        SaveTokenNum = TokenNum ' restore if error
        If TokenNum > LastToken Then GoTo ErrorFound
        HasNeg = False
        HasDecimal = False
        HasAbove = False
        HasBelow = False
NextItem:
        If Tokens(TokenNum) = "-" Then
            If HasNeg Then GoTo ErrorFound
            HasNeg = True
            GoTo DoneItem
        End If
        If InStr(Tokens(TokenNum), ".") > 0 Then
            intPos = InStr(Tokens(TokenNum), ".")
            If intPos > 1 Then
                If HasAbove Then GoTo ErrorFound
                If Not NumOnly(Left$(Tokens(TokenNum), intPos - 1)) Then GoTo ErrorFound
                intAboveVal = CInt(Val(Left$(Tokens(TokenNum), intPos - 1)))
                If intAboveVal > 14 Then GoTo ErrorFound
                HasAbove = True
            End If
            If HasDecimal Then GoTo ErrorFound
            HasDecimal = True
            If intPos < Len(Tokens(TokenNum)) Then
                If HasBelow Then GoTo ErrorFound
                If Not NumOnly(Mid$(Tokens(TokenNum), intPos + 1)) Then GoTo ErrorFound
                intBelowVal = CInt(Val(Mid$(Tokens(TokenNum), intPos + 1)))
                If intBelowVal > 7 Then GoTo ErrorFound
                If intAboveVal + intBelowVal > 14 Then GoTo ErrorFound
                HasBelow = True
            End If
            GoTo DoneItem
        End If
        If NumOnly(Tokens(TokenNum)) Then
            If Not HasAbove Then
                intAboveVal = CInt(Val(Tokens(TokenNum)))
                If intAboveVal > 14 Then GoTo ErrorFound
                HasAbove = True
            Else
                If Not HasDecimal Then GoTo ErrorFound
                If HasBelow Then GoTo ErrorFound
                intBelowVal = CInt(Val(Tokens(TokenNum)))
                If intBelowVal > 7 Then GoTo ErrorFound
                If intAboveVal + intBelowVal > 14 Then GoTo ErrorFound
                HasBelow = True
            End If
            GoTo DoneItem
        End If
        If Tokens(TokenNum) = ")" Then
            If Not HasAbove Then GoTo ErrorFound
            TokenNum += 1
            GoTo Done
        End If
DoneItem:
        TokenNum += 1
        If TokenNum > LastToken Then GoTo ErrorFound
        GoTo NextItem
Done:
        IsEnterFmt = True
        Exit Function
ErrorFound:
        TokenNum = SaveTokenNum
        IsEnterFmt = False
        Exit Function
    End Function

    Private Function IsCondition(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        Dim SaveTokenNum As Integer = TokenNum
        Dim LastToken As Integer = Tokens.GetUpperBound(0)
        ' ------------------------------------------------
        If TokenNum > LastToken Then GoTo ErrorFound
        ' --- check for alpha comparison ---
        If IsAlphaTerm(Tokens, TokenNum) Then
            If IsRelation(Tokens, TokenNum) Then
                If IsAlphaTerm(Tokens, TokenNum) Then
                    GoTo Done
                End If
            End If
        End If
        ' --- reset to beginning of expression ---
        TokenNum = SaveTokenNum
        ' --- check for numeric comparison ---
        If IsNumExpr(Tokens, TokenNum) Then
            If IsRelation(Tokens, TokenNum) Then
                If IsNumExpr(Tokens, TokenNum) Then
                    GoTo Done
                End If
            End If
        End If
        ' --- doesn't match any condition pattern ---
        GoTo ErrorFound
Done:
        Return True
ErrorFound:
        TokenNum = SaveTokenNum
        Return False
    End Function

    Private Function IsRelation(ByRef Tokens() As String, ByRef TokenNum As Integer) As Boolean
        TokenNum += 1
        Select Case Tokens(TokenNum - 1)
            Case "=" : Return True
            Case "#" : Return True
            Case "<>" : Return True
            Case ">" : Return True
            Case "<" : Return True
            Case ">=" : Return True
            Case "<=" : Return True
        End Select
        TokenNum -= 1
        Return False
    End Function

    Private Function GetExpr(ByRef Tokens() As String, ByVal FromToken As Integer, ByVal ToToken As Integer) As String
        ' --- this returns a section of the line in a single string ---
        ' --- it doesn't adjust any pointers or any Token items.    ---
        Dim Result As New StringBuilder
        ' -----------------------------
        For LoopNum As Integer = FromToken To ToToken
            If Result.Length > 0 Then
                Result.Append(vbTab)
            End If
            Result.Append(Tokens(LoopNum))
        Next
        Return Result.ToString
    End Function

#End Region

#Region " --- Third Pass Routines --- "

    Private Function PerformThirdPass(ByRef Lines1 As List(Of String), ByRef Lines2 As List(Of String)) As Boolean
        Dim CurrLine As String
        Dim Tokens() As String
        Dim LineNum As Integer = -1
        Dim LineComment As String = ""
        Dim ErrorInfo As ParseError
        ' ----------------------------
        Renumbers.Clear()
        Lines2.Clear()
        ' --- build a list of :labels and line numbers ---
        For LoopNum As Integer = 0 To Lines1.Count - 1
            CurrLine = Lines1(LoopNum)
            If CurrLine = "" Then Continue For
            If CurrLine.StartsWith("!!") Then
                LineComment = CurrLine.Substring(1) ' drop first "!"
                Continue For
            End If
            If CurrLine.StartsWith("!") Then
                Lines2.Add(CurrLine)
                Continue For
            End If
            LineNum += 1
            If LineComment <> "" Then
                CurrLine += vbTab + LineComment
                LineComment = ""
            End If
            Tokens = CurrLine.Split(CChar(vbTab))
            If Tokens(0).StartsWith(":") Then
                Try
                    Renumbers.Add(Tokens(0), LineNum.ToString)
                Catch ex As Exception
                    ErrorInfo = New ParseError
                    With ErrorInfo
                        .LineNum = LoopNum
                        .SourceLine = Lines1(LoopNum).Replace(":", "")
                        .ErrorDesc = "Duplicate Line Number Found"
                    End With
                    ParseErrors.Add(ErrorInfo)
                    ErrorInfo = Nothing
                    Return False
                End Try
            End If
            Lines2.Add(CurrLine)
        Next
        ' --- now replace the line numbers ---
        For LoopNum As Integer = 0 To Lines2.Count - 1
            Try
                Lines2(LoopNum) = ReplaceLineNumbers(Lines1(LoopNum))
            Catch ex As Exception
                ErrorInfo = New ParseError
                With ErrorInfo
                    .LineNum = LoopNum
                    .SourceLine = Lines1(LoopNum).Replace(":", "")
                    .ErrorDesc = "Line Number Not Found"
                End With
                ParseErrors.Add(ErrorInfo)
                ErrorInfo = Nothing
                Return False
            End Try
        Next
        Return True
    End Function

    Private Function ReplaceLineNumbers(ByVal Value As String) As String
        Dim Changed As Boolean = False
        Dim TokenNum As Integer
        Dim Tokens() As String = Value.Split(CChar(vbTab))
        ' -----------------------------------------
        For TokenNum = 0 To Tokens.GetUpperBound(0)
            If Tokens(TokenNum) = "!!" OrElse Tokens(TokenNum) = "!" Then
                Exit For
            End If
            If Tokens(TokenNum).StartsWith(":") Then
                Tokens(TokenNum) = Renumbers.Item(Tokens(TokenNum))
                Changed = True
            End If
        Next
        If Changed Then
            Dim Result As New StringBuilder
            For TokenNum = 0 To Tokens.GetUpperBound(0)
                If TokenNum > 0 Then
                    Result.Append(vbTab)
                End If
                Result.Append(Tokens(TokenNum))
            Next
            Return Result.ToString
        Else
            Return Value
        End If
    End Function

#End Region

#Region " --- Fourth Pass Routines --- "

    Private Function PerformFourthPass(ByRef Lines1 As List(Of String), ByRef Lines2 As List(Of String)) As Boolean
        Dim CurrLine As String = ""
        Dim Tokens() As String
        Dim TokenNum As Integer
        Dim LastToken As Integer
        Dim SaveTokenNum As Integer
        Dim SaveTokenNum2 As Integer
        Dim Result As StringBuilder
        Dim TempResult As String
        Dim TempResult2 As String
        Dim TempResult3 As String
        Dim TempResult4 As String
        Dim NeedEndif As Boolean
        Dim NeedExitSub As Boolean
        Dim NeedThisLineNum As Boolean
        Dim NeedNextLineNum As Boolean
        Dim SourceLinenum As Integer
        Dim ErrorMsg As String = ""
        ' ----------------------------
        Lines2.Clear()
        JumpPointList.Clear()
        GotoLineList.Clear()
        ' --- make sure there is always a line number 0 ---
        NeedThisLineNum = True
        JumpPointList.Add(0)
        ' --- clean up remaining items into good IL code ---
        For SourceLinenum = 0 To Lines1.Count - 1
            CurrLine = Lines1(SourceLinenum)

            ' --- check for pre-comments ---
            If CurrLine.StartsWith("!") Then
                Continue For
            End If

            Tokens = CurrLine.Split(CChar(vbTab))

            ' --- find the end of the command ---
            LastToken = Tokens.GetUpperBound(0)
            For TokenNum = 0 To Tokens.GetUpperBound(0)
                If Tokens(TokenNum) = "!" Then
                    LastToken = TokenNum - 1
                    Exit For
                End If
            Next

            Result = New StringBuilder
            NeedEndif = False
            NeedExitSub = False
            NeedNextLineNum = False

            For TokenNum = 0 To LastToken
                If (Tokens(TokenNum).StartsWith("'"c) AndAlso Tokens(TokenNum).EndsWith("'"c)) OrElse
                   (Tokens(TokenNum).StartsWith("%"c) AndAlso Tokens(TokenNum).EndsWith("%"c)) OrElse
                   (Tokens(TokenNum).StartsWith("$"c) AndAlso Tokens(TokenNum).EndsWith("$"c)) Then
                    Tokens(TokenNum) = """" + Replace(Tokens(TokenNum).Substring(1, Tokens(TokenNum).Length - 2), """", """""") + """"
                End If
                If TokenNum > 0 AndAlso Tokens(TokenNum - 1) = "DISPLAY" Then
                    If Tokens(TokenNum).StartsWith(""""c) OrElse Tokens(TokenNum).StartsWith("GET") Then
                        Tokens(TokenNum - 1) = "DISPLAYSTRING"
                    End If
                End If
                ' --- replace some VB reserved words ---
                Select Case Tokens(TokenNum)
                    Case "REM"
                        Tokens(TokenNum) = "REMVAL"
                    Case "DATE"
                        Tokens(TokenNum) = "DATEVAL"
                    Case "SPACE"
                        Tokens(TokenNum) = "DISPLAYSPACE"
                    Case "CLEAR"
                        Tokens(TokenNum) = "CLEARSCREEN"
                    Case "READ"
                        Tokens(TokenNum) = "READFILE"
                    Case "WRITE"
                        Tokens(TokenNum) = "WRITEFILE"
                    Case "CLOSE"
                        Tokens(TokenNum) = "CLOSEFILE"
                    Case "RETURN"
                        Tokens(TokenNum) = "RETURNPROG"
                    Case "TAB"
                        Tokens(TokenNum) = "TABCURSOR"
                    Case "TRUE"
                        Tokens(TokenNum) = "TRUEVAL"
                    Case "FALSE"
                        Tokens(TokenNum) = "FALSEVAL"
                    Case "NULL"
                        Tokens(TokenNum) = """"""
                    Case "#"
                        Tokens(TokenNum) = "<>"
                    Case "ENDIF"
                        Tokens(TokenNum) = "END IF"
                End Select
            Next

            For TokenNum = 0 To LastToken

                ' --- check for line number ---

                If TokenNum = 0 AndAlso IsNumeric(Tokens(TokenNum)) Then
                    AddGotoLine(CInt(Tokens(TokenNum)))
                    Continue For
                End If

                ' --- check for "IF", "THEN", "ELSE" ---

                If Tokens(TokenNum) = "IF" Then
                    Result.Append("IF ")
                    TokenNum += 1
NextIf:
                    Result.Append(GetCompare(Tokens, TokenNum))
                    Result.Append(" ")
                    If TokenNum <= LastToken Then
                        If Tokens(TokenNum) = "AND" OrElse Tokens(TokenNum) = "OR" Then
                            Result.Append(Tokens(TokenNum))
                            Result.Append(" ")
                            TokenNum += 1
                            GoTo NextIf
                        End If
                        If Tokens(TokenNum) <> "THEN" Then
                            Result.Append("THEN ")
                        End If
                    Else
                        Result.Append("THEN ")
                        GoTo DoneCmd
                    End If
                    GoTo CheckCommand
                End If

                If Tokens(TokenNum) = "THEN" Then
                    TokenNum += 1 ' skip it
                    GoTo CheckCommand
                End If

                If Tokens(TokenNum) = "ELSE" Then
                    If TokenNum < LastToken Then
                        Result.Append("ELSE")
                        Result.Append(vbCrLf)
                        Result.Append("       ")
                        TokenNum += 1
                        NeedEndif = True
                    End If
                    GoTo CheckCommand
                End If

CheckCommand:

                ' --- DCH check must come before Assignment check, as the char may be "=" ---

                If Tokens(TokenNum) = "DCH" Then
                    TokenNum += 1
                    If Tokens(TokenNum) = "(" Then ' dch (hex)
                        TokenNum += 1
                        If Tokens(TokenNum).Length <> 2 Then GoTo ErrorFound
                        If Tokens(TokenNum + 1) <> ")" Then GoTo ErrorFound
                        Result.Append("DCH_HEX """)
                        Result.Append(Tokens(TokenNum))
                        Result.Append(""", ")
                        TokenNum += 2
                    Else
                        If Tokens(TokenNum).Length <> 1 Then GoTo ErrorFound
                        If Tokens(TokenNum) = """" Then
                            Result.Append("DCH """""""", ")
                        Else
                            Result.Append("DCH """)
                            Result.Append(Tokens(TokenNum))
                            Result.Append(""", ")
                        End If
                        TokenNum += 1
                    End If
                    If TokenNum <= LastToken Then
                        Result.Append(GetExpression(Tokens, TokenNum))
                    Else
                        Result.Append("1")
                    End If
                    GoTo DoneCmd
                End If

                ' --- check for buffer assignment ---

                If Len(Tokens(TokenNum)) = 1 AndAlso InStr("RZXYWSTUV", Tokens(TokenNum)) > 0 Then

                    If TokenNum >= LastToken Then GoTo Assignment

                    If Tokens(TokenNum + 1) = "=" Then GoTo Assignment ' will change to ?A below

                    TempResult = Tokens(TokenNum) ' buffer name
                    TempResult2 = "" ' holds offset
                    TempResult3 = "" ' holds numeric byte value
                    TokenNum += 1

                    If Tokens(TokenNum) = "[" Then
                        TokenNum += 1
                        TempResult2 = "_OFS " + GetExpression(Tokens, TokenNum) + ","
                        If Tokens(TokenNum) <> "]" Then GoTo ErrorFound
                        TokenNum += 1
                    End If

                    If Tokens(TokenNum) = "(" Then
                        If Tokens(TokenNum + 2) <> ")" Then GoTo ErrorFound
                        ' --- check for alpha ---
                        If Tokens(TokenNum + 1) = "A" Then
                            TempResult = TempResult + "_A" ' ?_A
                            TokenNum = TokenNum + 3
                        Else ' --- numeric ---
                            TempResult3 = Tokens(TokenNum + 1) + ", "
                            TokenNum = TokenNum + 3
                        End If
                    Else
                        TempResult = TempResult + "_A" ' default to alpha
                    End If

                    If Tokens(TokenNum) <> "=" Then GoTo ErrorFound
                    TokenNum += 1

                    Result.Append("LET_")
                    Result.Append(TempResult)
                    Result.Append(TempResult2)
                    Result.Append(" ")
                    Result.Append(TempResult3)
                    Result.Append(GetExpression(Tokens, TokenNum))
                    Result.Append(" ")
                    GoTo DoneCmd

                End If

                ' --- check for alpha assignment with offset ---

                If TokenNum < LastToken Then
                    If Tokens(TokenNum + 1) = "[" Then
                        TempResult = Tokens(TokenNum)
                        TokenNum += 2
                        TempResult2 = GetExpression(Tokens, TokenNum)
                        If Tokens(TokenNum) <> "]" Then GoTo ErrorFound
                        Tokens(TokenNum) = TempResult + "_OFS " + TempResult2 + ","
                        GoTo Assignment
                    End If
                End If

                ' --- check for simple assignment ---

Assignment:

                If TokenNum < LastToken Then
                    If Tokens(TokenNum + 1) = "=" Then
                        Select Case Tokens(TokenNum)
                            ' --- buffer without qualifier is assumed to be alpha ---
                            Case "R", "Z", "X", "Y", "W", "S", "T", "U", "V"
                                Tokens(TokenNum) += "_A"
                            Case "ESC"
                                Tokens(TokenNum) = "ESCVAL"
                            Case "CAN"
                                Tokens(TokenNum) = "CANVAL"
                            Case "LOCK"
                                Tokens(TokenNum) = "LOCKVAL"
                        End Select
                        If Tokens(TokenNum) = "N" OrElse
                           (Tokens(TokenNum).StartsWith("N") AndAlso NumOnly(Tokens(TokenNum).Substring(1))) OrElse
                           Tokens(TokenNum) = "REC" OrElse
                           Tokens(TokenNum) = "REMVAL" Then
                            Result.Append(Tokens(TokenNum))
                            Result.Append(" = ")
                            TokenNum += 2
                        Else
                            Result.Append("LET_")
                            Result.Append(Tokens(TokenNum))
                            Result.Append(" ")
                            TokenNum += 2
                        End If
                        ' --- check for expression ---
                        Try
                            Result.Append(GetExpression(Tokens, TokenNum))
                        Catch ex As Exception
                            ErrorMsg = ex.Message
                            GoTo ErrorFound
                        End Try
                        Result.Append(" ")
                        GoTo DoneCmd
                    End If
                End If

                If Tokens(TokenNum) = "SYSVAR" Then
                    TokenNum += 1
                    If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    If Tokens(TokenNum) <> "=" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("MEM(")
                    Result.Append(TempResult)
                    Result.Append(") = ")
                    Result.Append(TempResult2)
                    GoTo DoneCmd
                End If

                ' --- check for transfer of control commands ---

                If Tokens(TokenNum) = "GOTO" Then
                    AddGotoLine(CInt(Tokens(TokenNum + 1)))
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "LOAD" Then
                    Result.Append("LOADPROG ")
                    TokenNum += 1
                    Result.Append(GetExpression(Tokens, TokenNum))
                    NeedExitSub = True
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "GOSUB" Then
                    Result.Append("GOSUBPROG ")
                    TokenNum += 1
                    Result.Append(GetExpression(Tokens, TokenNum)) ' gosub program number
                    Result.Append(", ")
                    Result.Append(Lines2.Count + 1) ' return line number
                    NeedExitSub = True
                    NeedNextLineNum = True
                    AddJumpPoint(Lines2.Count + 1) ' add to jumppointlist
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "GOS" Then
                    Result.Append("GOS ")
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum) ' gos line number
                    Result.Append(TempResult)
                    Result.Append(", ")
                    Result.Append(Lines2.Count + 1)
                    NeedExitSub = True
                    NeedNextLineNum = True
                    AddJumpPoint(CInt(TempResult))
                    AddJumpPoint(Lines2.Count + 1)
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "RELEASETERMINAL" Then
                    Result.Append(Tokens(TokenNum))
                    Result.Append(" ")
                    TokenNum += 1
                    TempResult = (Lines2.Count + 1).ToString ' spawn line number
                    TempResult2 = GetExpression(Tokens, TokenNum) ' error line number
                    Result.Append(TempResult)
                    Result.Append(", ")
                    Result.Append(TempResult2)
                    NeedExitSub = True
                    NeedNextLineNum = True
                    AddJumpPoint(CInt(TempResult))
                    AddJumpPoint(CInt(TempResult2))
                    GoTo DoneCmd
                End If

                ' --- convert buffer commands to subroutines ---

                If Tokens(TokenNum) = "INIT" OrElse
                   Tokens(TokenNum) = "SET" OrElse
                   Tokens(TokenNum) = "RESET" Then
                    If TokenNum = LastToken Then GoTo DoneCmd
                    If Right$(Tokens(TokenNum + 1), 1) = "P" Then
                        Tokens(TokenNum + 1) = Left$(Tokens(TokenNum + 1), Len(Tokens(TokenNum + 1)) - 1)
                    End If
                    Result.Append(Tokens(TokenNum))
                    Result.Append("_")
                    Result.Append(Tokens(TokenNum + 1))
                    Result.Append(" ")
                    TokenNum += 2
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "SKIP" Then
                    TokenNum += 1
                    TempResult = Tokens(TokenNum)
                    TokenNum += 2
                    If Tokens(TokenNum) = "A" Then
                        TempResult += "_A"
                        TokenNum += 2
                        If TokenNum <= LastToken Then
                            TempResult2 = GetExpression(Tokens, TokenNum)
                            Result.Append("SKIP_")
                            Result.Append(TempResult)
                            Result.Append(" ")
                            Result.Append(TempResult2)
                            Result.Append(" ")
                        Else
                            Result.Append("SKIP_")
                            Result.Append(TempResult)
                            Result.Append(" 1 ")
                        End If
                    Else
                        TempResult2 = GetExpression(Tokens, TokenNum)
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        Result.Append("SKIP_")
                        Result.Append(TempResult)
                        Result.Append(" ")
                        Result.Append(TempResult2)
                        Result.Append(" ")
                    End If
                    GoTo DoneCmd
                End If

                ' --- check for commands that have their own syntax ---

                If Tokens(TokenNum) = "ENTERNUM" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    Try
                        TempResult = GetNumericFormat(Tokens, TokenNum)
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        TempResult2 = GetTarget(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        If Not IsNumericItem(TempResult2) Then GoTo ErrorFound
                        ' --- see if extra internal overflow checking needed ---
                        If TempResult <> "0" AndAlso TempResult <> "1" AndAlso TempResult <> "2" AndAlso
                           IsNumericByteItem(TempResult2) Then
                            Result.Append("IF NOT ENTERBYTE(""")
                        Else
                            Result.Append("IF NOT ENTERNUM(""")
                        End If
                        Result.Append(TempResult)
                        Result.Append(""") THEN EXIT SUB")
                        Result.Append(vbCrLf)
                        Result.Append("       ")
                        If TempResult2 = "N" OrElse
                           (Left$(TempResult2, 1) = "N" AndAlso InStr("123456789_", Mid$(TempResult2, 2, 1)) > 0) OrElse
                           TempResult2 = "REC" OrElse
                           TempResult2 = "REMVAL" Then
                            Result.Append(TempResult2)
                            Result.Append(" = NUMERIC_RESULT ")
                        Else
                            Result.Append("LET_")
                            Result.Append(TempResult2)
                            Result.Append(" NUMERIC_RESULT ")
                        End If
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "ENTERALPHA" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    Try
                        TempResult = GetExpression(Tokens, TokenNum)
                        If InStr(TempResult, ".") > 0 Then GoTo ErrorFound
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        TempResult2 = GetTarget(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        Result.Append("IF NOT ENTERALPHA(")
                        Result.Append(TempResult)
                        Result.Append(") THEN EXIT SUB")
                        Result.Append(vbCrLf)
                        Result.Append("       ")
                        Result.Append("LET_")
                        Result.Append(TempResult2)
                        Result.Append(" ALPHA_RESULT ")
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "EDITNUM" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    ' --- check for numeric enter first ---
                    Try
                        TempResult = GetNumericFormat(Tokens, TokenNum)
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        SaveTokenNum2 = TokenNum
                        TempResult2 = GetTarget(Tokens, TokenNum)
                        TokenNum = SaveTokenNum2
                        TempResult4 = GetItem(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        If Not IsNumericItem(TempResult2) Then GoTo ErrorFound
                        ' --- see if extra internal overflow checking needed ---
                        If TempResult <> "0" AndAlso TempResult <> "1" AndAlso TempResult <> "2" AndAlso
                           IsNumericByteItem(TempResult2) Then
                            Result.Append("IF NOT EDITBYTE(""")
                        Else
                            Result.Append("IF NOT EDITNUM(""")
                        End If
                        Result.Append(TempResult)
                        Result.Append(""", ")
                        Result.Append(TempResult4)
                        Result.Append(") THEN EXIT SUB")
                        Result.Append(vbCrLf)
                        Result.Append("       ")
                        If TempResult2 = "N" OrElse
                           (Left$(TempResult2, 1) = "N" AndAlso InStr("123456789_", Mid$(TempResult2, 2, 1)) > 0) OrElse
                           TempResult2 = "REC" OrElse
                           TempResult2 = "REMVAL" Then
                            Result.Append(TempResult2)
                            Result.Append(" = NUMERIC_RESULT ")
                        Else
                            Result.Append("LET_")
                            Result.Append(TempResult2)
                            Result.Append(" NUMERIC_RESULT ")
                        End If
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "EDITALPHA" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    Try
                        TempResult = GetExpression(Tokens, TokenNum)
                        If InStr(TempResult, ".") > 0 Then GoTo ErrorFound
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        SaveTokenNum2 = TokenNum
                        TempResult2 = GetTarget(Tokens, TokenNum)
                        TokenNum = SaveTokenNum2
                        TempResult4 = GetItem(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        Result.Append("IF NOT EDITALPHA(")
                        Result.Append(TempResult)
                        Result.Append(", ")
                        Result.Append(TempResult4)
                        Result.Append(") THEN EXIT SUB")
                        Result.Append(vbCrLf)
                        Result.Append("       ")
                        Result.Append("LET_")
                        Result.Append(TempResult2)
                        Result.Append(" ALPHA_RESULT ")
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "DISPLAY" Then
                    If Tokens(TokenNum + 1) <> "(" Then
                        TokenNum += 1
                        TempResult = GetExpression(Tokens, TokenNum)
                        TempResult2 = "DISPLAY"
                        If Left$(TempResult, 1) = """" OrElse Left$(TempResult, 3) = "GET" Then
                            TempResult2 = "DISPLAYSTRING"
                        End If
                        Result.Append(TempResult2)
                        Result.Append(" ")
                        Result.Append(TempResult)
                        Result.Append(" ")
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        GoTo DoneCmd
                    End If
                    TokenNum += 2
                    TempResult = GetNumericFormat(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("DISPLAYNUM """)
                    Result.Append(TempResult)
                    Result.Append(""", ")
                    Result.Append(TempResult2)
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "CURSORAT" Then
                    ' --- have to parse to handle expressions properly ---
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "," Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("CURSORAT ")
                    Result.Append(TempResult)
                    Result.Append(", ")
                    Result.Append(TempResult2)
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "SPOOL" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    ' --- check if numeric spool first ---
                    Try
                        SaveTokenNum = TokenNum
                        TempResult = GetNumericFormat(Tokens, TokenNum)
                        If Tokens(TokenNum) <> ")" Then GoTo AlphaSpool
                        TokenNum += 1
                        TempResult2 = GetExpression(Tokens, TokenNum)
                        If Not IsNumericExpr(TempResult2) Then GoTo AlphaSpool
                        If Tokens(TokenNum) <> "TO" Then GoTo AlphaSpool
                        TokenNum += 1
                        TempResult3 = GetTarget(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo AlphaSpool
                        Result.Append("SPOOL_")
                        Result.Append(TempResult3)
                        Result.Append(" ")
                        Result.Append(FormatLen(TempResult).ToString)
                        Result.Append(", ")
                        Result.Append("FORMATNUM(""")
                        Result.Append(TempResult)
                        Result.Append(""", ")
                        Result.Append(TempResult2)
                        Result.Append(") ")
                        GoTo DoneCmd
                    Catch ex As Exception
                        GoTo AlphaSpool
                    End Try
AlphaSpool:
                    Try
                        TokenNum = SaveTokenNum
                        TempResult = GetExpression(Tokens, TokenNum)
                        If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                        TokenNum += 1
                        TempResult2 = GetExpression(Tokens, TokenNum)
                        If Tokens(TokenNum) <> "TO" Then GoTo ErrorFound
                        TokenNum += 1
                        TempResult3 = GetTarget(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        Result.Append("SPOOL_")
                        Result.Append(TempResult3)
                        Result.Append(" ")
                        Result.Append(TempResult)
                        Result.Append(", ")
                        Result.Append(TempResult2)
                        Result.Append(" ")
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "PACK" Then
                    TokenNum += 1
                    If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetMemPos(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "TO" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult3 = GetTarget(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("LET_")
                    Result.Append(TempResult3)
                    Result.Append(" PACK(")
                    Result.Append(TempResult)
                    Result.Append(", ")
                    Result.Append(TempResult2)
                    Result.Append(") ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "CONVERT" Then
                    If Tokens(TokenNum + 1) <> "(" Then GoTo ErrorFound
                    TokenNum += 2
                    TempResult = GetNumericFormat(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetMemPos(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "TO" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult3 = GetTarget(Tokens, TokenNum) ' target numeric
                    TempResult4 = Tokens(TokenNum) ' error line
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("IF NOT CONVERT(""")
                    Result.Append(TempResult)
                    Result.Append(""", ")
                    Result.Append(TempResult2)
                    Result.Append(") THEN GOTO ")
                    Result.Append(TempResult4)
                    AddGotoLine(CInt(TempResult4))
                    Result.Append(vbCrLf)
                    Result.Append("       ")
                    If TempResult3 = "N" OrElse
                       (Left$(TempResult3, 1) = "N" AndAlso InStr("123456789", Mid$(TempResult3, 2, 1)) > 0) OrElse
                       TempResult3 = "REC" OrElse
                       TempResult3 = "REMVAL" Then
                        Result.Append(TempResult3)
                        Result.Append(" = NUMERIC_RESULT ")
                    Else
                        Result.Append("LET_")
                        Result.Append(TempResult3)
                        Result.Append(" NUMERIC_RESULT ")
                    End If
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "WHENCANCEL" OrElse
                   Tokens(TokenNum) = "WHENESCAPE" OrElse
                   Tokens(TokenNum) = "WHENERROR" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    If Tokens(TokenNum) = "GOTO" Then
                        TokenNum += 1
                        TempResult2 = "PROG"
                        TempResult3 = Tokens(TokenNum) ' line number
                        AddJumpPoint(CInt(TempResult3))
                    ElseIf Tokens(TokenNum) = "LOAD" Then
                        TokenNum += 1
                        TempResult2 = Tokens(TokenNum) ' program number
                        TempResult3 = "0" ' line number
                    ElseIf Tokens(TokenNum) = "TRAP" Then
                        TempResult2 = "-1" ' program number
                        TempResult3 = "-1" ' line number
                    Else
                        GoTo ErrorFound
                    End If
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append(TempResult)
                    Result.Append(" ")
                    Result.Append(TempResult2)
                    Result.Append(", ")
                    Result.Append(TempResult3)
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "SORT" Then
                    TokenNum += 1
                    If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    ' --- check for numeric sort ---
                    Try
                        SaveTokenNum = TokenNum
                        TempResult2 = GetExpression(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo SortAlpha
                        If Not IsNumericExpr(TempResult2) Then GoTo SortAlpha
                        Result.Append("SORTNUM ")
                        Result.Append(TempResult)
                        Result.Append(", ")
                        Result.Append(TempResult2)
                        GoTo DoneCmd
                    Catch ex As Exception
                        GoTo SortAlpha
                    End Try
SortAlpha:
                    Try
                        TokenNum = SaveTokenNum
                        TempResult2 = GetItem(Tokens, TokenNum)
                        If TokenNum <= LastToken Then GoTo ErrorFound
                        Result.Append("SORTALPHA ")
                        Result.Append(TempResult)
                        Result.Append(", ")
                        Result.Append(TempResult2)
                        GoTo DoneCmd
                    Catch ex As Exception
                        ErrorMsg = ex.Message
                        GoTo ErrorFound
                    End Try
                End If

                If Tokens(TokenNum) = "MOVE" Then
                    TokenNum += 1
                    If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetMemPos(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "TO" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult3 = GetMemPos(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("MOVE ")
                    Result.Append(TempResult)
                    Result.Append(", ")
                    Result.Append(TempResult2)
                    Result.Append(", ")
                    Result.Append(TempResult3)
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "UPDATE" Then
                    TokenNum += 1
                    If Tokens(TokenNum) <> "(" Then
                        TempResult = Tokens(TokenNum) ' buffer name
                        TokenNum += 1
                    Else
                        TempResult = "R" ' default buffer
                    End If
                    If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum) ' number of bytes
                    If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult3 = GetExpression(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("SET_")
                    Result.Append(TempResult)
                    Result.Append(" : UPDATE_VALUE = ")
                    Result.Append(TempResult3)
                    Result.Append(" : RESET_")
                    Result.Append(TempResult)
                    Result.Append(" : LET_")
                    Result.Append(TempResult)
                    Result.Append(" " + TempResult2)
                    Result.Append(", UPDATE_VALUE ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "FLIP" OrElse Tokens(TokenNum) = "FLOP" OrElse
                   Tokens(TokenNum) = "NOSIGN" OrElse Tokens(TokenNum) = "SIGNSET" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "TO" Then GoTo ErrorFound
                    TokenNum += 1
                    TempResult3 = GetTarget(Tokens, TokenNum)
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    If TempResult3 = "N" OrElse
                       (Left$(TempResult3, 1) = "N" AndAlso InStr("123456789", Mid$(TempResult3, 2, 1)) > 0) OrElse
                       TempResult3 = "REC" OrElse
                       TempResult3 = "REMVAL" Then
                        Result.Append(TempResult3)
                        Result.Append(" = ")
                    Else
                        Result.Append("LET_")
                        Result.Append(TempResult3)
                        Result.Append(" ")
                    End If
                    Result.Append(TempResult)
                    Result.Append("( ")
                    Result.Append(TempResult2)
                    Result.Append(" ) ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "ON" Then
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) <> "GOTO" Then GoTo ErrorFound
                    TokenNum += 1
                    Result.Append("ON TO_BYTE(")
                    Result.Append(TempResult)
                    Result.Append(" + 1) GOTO ")
                    Do While TokenNum <= LastToken
                        If Tokens(TokenNum) = "," Then
                            Result.Append(", ")
                        Else
                            Result.Append(Tokens(TokenNum))
                            AddGotoLine(CInt(Tokens(TokenNum)))
                        End If
                        TokenNum += 1
                    Loop
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                ' --- check for "command expression", with expression required ---

                If Tokens(TokenNum) = "LEFT" Then Tokens(TokenNum) = "MOVELEFT"
                If Tokens(TokenNum) = "RIGHT" Then Tokens(TokenNum) = "MOVERIGHT"
                If Tokens(TokenNum) = "UP" Then Tokens(TokenNum) = "MOVEUP"
                If Tokens(TokenNum) = "DOWN" Then Tokens(TokenNum) = "MOVEDOWN"

                If Tokens(TokenNum) = "ATT" OrElse
                   Tokens(TokenNum) = "BACKSPACE" OrElse
                   Tokens(TokenNum) = "BELL" OrElse
                   Tokens(TokenNum) = "CHARDELETE" OrElse
                   Tokens(TokenNum) = "CHARINSERT" OrElse
                   Tokens(TokenNum) = "CLOSEFILE" OrElse
                   Tokens(TokenNum) = "CLOSECHANNEL" OrElse
                   Tokens(TokenNum) = "CLOSEDEVICE" OrElse
                   Tokens(TokenNum) = "CRD" OrElse
                   Tokens(TokenNum) = "DELAY" OrElse
                   Tokens(TokenNum) = "DISPLAYSPACE" OrElse
                   Tokens(TokenNum) = "FF" OrElse
                   Tokens(TokenNum) = "INITSORT" OrElse
                   Tokens(TokenNum) = "LINEDELETE" OrElse
                   Tokens(TokenNum) = "LINEINSERT" OrElse
                   Tokens(TokenNum) = "MOVEDOWN" OrElse
                   Tokens(TokenNum) = "MOVELEFT" OrElse
                   Tokens(TokenNum) = "MOVERIGHT" OrElse
                   Tokens(TokenNum) = "MOVEUP" OrElse
                   Tokens(TokenNum) = "NL" OrElse
                   Tokens(TokenNum) = "PAD" OrElse
                   Tokens(TokenNum) = "SCROLLDOWN" OrElse
                   Tokens(TokenNum) = "SCROLLUP" OrElse
                   Tokens(TokenNum) = "TABCURSOR" Then
                    Result.Append(Tokens(TokenNum))
                    Result.Append(" ")
                    TokenNum += 1
                    If TokenNum > LastToken Then
                        Result.Append("1")
                    Else
                        Result.Append(GetExpression(Tokens, TokenNum))
                    End If
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    GoTo DoneCmd
                End If

                ' --- check for "command errorline" ---

                If Tokens(TokenNum) = "CREATETFA" OrElse
                   Tokens(TokenNum) = "OPENTFA" OrElse
                   Tokens(TokenNum) = "OPENVOLUME" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = Tokens(TokenNum)
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("IF NOT ")
                    Result.Append(TempResult)
                    Result.Append(" THEN GOTO ")
                    Result.Append(TempResult2)
                    AddGotoLine(CInt(TempResult2))
                    GoTo DoneCmd
                End If

                ' --- check for "command expression errorline" ---

                If Tokens(TokenNum) = "ASSIGNDEVICE" OrElse Tokens(TokenNum) = "CONTROLCHANNEL" OrElse
                   Tokens(TokenNum) = "CREATECHANNEL" OrElse Tokens(TokenNum) = "DELETE" OrElse
                   Tokens(TokenNum) = "DELETECHANNEL" OrElse Tokens(TokenNum) = "EOFCHANNEL" OrElse
                   Tokens(TokenNum) = "FETCH" OrElse Tokens(TokenNum) = "OPENCHANNEL" OrElse
                   Tokens(TokenNum) = "OPENDATA" OrElse Tokens(TokenNum) = "OPENDEVICE" OrElse
                   Tokens(TokenNum) = "OPENDIRDEVICE" OrElse Tokens(TokenNum) = "OPENDIRECTORY" OrElse
                   Tokens(TokenNum) = "OPENDIRLIB" OrElse Tokens(TokenNum) = "OPENDIRTFA" OrElse
                   Tokens(TokenNum) = "OPENDIRVOLUME" OrElse Tokens(TokenNum) = "OPENSORTFILE" OrElse
                   Tokens(TokenNum) = "READFILE" OrElse Tokens(TokenNum) = "READKEY" OrElse
                   Tokens(TokenNum) = "READREC" OrElse Tokens(TokenNum) = "RENAMECHANNEL" OrElse
                   Tokens(TokenNum) = "REWINDCHANNEL" OrElse Tokens(TokenNum) = "WINDCHANNEL" OrElse
                   Tokens(TokenNum) = "WRITEFILE" OrElse Tokens(TokenNum) = "WRITEKEY" OrElse
                   Tokens(TokenNum) = "WRITEREC" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    TempResult3 = Tokens(TokenNum)
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("IF NOT ")
                    Result.Append(TempResult)
                    Result.Append("(")
                    Result.Append(TempResult2)
                    Result.Append(") THEN GOTO ")
                    Result.Append(TempResult3)
                    AddGotoLine(CInt(TempResult3))
                    GoTo DoneCmd
                End If

                ' --- check for "token/expr/until/label" ---

                If Tokens(TokenNum) = "BACKSPACECHANNEL" Then
                    Result.Append("IF NOT ")
                    Result.Append(Tokens(TokenNum)) ' no trailing space
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) = "UNTIL" Then
                        TokenNum += 1
                        TempResult = TempResult + ", MEMPOS_" + Tokens(TokenNum) ' until buffer
                        TokenNum += 1
                    Else
                        TempResult = TempResult + ", -1" ' default to standard list
                    End If
                    TempResult2 = Tokens(TokenNum) ' line number
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("( ")
                    Result.Append(TempResult)
                    Result.Append(" ) THEN GOTO ")
                    Result.Append(TempResult2)
                    AddGotoLine(CInt(TempResult2))
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                ' --- check for token/expr/to/until/label ---

                If Tokens(TokenNum) = "READCHANNEL" Then
                    Result.Append("IF NOT ")
                    Result.Append(Tokens(TokenNum)) ' no trailing space
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) = "TO" Then
                        TokenNum += 1
                        TempResult = TempResult + ", MEMPOS_" + Tokens(TokenNum) ' to buffer
                        TokenNum += 1
                    Else
                        TempResult = TempResult + ", MEMPOS_R" ' default to buffer
                    End If
                    If Tokens(TokenNum) = "UNTIL" Then
                        TokenNum += 1
                        TempResult = TempResult + ", MEMPOS_" + Tokens(TokenNum) ' until buffer
                        TokenNum += 1
                    Else
                        TempResult = TempResult + ", -1" ' default to standard list
                    End If
                    TempResult2 = Tokens(TokenNum) ' line number
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("( ")
                    Result.Append(TempResult)
                    Result.Append(" ) THEN GOTO ")
                    Result.Append(TempResult2)
                    AddGotoLine(CInt(TempResult2))
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                ' --- check for token/expr/from/label ---

                If Tokens(TokenNum) = "WRITECHANNEL" Then
                    Result.Append("IF NOT ")
                    Result.Append(Tokens(TokenNum)) ' no trailing space
                    TokenNum += 1
                    TempResult = GetExpression(Tokens, TokenNum)
                    If Tokens(TokenNum) = "FROM" Then
                        TokenNum += 1
                        TempResult = TempResult + ", MEMPOS_" + Tokens(TokenNum) ' from buffer
                        TokenNum += 1
                    Else
                        TempResult = TempResult + ", MEMPOS_W" ' default from buffer
                    End If
                    TempResult2 = Tokens(TokenNum) ' line number
                    TokenNum += 1
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append("( ")
                    Result.Append(TempResult)
                    Result.Append(" ) THEN GOTO ")
                    Result.Append(TempResult2)
                    AddGotoLine(CInt(TempResult2))
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                ' --- check for invalid but known commands ---

                If Tokens(TokenNum) = "HEX" Then
                    Result.Append("THROWERROR ""SYNTAX ERROR"", ""HEX")
                    TokenNum += 1
                    Do While TokenNum <= LastToken
                        Result.Append(" ")
                        Result.Append(Tokens(TokenNum))
                        TokenNum += 1
                    Loop
                    Result.Append("""")
                    GoTo DoneCmd
                End If

                ' --- special IDRIS commands ---

                If Tokens(TokenNum) = "DEBUG" Then
                    TokenNum += 1
                    If TokenNum > LastToken Then GoTo ErrorFound
                    Select Case Tokens(TokenNum)
                        Case "PRINT"
                            If TokenNum = LastToken Then GoTo ErrorFound
                            Result.Append("Debug.Print """)
                            Result.Append(Tokens(TokenNum + 1))
                            Result.Append(""" ")
                            TokenNum = LastToken + 1
                            GoTo DoneCmd
                        Case "BREAK"
                            Result.Append("Debug.Assert False ")
                            TokenNum = LastToken + 1
                            GoTo DoneCmd
                        Case "CMD"
                            If TokenNum = LastToken Then GoTo ErrorFound
                            TempResult = Tokens(TokenNum + 1)
                            TempResult = Replace(TempResult, vbTab, " ")
                            TempResult = Replace(TempResult, """", """""")
                            Result.Append("SendToServer ""DEBUG"" & vbTab & ""CMD"" & vbTab & """)
                            Result.Append(TempResult)
                            Result.Append(""" ")
                            TokenNum = LastToken + 1
                            GoTo DoneCmd
                        Case "COMMENT"
                            If TokenNum = LastToken Then GoTo ErrorFound
                            TempResult = Tokens(TokenNum + 1)
                            TempResult = Replace(TempResult, vbTab, " ")
                            Result.Append("' --- ")
                            Result.Append(TempResult)
                            Result.Append(" --- ")
                            TokenNum = LastToken + 1
                            GoTo DoneCmd
                        Case Else
                            GoTo ErrorFound
                    End Select
                End If

                If Tokens(TokenNum) = "SETSUBQUERY" OrElse Tokens(TokenNum) = "SETSUBQUERYFILE" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetItem(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append(TempResult)
                    Result.Append(" ")
                    Result.Append(TempResult2)
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "SAVEFILEINFO" OrElse Tokens(TokenNum) = "RESTOREFILEINFO" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetExpression(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append(TempResult)
                    Result.Append(" ")
                    Result.Append(TempResult2)
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "EXECSQL" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetItem(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append(TempResult)
                    Result.Append(" ")
                    Result.Append(TempResult2)
                    Result.Append(" ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "FATALERROR" Then
                    TempResult = Tokens(TokenNum)
                    TokenNum += 1
                    TempResult2 = GetItem(Tokens, TokenNum)
                    If TokenNum <= LastToken Then GoTo ErrorFound
                    Result.Append(TempResult)
                    Result.Append(" ")
                    Result.Append(TempResult2)
                    Result.Append(" : EXIT SUB ")
                    GoTo DoneCmd
                End If

                If Tokens(TokenNum) = "EXITRUNTIME" Then
                    If TokenNum < LastToken Then GoTo ErrorFound
                    Result.Append(Tokens(TokenNum))
                    Result.Append(" : EXIT SUB ")
                    TokenNum += 1
                    GoTo DoneCmd
                End If

                If TokenNum > 0 AndAlso Tokens(TokenNum) = "=" Then
                    Select Case Tokens(TokenNum - 1)
                        Case "ESC"
                            Tokens(TokenNum - 1) = "ESCVAL"
                        Case "CAN"
                            Tokens(TokenNum - 1) = "CANVAL"
                        Case "LOCK"
                            Tokens(TokenNum - 1) = "LOCKVAL"
                    End Select
                    ' --- collapse unary minus after equals sign ---
                    If TokenNum < LastToken - 1 Then
                        If Tokens(TokenNum + 1) = "-" Then
                            Tokens(TokenNum + 2) = "-" + Tokens(TokenNum + 2)
                            Tokens(TokenNum + 1) = ""
                        End If
                    End If
                End If
                If TokenNum = LastToken Then
                    Select Case Tokens(TokenNum)
                        Case "LOCK"
                            Tokens(TokenNum) = "LOCKREC"
                        Case "UNLOCK"
                            Tokens(TokenNum) = "UNLOCKREC"
                        Case "ESCAPE"
                            Tokens(TokenNum) = "ESC"
                        Case "CAN"
                            Tokens(TokenNum) = "CANCEL"
                    End Select
                    If Tokens(TokenNum) = "ESC" OrElse
                       Tokens(TokenNum) = "CANCEL" OrElse
                       Tokens(TokenNum) = "WRITEBACK" OrElse
                       Tokens(TokenNum) = "RELEASEDEVICE" Then
                        Tokens(TokenNum) += " : IF MUSTEXIT THEN EXIT SUB "
                        GoTo DoneCmd
                    End If
                    If Tokens(TokenNum) = "RETURNPROG" Then
                        NeedExitSub = True
                        GoTo DoneCmd
                    End If
                End If
            Next

DoneCmd:

            If Result.Length > 0 Then GoTo BuildResult
            For TokenNum = 0 To LastToken
                If TokenNum = 0 Then
                    If IsNumeric(Tokens(TokenNum)) Then
                        Result.Append(Tokens(TokenNum))
                        Result.Append(":"c)
                        Result.Append(Space(5 - Tokens(TokenNum).Length))
                        Continue For
                    End If
                End If
                If Tokens(TokenNum) = "" Then Continue For
                If Result.Length > 0 Then
                    Result.Append(" "c)
                End If
                Result.Append(Tokens(TokenNum).TrimEnd)
            Next
            GoTo CheckForEndings
BuildResult:
            ' --- get rest of command, if any ---
            Do While TokenNum <= LastToken
                If Tokens(TokenNum) = ")" Then
                    Do While Result.Chars(Result.Length - 1) = " "c
                        Result.Remove(Result.Length - 1, 1)
                    Loop
                End If
                Result.Append(Tokens(TokenNum))
                Result.Append(" ")
                TokenNum += 1
            Loop
            GoTo CheckForEndings
CheckForEndings:
            If Result.Length > 0 Then
                If NeedEndif Then
                    Result.Append(vbCrLf)
                    Result.Append("       END IF ")
                End If
                If NeedExitSub Then
                    Result.Append(" : EXIT SUB ")
                End If
                Dim ResultPrefix As String = "       "
                If Result.Chars(0) >= "0"c AndAlso Result.Chars(0) <= "9"c Then
                    ResultPrefix = ""
                ElseIf NeedThisLineNum Then
                    ResultPrefix = (Lines2.Count.ToString + ":       ").Substring(0, 7)
                End If
                ' --- add line to list ---
                Lines2.Add(ResultPrefix + Result.ToString.TrimEnd)
                NeedThisLineNum = NeedNextLineNum
            End If
        Next
        Return True
ErrorFound:
        Dim ErrorInfo As New ParseError
        With ErrorInfo
            .LineNum = SourceLinenum
            .SourceLine = CurrLine.Replace(vbTab, " "c).Trim
            .ErrorDesc = "***ERROR*** - " + ErrorMsg
        End With
        ParseErrors.Add(ErrorInfo)
        ErrorInfo = Nothing
        Return False
    End Function

    Private Function GetMemPos(ByRef Tokens() As String, ByRef TokenNum As Integer) As String
        Dim strResult As String
        ' ---------------------
        strResult = GetItem(Tokens, TokenNum)
        If InStr("RZXYWSTUV", Left$(strResult, 1)) > 0 AndAlso Mid$(strResult, 2, 2) = "_A" Then
            strResult = Left$(strResult, 1) + Mid$(strResult, 4)
        End If
        If InStr(strResult, "_OFS(") > 0 Then
            strResult = Replace(strResult, "_OFS(", ", ")
            strResult = Left$(strResult, Len(strResult) - 1) ' remove ")"
        Else
            strResult = strResult + ", 0"
        End If
        strResult = "MEMPOS_" + strResult
        Return strResult
    End Function

    Private Function FormatLen(ByVal DisplayFmt As String) As Integer

        Dim intResult As Integer
        Dim TempChar As String
        Dim intLoop As Integer

        Dim FrontNeg As Boolean
        Dim FrontParen As Boolean
        Dim ZeroFill As Boolean
        Dim StarFill As Boolean
        Dim Comma As Boolean
        Dim DigitsFound As Boolean
        Dim DigitsAbove As Integer
        Dim DecimalPoint As Boolean
        Dim DigitsBelow As Integer
        Dim RearNeg As Boolean
        Dim RearParen As Boolean
        ' -------------------------

        ' --- clear flags ---
        FrontNeg = False
        FrontParen = False
        ZeroFill = False
        StarFill = False
        Comma = False
        DigitsFound = False
        DigitsAbove = 0
        DecimalPoint = False
        DigitsBelow = 0
        RearNeg = False
        RearParen = False

        ' --- parse display format into flags ---
        For intLoop = 1 To Len(DisplayFmt)
            TempChar = Mid$(DisplayFmt, intLoop, 1)
            Select Case TempChar
                Case "("
                    If DigitsFound OrElse FrontParen OrElse FrontNeg Then GoTo ErrorFound
                    FrontParen = True
                Case "-"
                    If Not DigitsFound Then
                        If FrontParen OrElse FrontNeg Then GoTo ErrorFound
                        FrontNeg = True
                    Else
                        If RearNeg OrElse RearParen Then GoTo ErrorFound
                        RearNeg = True
                    End If
                Case "*"
                    If DigitsFound OrElse StarFill OrElse ZeroFill Then GoTo ErrorFound
                    StarFill = True
                Case "z", "Z"
                    If DigitsFound OrElse StarFill OrElse ZeroFill Then GoTo ErrorFound
                    ZeroFill = True
                Case ","
                    If DigitsFound OrElse Comma Then GoTo ErrorFound
                    Comma = True
                Case "0" To "9"
                    If Not DecimalPoint Then
                        DigitsAbove = CInt((DigitsAbove * 10) + Val(TempChar))
                    Else
                        If RearNeg OrElse RearParen Then GoTo ErrorFound
                        DigitsBelow = CInt((DigitsBelow * 10) + Val(TempChar))
                    End If
                    DigitsFound = True
                Case "."
                    If Not DigitsFound Then GoTo ErrorFound
                    If DecimalPoint Then GoTo ErrorFound
                    DecimalPoint = True
                Case ")"
                    If Not DigitsFound Then GoTo ErrorFound
                    If RearNeg OrElse RearParen Then GoTo ErrorFound
                    RearParen = True
                Case Else
                    GoTo ErrorFound
            End Select
        Next intLoop

        ' --- check for digit overflows ---
        If DigitsAbove + DigitsBelow > 14 Then GoTo ErrorFound
        If DigitsBelow > 7 Then GoTo ErrorFound

        ' --- clean up decimal point ---
        If DigitsBelow = 0 Then DecimalPoint = False

        ' --- clean up negative display ---
        If FrontParen OrElse RearParen Then
            FrontParen = True
            RearParen = True
            FrontNeg = False
            RearNeg = False
        End If
        If RearNeg Then
            FrontNeg = False
        End If

        ' --- calculate length of formatted number ---
        intResult = DigitsAbove + DigitsBelow
        If DecimalPoint Then intResult = intResult + 1
        If FrontParen Then intResult = intResult + 1
        If RearParen Then intResult = intResult + 1
        If FrontNeg Then intResult = intResult + 1
        If RearNeg Then intResult = intResult + 1
        If Comma AndAlso DigitsAbove > 3 Then
            intResult = intResult + ((DigitsAbove - 1) \ 3)
        End If

        ' --- done ---
        Return intResult

        Exit Function

ErrorFound:

        Return -1 ' invalid

    End Function

    Private Function GetNumericFormat(ByRef Tokens() As String, ByRef TokenNum As Integer) As String
        Dim Result As New StringBuilder
        ' -----------------------------
        Do
            Result.Append(Tokens(TokenNum))
            TokenNum += 1
        Loop Until Tokens(TokenNum) = ")"
        If Tokens(TokenNum + 1) = ")" Then
            Result.Append(Tokens(TokenNum))
            TokenNum += 1
        End If
        Return Result.ToString
    End Function

    Private Function IsNumericItem(ByVal Value As String) As Boolean
        IsNumericItem = False
        If Value = "N" Then Return True
        If Value = "F" Then Return True
        If Value = "G" Then Return True
        If Value = "REC" Then Return True
        If Value = "REMVAL" Then Return True
        If InStr("NFG", Left$(Value, 1)) > 0 AndAlso InStr("123456789", Mid$(Value, 2, 1)) > 0 Then
            Return True
        End If
        If Left$(Value, 6) = "N_OFS(" Then Return True
        ' --- "W(1)" format ---
        If InStr("RZXYWSTUV", Left$(Value, 1)) > 0 AndAlso Mid$(Value, 2, 1) = "(" AndAlso Mid$(Value, 4, 1) = ")" Then
            If InStr("123456", Mid$(Value, 3, 1)) > 0 Then Return True
        End If
        ' --- "W 1," format ---
        If InStr("RZXYWSTUV", Left$(Value, 1)) > 0 AndAlso Mid$(Value, 2, 1) = " " AndAlso Mid$(Value, 4, 1) = "," Then
            If InStr("123456", Mid$(Value, 3, 1)) > 0 Then Return True
        End If
        Return IsSysVar(Value)
    End Function

    Private Function IsNumericByteItem(ByVal Value As String) As Boolean
        ' --- check for flag registers ---
        If Value = "F" Then Return True
        If Value = "G" Then Return False ' G is two bytes
        If InStr("FG", Left$(Value, 1)) > 0 AndAlso InStr("123456789", Mid$(Value, 2, 1)) > 0 Then
            Return True
        End If
        ' --- "W(1)" format ---
        If InStr("RZXYWSTUV", Left$(Value, 1)) > 0 AndAlso Mid$(Value, 2, 1) = "(" AndAlso Mid$(Value, 4, 1) = ")" Then
            If InStr("1", Mid$(Value, 3, 1)) > 0 Then Return True
        End If
        ' --- check for buffer pointers ---
        If IsBufferPtrByValue(Value) Then Return True
        ' --- non-byte system variables ---
        If Value = "OPER" Then Return False
        If Value = "ORIG" Then Return False
        If Value = "PROG" Then Return False
        If Value = "USER" Then Return False
        ' --- otherwise check if system variable ---
        Return IsSysVar(Value)
    End Function

    Private Function IsSysVar(ByVal Value As String) As Boolean
        If Value = "CANVAL" Then Return True
        If Value = "CHAR" Then Return True
        If Value = "ESCVAL" Then Return True
        If Value = "ITYPE" Then Return True
        If Value = "KBC" Then Return True
        If Value = "KBCX" Then Return True
        If Value = "LANG" Then Return True
        If Value = "LENGTH" Then Return True
        If Value = "LIB" Then Return True
        If Value = "LOCKVAL" Then Return True
        If Value = "MACHTYPE" Then Return True
        If Value = "OPER" Then Return True
        If Value = "ORIG" Then Return True
        If Value = "PRIVG" Then Return True
        If Value = "PROG" Then Return True
        If Value = "PRTNUM" Then Return True
        If Value = "PVOL" Then Return True
        If Value = "REQVOL" Then Return True
        If Value = "SEG" Then Return True
        If Value = "STATUS" Then Return True
        If Value = "SYSREL" Then Return True
        If Value = "SYSREV" Then Return True
        If Value = "TCHAN" Then Return True
        If Value = "TERM" Then Return True
        If Value = "TFA" Then Return True
        If Value = "USER" Then Return True
        If Value = "VOL" Then Return True
        Return False
    End Function

    Private Function IsNumericExpr(ByVal Value As String) As Boolean
        Dim TempValue As Integer
        ' ----------------------
        If Left$(Value, 1) = "(" Then Return True
        If Left$(Value, 1) = "-" Then Return True
        If Left$(Value, 6) = "DIVREM" Then Return True
        If Left$(Value, 1) >= "0" AndAlso Left$(Value, 1) <= "9" Then Return True
        TempValue = InStr(Value, " ")
        If TempValue > 0 Then
            If IsNumericItem(Left$(Value, TempValue - 1)) Then Return True
        Else
            If IsNumericItem(Value) Then Return True
        End If
        Return False
    End Function

    Private Function GetTarget(ByRef Tokens() As String, ByRef TokenNum As Integer) As String
        Dim Result As String
        ' -----------------------------
        Result = GetItem(Tokens, TokenNum)
        If InStr(Result, "_OFS(") > 0 AndAlso Right$(Result, 1) = ")" Then
            Result = Replace(Result, "_OFS(", "_OFS ")
            Result = Left$(Result, Len(Result) - 1) + ","
        End If
        If InStr(Result, "(") = 2 AndAlso Right$(Result, 1) = ")" Then
            Result = Replace(Result, "(", " ")
            Result = Left$(Result, Len(Result) - 1) + ","
        End If
        Return Result
    End Function

    Private Sub AddJumpPoint(ByVal Value As Integer)
        ' --- keep JumpPointList sorted ---
        Dim TempValue As Integer = Value
        Dim CurrValue As Integer
        For LoopNum As Integer = 0 To JumpPointList.Count - 1
            CurrValue = JumpPointList(LoopNum)
            If CurrValue = TempValue Then Exit Sub ' already exists
            If CurrValue > TempValue Then
                JumpPointList(LoopNum) = TempValue
                TempValue = CurrValue
            End If
        Next
        JumpPointList.Add(TempValue)
    End Sub

    Private Sub AddGotoLine(ByVal Value As Integer)
        ' --- keep GotoLineList sorted ---
        Dim TempValue As Integer = Value
        Dim CurrValue As Integer
        For LoopNum As Integer = 0 To GotoLineList.Count - 1
            CurrValue = GotoLineList(LoopNum)
            If CurrValue = TempValue Then Exit Sub ' already exists
            If CurrValue > TempValue Then
                GotoLineList(LoopNum) = TempValue
                TempValue = CurrValue
            End If
        Next
        GotoLineList.Add(TempValue)
    End Sub

    Private Function GetCompare(ByRef Tokens() As String, ByRef TokenNum As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim Result As New StringBuilder
        Dim FinalResult As String
        ' -----------------------------
        Result.Append(GetExpression(Tokens, TokenNum))
        Result.Append(" ")
        ' --- check for comparison operator ---
        If Tokens(TokenNum) <> "=" AndAlso
           Tokens(TokenNum) <> "<>" AndAlso
           Tokens(TokenNum) <> "#" AndAlso
           Tokens(TokenNum) <> "<" AndAlso
           Tokens(TokenNum) <> "<=" AndAlso
           Tokens(TokenNum) <> ">" AndAlso
           Tokens(TokenNum) <> ">=" Then
            GoTo ErrorFound
        End If
        If Tokens(TokenNum) = "#" Then
            Tokens(TokenNum) = "<>"
        End If
        Result.Append(Tokens(TokenNum)) ' operator
        Result.Append(" ")
        TokenNum += 1
        Result.Append(GetExpression(Tokens, TokenNum))
        FinalResult = Result.ToString
        If UsesSameBufferTwice(FinalResult) Then GoTo ErrorFound
        Return FinalResult
        Exit Function
ErrorFound:
        Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Invalid Comparison")
    End Function

    Private Function GetExpression(ByRef Tokens() As String, ByRef TokenNum As Integer) As String

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Const LastLevel As Integer = 31
        Dim Levels(LastLevel) As String
        Dim Oper(LastLevel) As String
        Dim Unary(LastLevel) As String
        Dim Pending(LastLevel) As String
        Dim strCurr As String
        Dim intLevel As Integer
        Dim BuffCount(8) As Integer
        ' -----------------------

        For intLevel = 0 To LastLevel
            Levels(intLevel) = ""
            Oper(intLevel) = ""
            Unary(intLevel) = ""
            Pending(intLevel) = ""
        Next intLevel
        intLevel = 0

NextToken:

        If TokenNum > UBound(Tokens) Then GoTo Done

        strCurr = GetItem(Tokens, TokenNum)

        ' --- check for unary minus ---
        If strCurr = "-" Then
            If Levels(intLevel) = "" Then
                Levels(intLevel) = "-"
                GoTo NextToken
            End If
            If Oper(intLevel) <> "" AndAlso Pending(intLevel) = "" Then
                Unary(intLevel) = "-"
                GoTo NextToken
            End If
        End If

        ' --- check for string concatenation ---
        If strCurr = "&" Then
            If Oper(intLevel) = "" Then
                Oper(intLevel) = strCurr
                GoTo NextToken
            End If
            ' --- check for errors ---
            If Oper(intLevel) <> "&" Then GoTo ErrorFound
            If Pending(intLevel) = "" Then GoTo ErrorFound
            ' --- concatenate the items ---
            Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Pending(intLevel)
            Oper(intLevel) = strCurr
            Pending(intLevel) = ""
            GoTo NextToken
        End If

        ' --- check for numeric operators ---
        If strCurr = "+" OrElse strCurr = "-" OrElse strCurr = "*" OrElse strCurr = "/" Then
            If Oper(intLevel) = "" Then
                Oper(intLevel) = strCurr
                GoTo NextToken
            End If
            If Pending(intLevel) = "" Then GoTo ErrorFound
        End If

        ' --- check for order of precedence problems ---
        If strCurr = "+" OrElse strCurr = "-" Then
            If Oper(intLevel) = "*" OrElse Oper(intLevel) = "/" Then
                If intLevel > 0 Then
                    ' --- problem only occurs when multi-level subtraction is involved ---
                    If strCurr = "-" OrElse Oper(intLevel - 1) = "-" Then
                        GoTo ErrorFound
                    End If
                End If
            End If
        End If

        ' --- check for addition/subtraction statements ---
        If strCurr = "+" OrElse strCurr = "-" Then
            If Oper(intLevel) = "+" OrElse Oper(intLevel) = "-" Then
                Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + Pending(intLevel)
            ElseIf Oper(intLevel) = "*" OrElse Oper(intLevel) = "/" Then
                If Oper(intLevel) = "*" Then
                    Levels(intLevel) = "(" + Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
                Else
                    Levels(intLevel) = "DIVREM(" + Levels(intLevel) + ", " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
                End If
                ' --- handle the (x*y) or DIVREM(x,y) as a single unit ---
                If intLevel > 0 Then
                    If Oper(intLevel - 1) = "+" OrElse Oper(intLevel - 1) = "-" Then
                        Levels(intLevel - 1) = Levels(intLevel - 1) + " " + Oper(intLevel - 1) + " " +
                                               Unary(intLevel - 1) + Pending(intLevel - 1) + Levels(intLevel)
                        Oper(intLevel - 1) = ""
                        Unary(intLevel - 1) = ""
                        Pending(intLevel - 1) = ""
                        Levels(intLevel) = ""
                        Oper(intLevel) = ""
                        Unary(intLevel) = ""
                        Pending(intLevel) = ""
                        intLevel = intLevel - 1
                    End If
                End If
            Else
                GoTo ErrorFound
            End If
            Oper(intLevel) = strCurr
            Unary(intLevel) = ""
            Pending(intLevel) = ""
            GoTo NextToken
        End If

        ' --- check for multiplication/division intermixed with addition/subtraction ---
        If strCurr = "*" OrElse strCurr = "/" Then
            If Oper(intLevel) = "+" OrElse Oper(intLevel) = "-" Then
                intLevel = intLevel + 1
                If intLevel > LastLevel Then GoTo ErrorFound
                Levels(intLevel) = Unary(intLevel - 1) + Pending(intLevel - 1)
                Oper(intLevel) = strCurr
                Unary(intLevel) = ""
                Pending(intLevel) = ""
                Unary(intLevel - 1) = ""
                Pending(intLevel - 1) = ""
                GoTo NextToken
            End If
        End If

        ' --- check for chained multiplication/division statements ---
        If strCurr = "*" OrElse strCurr = "/" Then
            If Oper(intLevel) = "*" Then
                Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + Pending(intLevel)
            ElseIf Oper(intLevel) = "/" Then
                Levels(intLevel) = "DIVREM(" + Levels(intLevel) + ", " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
            Else
                GoTo ErrorFound
            End If
            Oper(intLevel) = strCurr
            Unary(intLevel) = ""
            Pending(intLevel) = ""
            GoTo NextToken
        End If

        ' --- check for left parenthesis ---
        If strCurr = "(" Then
            If Pending(intLevel) <> "" Then GoTo ErrorFound
            Pending(intLevel) = "("
            intLevel = intLevel + 1
            If intLevel > LastLevel Then GoTo ErrorFound
            Levels(intLevel) = ""
            Oper(intLevel) = ""
            Unary(intLevel) = ""
            Pending(intLevel) = ""
            GoTo NextToken
        End If

        ' --- check for right parenthesis ---
        If strCurr = ")" Then
RightParenAgain:
            If Oper(intLevel) = "+" OrElse Oper(intLevel) = "-" OrElse Oper(intLevel) = "*" Then
                Levels(intLevel) = "(" + Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
                Oper(intLevel) = ""
                Unary(intLevel) = ""
                Pending(intLevel) = ""
            End If
            If Oper(intLevel) = "/" Then
                Levels(intLevel) = "DIVREM(" + Levels(intLevel) + ", " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
                Oper(intLevel) = ""
                Unary(intLevel) = ""
                Pending(intLevel) = ""
            End If
            If Oper(intLevel) <> "" Then GoTo ErrorFound
            intLevel = intLevel - 1
            If intLevel < 0 Then GoTo ErrorFound
            ' --- check for pushed level due to order of precidence ---
            If Pending(intLevel) = "" Then
                Pending(intLevel) = Levels(intLevel + 1)
                Levels(intLevel + 1) = ""
                GoTo RightParenAgain
            End If
            ' --- found proper level for right parenthesis ---
            If Pending(intLevel) <> "(" Then GoTo ErrorFound
            Pending(intLevel) = RTrim$(Levels(intLevel + 1))
            Levels(intLevel + 1) = ""
            ' --- check if level started with a parenthesis ---
            If (Levels(intLevel) = "" OrElse Levels(intLevel) = "-") AndAlso Oper(intLevel) = "" Then
                Levels(intLevel) = Levels(intLevel) + Unary(intLevel) + Pending(intLevel)
                Oper(intLevel) = ""
                Unary(intLevel) = ""
                Pending(intLevel) = ""
            End If
            If intLevel = 0 Then GoTo CheckIfDone
            GoTo NextToken
        End If

        ' --- add value to current level ---
        If Pending(intLevel) <> "" Then GoTo ErrorFound
        If Levels(intLevel) = "" Then
            Levels(intLevel) = strCurr
        ElseIf Levels(intLevel) = "-" Then
            Levels(intLevel) = "-" + strCurr
        Else
            Pending(intLevel) = strCurr
        End If

CheckIfDone:

        If TokenNum > UBound(Tokens) Then GoTo Done
        strCurr = Tokens(TokenNum)
        If strCurr = "!" Then GoTo Done
        If strCurr = "+" OrElse strCurr = "-" OrElse strCurr = "*" OrElse strCurr = "/" Then
            GoTo NextToken
        End If
        If strCurr = "&" Then
            GoTo NextToken
        End If
        If strCurr = ")" AndAlso intLevel > 0 Then
            GoTo NextToken
        End If

Done:

        Do While intLevel > 0
            If Oper(intLevel) = "+" OrElse Oper(intLevel) = "-" OrElse Oper(intLevel) = "*" Then
                Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + Pending(intLevel)
                Oper(intLevel) = ""
                Unary(intLevel) = ""
                Pending(intLevel) = ""
                Pending(intLevel - 1) = "(" + RTrim$(Levels(intLevel)) + ")"
                Levels(intLevel) = ""
            End If
            If Oper(intLevel) = "/" Then
                Levels(intLevel) = "DIVREM(" + Levels(intLevel) + ", " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
                Oper(intLevel) = ""
                Unary(intLevel) = ""
                Pending(intLevel) = ""
                Pending(intLevel - 1) = Levels(intLevel) ' don't need parenthesis with DIVREM
                Levels(intLevel) = ""
            End If
            If Oper(intLevel) <> "" Then GoTo ErrorFound
            intLevel = intLevel - 1
        Loop

        If Oper(intLevel) = "+" OrElse Oper(intLevel) = "-" OrElse Oper(intLevel) = "*" Then
            Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + Pending(intLevel)
            Oper(intLevel) = ""
            Unary(intLevel) = ""
            Pending(intLevel) = ""
        End If
        If Oper(intLevel) = "&" Then
            Levels(intLevel) = Levels(intLevel) + " " + Oper(intLevel) + " " + Unary(intLevel) + Pending(intLevel)
            Oper(intLevel) = ""
            Unary(intLevel) = ""
            Pending(intLevel) = ""
        End If
        If Oper(intLevel) = "/" Then
            Levels(intLevel) = "DIVREM(" + Levels(intLevel) + ", " + Unary(intLevel) + RTrim$(Pending(intLevel)) + ")"
            Oper(intLevel) = ""
            Unary(intLevel) = ""
            Pending(intLevel) = ""
        End If

        ' --- get final expression result ---
        strCurr = RTrim$(Levels(0))

        ' --- check for multiple calls to same buffer in the same expression ---
        If Left$(strCurr, 1) <> """" Then
            If UsesSameBufferTwice(strCurr) Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Same buffer occurs more than once in the expression.")
            End If
        End If

        ' --- done ---
        Return strCurr

        Exit Function

ErrorFound:

        Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Invalid Expression")

    End Function

    Private Function GetItem(ByRef Tokens() As String, ByRef TokenNum As Integer) As String

        ' --- Note: Must use "... = GetItem()" to call a function recursively! ---
        ' ---       "GetItem" is a value which can be referenced, "GetItem()"  ---
        ' ---       is a recursive call to this function.                      ---

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim strTemp As String
        Dim strTemp2 As String
        Dim strTemp3 As String
        Dim strResult As String
        ' ---------------------

        strResult = Tokens(TokenNum)
        TokenNum += 1

        ' --- check for literal string ---
        If strResult.StartsWith("""") AndAlso strResult.EndsWith("""") Then
            GoTo Done
        End If
        'If (strResult.StartsWith("'") AndAlso strResult.EndsWith("'")) OrElse _
        '   (strResult.StartsWith("%") AndAlso strResult.EndsWith("%")) OrElse _
        '   (strResult.StartsWith("$") AndAlso strResult.EndsWith("$")) Then
        '    strResult = strResult.Replace("""", """""")
        '    strResult = """" + strResult.Substring(1, strResult.Length - 2) + """"
        '    GoTo Done
        'End If

        ' --- check for ambiguous commands ---
        If strResult = "ESC" Then strResult = "ESCVAL"
        If strResult = "CAN" Then strResult = "CANVAL"
        If strResult = "LOCK" Then strResult = "LOCKVAL"
        If strResult = "REM" Then strResult = "REMVAL"
        If strResult = "DATE" Then strResult = "DATEVAL"

        ' --- check for buffer ---
        If Len(strResult) = 1 AndAlso InStr("RZXYWSTUV", strResult) > 0 Then

            If TokenNum > UBound(Tokens) Then
                strResult = strResult + "_A"
                GoTo Done
            End If

            strTemp2 = "" ' holds offset
            strTemp3 = "" ' holds numeric byte value

            If Tokens(TokenNum) = "[" Then
                TokenNum += 1
                strTemp2 = "_OFS(" + GetExpression(Tokens, TokenNum)
                strTemp3 = ")"
                If Tokens(TokenNum) <> "]" Then GoTo ErrorFound
                TokenNum += 1
            End If

            If TokenNum > UBound(Tokens) Then
                strResult = strResult + "_A"
            Else
                If Tokens(TokenNum) = "(" Then
                    If Tokens(TokenNum + 2) <> ")" Then GoTo ErrorFound
                    ' --- check for alpha ---
                    If Tokens(TokenNum + 1) = "A" Then
                        strResult = strResult + "_A" ' ?_A
                        TokenNum = TokenNum + 3
                    Else ' --- numeric ---
                        If strTemp2 = "" Then
                            strTemp2 = "("
                        Else
                            strTemp2 = strTemp2 + ","
                        End If
                        strTemp3 = RTrim$(Tokens(TokenNum + 1)) + ")"
                        TokenNum = TokenNum + 3
                    End If
                Else
                    strResult = strResult + "_A"
                End If
            End If

            strResult = strResult + strTemp2 + strTemp3
            GoTo Done

        End If

        ' --- check for alpha with offset ---
        If TokenNum < UBound(Tokens) Then
            If Tokens(TokenNum) = "[" Then
                TokenNum += 1
                strTemp2 = GetExpression(Tokens, TokenNum)
                If Tokens(TokenNum) <> "]" Then GoTo ErrorFound
                TokenNum += 1
                strResult = strResult + "_OFS(" + RTrim$(strTemp2) + ")"
                GoTo Done
            End If
        End If

        ' --- check for direct memory accesses ---
        If strResult = "SYSVAR" Then
            If Tokens(TokenNum) <> "(" Then GoTo ErrorFound
            TokenNum += 1
            strTemp = GetExpression(Tokens, TokenNum)
            If Tokens(TokenNum) <> ")" Then GoTo ErrorFound
            TokenNum += 1
            strResult = "MEM(" + RTrim$(strTemp) + ")"
            GoTo Done
        End If

        ' --- check for large numeric value ---
        If NumOnly(strResult) AndAlso Len(strResult) > 9 Then
            strResult = strResult + "@" ' currency
        End If

Done:

        Return strResult

        Exit Function

ErrorFound:

        Throw New SystemException(FuncName + vbCrLf + vbCrLf + "Invalid Expression")

    End Function

    Private Function UsesSameBufferTwice(ByVal Value As String) As Boolean
        Dim strTemp As String
        ' -------------------
        strTemp = Value
        ' --- need to add spaces so "S(" isn't found for "A_OFS(" ---
        Do While InStr(strTemp, "_OFS(") > 0
            strTemp = Replace(strTemp, "_OFS(", "_OFS (")
        Loop
        ' --- also need to handle "KEY_OFS" not looking like "Y_OFS" ---
        Do While InStr(strTemp, "KEY_OFS (") > 0
            strTemp = Replace(strTemp, "KEY_OFS (", "KEY_OFS  (")
        Loop
        ' --- look for same buffer used twice in the line ---
        If InStr(strTemp, "R_OFS (") <> InStrRev(strTemp, "R_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "Z_OFS (") <> InStrRev(strTemp, "Z_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "X_OFS (") <> InStrRev(strTemp, "X_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "Y_OFS (") <> InStrRev(strTemp, "Y_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "W_OFS (") <> InStrRev(strTemp, "W_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "S_OFS (") <> InStrRev(strTemp, "S_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "T_OFS (") <> InStrRev(strTemp, "T_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "U_OFS (") <> InStrRev(strTemp, "U_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "V_OFS (") <> InStrRev(strTemp, "V_OFS (") Then GoTo ErrorFound
        If InStr(strTemp, "R(") > 0 AndAlso InStr(strTemp, "R_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "Z(") > 0 AndAlso InStr(strTemp, "Z_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "X(") > 0 AndAlso InStr(strTemp, "X_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "Y(") > 0 AndAlso InStr(strTemp, "Y_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "W(") > 0 AndAlso InStr(strTemp, "W_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "S(") > 0 AndAlso InStr(strTemp, "S_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "T(") > 0 AndAlso InStr(strTemp, "T_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "U(") > 0 AndAlso InStr(strTemp, "U_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "V(") > 0 AndAlso InStr(strTemp, "V_OFS (") > 0 Then GoTo ErrorFound
        If InStr(strTemp, "R(") <> InStrRev(strTemp, "R(") Then GoTo ErrorFound
        If InStr(strTemp, "Z(") <> InStrRev(strTemp, "Z(") Then GoTo ErrorFound
        If InStr(strTemp, "X(") <> InStrRev(strTemp, "X(") Then GoTo ErrorFound
        If InStr(strTemp, "Y(") <> InStrRev(strTemp, "Y(") Then GoTo ErrorFound
        If InStr(strTemp, "W(") <> InStrRev(strTemp, "W(") Then GoTo ErrorFound
        If InStr(strTemp, "S(") <> InStrRev(strTemp, "S(") Then GoTo ErrorFound
        If InStr(strTemp, "T(") <> InStrRev(strTemp, "T(") Then GoTo ErrorFound
        If InStr(strTemp, "U(") <> InStrRev(strTemp, "U(") Then GoTo ErrorFound
        If InStr(strTemp, "V(") <> InStrRev(strTemp, "V(") Then GoTo ErrorFound
        Return False
        Exit Function
ErrorFound:
        Return True
    End Function

#End Region

#Region " --- Fifth Pass Routines --- "

    Private Sub PerformFifthPass(ByRef Lines1 As List(Of String), ByRef Lines2 As List(Of String))
        Dim LineNum As Integer
        Dim CurrLine As String
        Dim TempProgNum As String
        ' -----------------------
        TempProgNum = ObjProgNum.ToString.PadLeft(3, "0"c)
        Lines2.Clear()
        If Not m_ToPath.EndsWith("\_IDRISYS") Then
            Lines2.Add("Attribute VB_Name = ""modProg" + TempProgNum + """")
        Else
            Lines2.Add("Attribute VB_Name = ""modSysProg" + TempProgNum + """")
        End If
        Lines2.Add("Option Explicit")
        Lines2.Add("")
        If Not m_ToPath.EndsWith("\_IDRISYS") Then
            Lines2.Add("Public Sub PROG_" + TempProgNum + "(ByVal JUMPPOINT As Long)")
        Else
            Lines2.Add("Public Sub SYSPROG_" + TempProgNum + "(ByVal JUMPPOINT As Long)")
        End If
        For Each LineNum In JumpPointList
            Lines2.Add("       IF JUMPPOINT = " + LineNum.ToString + " THEN GOTO " + LineNum.ToString)
        Next
        Lines2.Add("       FATALERROR ""UNKNOWN JUMPPOINT"" : EXIT SUB")
        For LineNum = 0 To Lines1.Count - 1
            CurrLine = Lines1(LineNum)
            If CurrLine.StartsWith(" ") Then
                If JumpPointList.Contains(LineNum) OrElse GotoLineList.Contains(LineNum) Then
                    CurrLine = (LineNum.ToString + ":       ").Substring(0, 7) + CurrLine.Substring(7)
                ElseIf CurrLine.IndexOf(" : ") > 0 Then
                    ' --- check for single-word command followed by ":" ---
                    If CurrLine.Substring(0, CurrLine.IndexOf(" : ")).Trim.IndexOf(" ") < 0 Then
                        CurrLine = (LineNum.ToString + ":       ").Substring(0, 7) + CurrLine.Substring(7)
                    End If
                End If
            End If
            Lines2.Add(CurrLine)
        Next
        Lines2.Add("End Sub")
    End Sub

#End Region

#Region " --- Build VB6 Project File --- "

    Public Sub BuildVB6ProjectFile(ByVal DestPath As String,
                                   ByVal DestLibName As String,
                                   ByVal CommonPath As String)

        Dim VBDestFilename As String
        Dim VBDestFile As StreamWriter
        Dim ProgramNum As Integer
        Dim Filenames() As String
        Dim TempFilename As String
        Dim IDRISYSPath As String
        Dim FixedDestLibName As String
        ' ----------------------------

        FixedDestLibName = DestLibName.Replace("/"c, "_"c).Replace("%"c, "_"c)

        ' --- get previous version from VB6 project file ---
        VBDestFilename = DestPath + "\" + DestLibName + "\LIB_" + FixedDestLibName + ".vbp"

        ' --- build VB6 project file ---
        VBDestFile = New StreamWriter(VBDestFilename)

        With VBDestFile

            .WriteLine("Type=Exe")

            ' --- object references ---
            .WriteLine("Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#" +
                       "C:\WINDOWS\system32\STDOLE2.TLB#OLE Automation")
            .WriteLine("Reference=*\G{EF53050B-882E-4776-B643-EDA472E8E3F2}#2.7#0#" +
                       "C:\Program Files\Common Files\System\ADO\msado27.tlb#" +
                       "Microsoft ActiveX Data Objects 2.7 Library")
            .WriteLine("Object={248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0; MSWINSCK.OCX")

            ' --- runtime classes ---
            Filenames = Directory.GetFiles(CommonPath, "rt*.cls", SearchOption.TopDirectoryOnly)
            For Each TempFilename In Filenames
                TempFilename = TempFilename.Substring(TempFilename.LastIndexOf("\"c) + 1)
                .WriteLine("Class=" + TempFilename.Substring(0, Len(TempFilename) - 4) +
                           "; ..\..\..\COMMON\" + TempFilename)
            Next

            ' --- runtime modules ---
            Filenames = Directory.GetFiles(CommonPath, "rt*.bas", SearchOption.TopDirectoryOnly)
            For Each TempFilename In Filenames
                TempFilename = TempFilename.Substring(TempFilename.LastIndexOf("\"c) + 1)
                .WriteLine("Module=" + TempFilename.Substring(0, Len(TempFilename) - 4) +
                           "; ..\..\..\COMMON\" + TempFilename)
            Next

            ' --- runtime forms ---
            Filenames = Directory.GetFiles(CommonPath, "rt*.frm", SearchOption.TopDirectoryOnly)
            For Each TempFilename In Filenames
                TempFilename = TempFilename.Substring(TempFilename.LastIndexOf("\"c) + 1)
                .WriteLine("Form=..\..\..\COMMON\" + TempFilename)
            Next
            ' --- find property _IDRISYS path ---
            IDRISYSPath = "..\.." ' always assume DEVICE00
            ' --- look for "_SYSVOL" ---
            IDRISYSPath += "\_SYSVOL"
            ' --- look for "_IDRISYS" ---
            IDRISYSPath += "\_IDRISYS"
            ' --- _USERLIB routines ---
            .WriteLine("Module=modSysProg000; " + IDRISYSPath + "\modSysProg000.bas")
            .WriteLine("Module=modSysProg001; " + IDRISYSPath + "\modSysProg001.bas")
            ' --- _IDRISYS routines ---
            .WriteLine("Module=modSysProg005; " + IDRISYSPath + "\modSysProg005.bas")
            .WriteLine("Module=modSysProg009; " + IDRISYSPath + "\modSysProg009.bas")
            .WriteLine("Module=modSysProg010; " + IDRISYSPath + "\modSysProg010.bas")
            .WriteLine("Module=modSysProg056; " + IDRISYSPath + "\modSysProg056.bas")
            .WriteLine("Module=modSysProg091; " + IDRISYSPath + "\modSysProg091.bas")
            .WriteLine("Module=modSysProg112; " + IDRISYSPath + "\modSysProg112.bas")
            .WriteLine("Module=modSysProg141; " + IDRISYSPath + "\modSysProg141.bas")
            .WriteLine("Module=modSysProg142; " + IDRISYSPath + "\modSysProg142.bas")
            .WriteLine("Module=modSysProg156; " + IDRISYSPath + "\modSysProg156.bas")
            .WriteLine("Module=modSysProg160; " + IDRISYSPath + "\modSysProg160.bas")
            .WriteLine("Module=modSysProg161; " + IDRISYSPath + "\modSysProg161.bas")
            .WriteLine("Module=modSysProg185; " + IDRISYSPath + "\modSysProg185.bas")
            .WriteLine("Module=modSysProg186; " + IDRISYSPath + "\modSysProg186.bas")
            .WriteLine("Module=modSysProg187; " + IDRISYSPath + "\modSysProg187.bas")
            .WriteLine("Module=modSysProg195; " + IDRISYSPath + "\modSysProg195.bas")
            .WriteLine("Module=modSysProg203; " + IDRISYSPath + "\modSysProg203.bas")
            .WriteLine("Module=modSysProg204; " + IDRISYSPath + "\modSysProg204.bas")
            .WriteLine("Module=modSysProg205; " + IDRISYSPath + "\modSysProg205.bas")
            .WriteLine("Module=modSysProg212; " + IDRISYSPath + "\modSysProg212.bas")
            .WriteLine("Module=modSysProg222; " + IDRISYSPath + "\modSysProg222.bas")
            .WriteLine("Module=modSysProg224; " + IDRISYSPath + "\modSysProg224.bas")
            .WriteLine("Module=modSysProg225; " + IDRISYSPath + "\modSysProg225.bas")
            .WriteLine("Module=modSysProg226; " + IDRISYSPath + "\modSysProg226.bas")
            ' --- more _USERLIB routines ---
            .WriteLine("Module=modSysProg230; " + IDRISYSPath + "\modSysProg230.bas")
            .WriteLine("Module=modSysProg231; " + IDRISYSPath + "\modSysProg231.bas")
            .WriteLine("Module=modSysProg232; " + IDRISYSPath + "\modSysProg232.bas")
            .WriteLine("Module=modSysProg233; " + IDRISYSPath + "\modSysProg233.bas")
            .WriteLine("Module=modSysProg234; " + IDRISYSPath + "\modSysProg234.bas")
            .WriteLine("Module=modSysProg235; " + IDRISYSPath + "\modSysProg235.bas")
            .WriteLine("Module=modSysProg236; " + IDRISYSPath + "\modSysProg236.bas")
            .WriteLine("Module=modSysProg239; " + IDRISYSPath + "\modSysProg239.bas")
            .WriteLine("Module=modSysProg241; " + IDRISYSPath + "\modSysProg241.bas")
            .WriteLine("Module=modSysProg242; " + IDRISYSPath + "\modSysProg242.bas")
            .WriteLine("Module=modSysProg243; " + IDRISYSPath + "\modSysProg243.bas")
            .WriteLine("Module=modSysProg244; " + IDRISYSPath + "\modSysProg244.bas")
            .WriteLine("Module=modSysProg245; " + IDRISYSPath + "\modSysProg245.bas")
            ' --- more _IDRISYS routines ---
            .WriteLine("Module=modSysProg248; " + IDRISYSPath + "\modSysProg248.bas")
            .WriteLine("Module=modSysProg249; " + IDRISYSPath + "\modSysProg249.bas")
            .WriteLine("Module=modSysProg250; " + IDRISYSPath + "\modSysProg250.bas")
            .WriteLine("Module=modSysProg251; " + IDRISYSPath + "\modSysProg251.bas")
            .WriteLine("Module=modSysProg254; " + IDRISYSPath + "\modSysProg254.bas")

            For ProgramNum = 0 To 255
                .WriteLine("Module=modProg" + ProgramNum.ToString.PadLeft(3, "0"c) +
                           "; modProg" + ProgramNum.ToString.PadLeft(3, "0"c) + ".bas")
            Next

            .WriteLine("Startup=""Sub Main""")
            .WriteLine("ExeName32=""LIB_" + FixedDestLibName.ToUpper + ".exe""")
            .WriteLine("Path32=""..""")
            .WriteLine("Command32=""""")
            .WriteLine("Name=""LIB_" + FixedDestLibName.ToUpper + """")
            .WriteLine("HelpContextID=""0""")
            .WriteLine("CompatibleMode=""0""")
            .WriteLine("MajorVer=" + Year(Today).ToString)
            .WriteLine("MinorVer=" + Month(Today).ToString)
            .WriteLine("RevisionVer=" + Day(Today).ToString)
            .WriteLine("AutoIncrementVer=0") ' don't turn this on - causes .VBP to change when compiled
            .WriteLine("ServerSupportFiles=0")
            .WriteLine("VersionCompanyName=""Custom Disability Solutions""")
            .WriteLine("CompilationType=0")
            .WriteLine("OptimizationType=0")
            .WriteLine("FavorPentiumPro(tm)=0")
            .WriteLine("CodeViewDebugInfo=0")
            .WriteLine("NoAliasing=0")
            .WriteLine("BoundsCheck=0")
            .WriteLine("OverflowCheck=0")
            .WriteLine("FlPointCheck=0")
            .WriteLine("FDIVCheck=0")
            .WriteLine("UnroundedFP=0")
            .WriteLine("StartMode=0")
            .WriteLine("Unattended=0")
            .WriteLine("Retained=0")
            .WriteLine("ThreadPerObject=0")
            .WriteLine("MaxNumberOfThreads=1")
            .WriteLine()
            .WriteLine("[MS Transaction Server]")
            .WriteLine("AutoRefresh=1")

            .Close()

        End With

    End Sub

#End Region

#Region " --- Common Atomic Routines --- "

    Private Function NumOnly(ByVal Value As String) As Boolean
        Dim CharNum As Integer
        ' --------------------
        For CharNum = 0 To Value.Length - 1
            If Not Char.IsDigit(Value(CharNum)) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Function IsBufferPtrByValue(ByVal Value As String) As Boolean
        Select Case Value
            Case "RP"
            Case "RP2"
            Case "IRP"
            Case "IRP2"
            Case "ZP"
            Case "ZP2"
            Case "IZP"
            Case "IZP2"
            Case "XP"
            Case "XP2"
            Case "IXP"
            Case "IXP2"
            Case "YP"
            Case "YP2"
            Case "IYP"
            Case "IYP2"
            Case "WP"
            Case "WP2"
            Case "IWP"
            Case "IWP2"
            Case "SP"
            Case "SP2"
            Case "ISP"
            Case "ISP2"
            Case "TP"
            Case "TP2"
            Case "ITP"
            Case "ITP2"
            Case "UP"
            Case "UP2"
            Case "IUP"
            Case "IUP2"
            Case "VP"
            Case "VP2"
            Case "IVP"
            Case "IVP2"
            Case Else
                Return False
        End Select
        Return True
    End Function

    Private Function IsSystemVarByValue(ByVal Value As String) As Boolean
        Select Case Value
            Case "CAN"
            Case "CHAR"
            Case "ESC"
            Case "ITYPE"
            Case "KBC"
            Case "KBCX"
            Case "LANG"
            Case "LENGTH"
            Case "LIB"
            Case "LOCK"
            Case "MACHTYPE"
            Case "OPER"
            Case "ORIG"
            Case "PRIVG"
            Case "PROG"
            Case "PRTNUM"
            Case "PVOL"
            Case "REQVOL"
            Case "SEG"
            Case "STATUS"
            Case "SYSREL"
            Case "SYSREV"
            Case "TCHAN"
            Case "TERM"
            Case "TFA"
            Case "USER"
            Case "VOL"
            Case Else
                Return False
        End Select
        Return True
    End Function

#End Region

#Region " --- RENAME Fill with defaults --- "

    Private Sub FillRenamesWithDefaults()
        Renames.Clear()
        For Index As Integer = 0 To 99
            If Index = 0 Then
                Renames.Add("N", "N")
            Else
                Renames.Add($"N{Index}", $"N{Index}")
            End If
        Next
        For Index As Integer = 0 To 99
            If Index = 0 Then
                Renames.Add("F", "F")
            Else
                Renames.Add($"F{Index}", $"F{Index}")
            End If
        Next
        For Index As Integer = 0 To 9
            If Index = 0 Then
                Renames.Add("A", "A")
                Renames.Add("B", "B")
                Renames.Add("C", "C")
                Renames.Add("D", "D")
                Renames.Add("E", "E")
            Else
                Renames.Add($"A{Index}", $"A{Index}")
                Renames.Add($"B{Index}", $"B{Index}")
                Renames.Add($"C{Index}", $"C{Index}")
                Renames.Add($"D{Index}", $"D{Index}")
                Renames.Add($"E{Index}", $"E{Index}")
            End If
        Next
    End Sub

#End Region

End Class
