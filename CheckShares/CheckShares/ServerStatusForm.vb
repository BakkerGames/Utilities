Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Threading
Imports System.ComponentModel

Public Class ServerStatusForm
    Private Const SERVER_LIST_TEXT_FILENAME As String = "shares.txt"
    Private Const OFFLINE As Boolean = False
    Private _pathList As List(Of String)
    Private _pathInfo As Dictionary(Of String, PathInfo)
    Private _statusRecords As Dictionary(Of String, StatusRecord)
    Protected _rwLock As ReaderWriterLock
    Private WithEvents bgWorker As BackgroundWorker


    Public Sub New()
        InitializeComponent()

        _rwLock = New ReaderWriterLock
        _statusRecords = New Dictionary(Of String, StatusRecord)
    End Sub

    Private Sub ServerStatusForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Enabled = False

        ReadServerList()
        PopulateFormWithServerList()

        Enabled = True

        AfterLoadTimer.Interval = 250  ' 1/4 of a second
        AfterLoadTimer.Enabled = True

    End Sub

    Private Sub AfterLoadTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles AfterLoadTimer.Tick
        ' Turn off timer, so it only fires once, after loading the form
        AfterLoadTimer.Enabled = False

        GetAllPathInfo()
    End Sub

    Private Sub ReadServerList()
        _pathList = New List(Of String)
        _pathInfo = New Dictionary(Of String, PathInfo)
        Dim path As String

        Using reader As New StreamReader(SERVER_LIST_TEXT_FILENAME)
            Dim line As String = reader.ReadLine
            While line IsNot Nothing
                path = line.Trim

                If path.Length > 0 _
                AndAlso (path.StartsWith("\\") OrElse (path.StartsWith("""\\") AndAlso path.EndsWith(""""))) _
                AndAlso Not _pathInfo.ContainsKey(path) Then
                    _pathList.Add(path)
                    _pathInfo.Add(path, New PathInfo)
                End If

                line = reader.ReadLine
            End While
        End Using
    End Sub


    Private Sub PopulateFormWithServerList()
        For Each curPath As String In _pathInfo.Keys
            _statusRecords.Add(curPath, New StatusRecord(curPath))
            StatusPanel.Controls.Add(_statusRecords(curPath).MainTable)
        Next

        Dim totalHeight As Integer = 0
        For Each ctrl As Control In StatusPanel.Controls
            totalHeight += ctrl.Height
        Next
        If totalHeight > StatusPanel.Height Then
            StatusPanel.Height = totalHeight
        End If
    End Sub

    Private Sub GetAllPathInfo()
        Dim path As String
        For Each path In _pathList
            If path.Length > 0 Then
                bgWorker = New BackgroundWorker
                AddHandler bgWorker.DoWork, AddressOf GetPathInfo
                AddHandler bgWorker.RunWorkerCompleted, AddressOf RunWorkerCompleted
                bgWorker.RunWorkerAsync(path)
            End If
        Next
    End Sub


    Private Sub RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        If e.Result Is Nothing Then
            MessageBox.Show("e.Result is Nothing")
        End If

        Dim info As PathInfo = CType(e.Result, PathInfo)
        If info.ErrorMessage IsNot Nothing Then
            _statusRecords(info.Path).SetError(info.ErrorMessage, info.Status)
        Else
            _statusRecords(info.Path).SetFreeSpace(info.FreeBytes, info.FreePercent)
        End If

        'e.Result
    End Sub

    Private Sub GetPathInfo(ByVal sender As Object, ByVal e As DoWorkEventArgs) 'ByVal pathAsObject As Object)
        Dim path As String = e.Argument.ToString
        Dim drive As MappedDrive = Nothing
        Dim info As DriveInfo
        Try
            drive = New MappedDrive(path)
            info = My.Computer.FileSystem.GetDriveInfo(drive.DriveLetter.ToString)

            e.Result = New PathInfo(path, info.TotalFreeSpace, CInt(100 * info.AvailableFreeSpace / info.TotalSize))

            drive.Disconnect()
        Catch lex As OutOfLettersException
            e.Result = New PathInfo(path, ServerStatus.Unknown, "Timed out waiting for an available drive letter")
        Catch mapEx As MappingRejectedException
            e.Result = New PathInfo(path, ServerStatus.Bad, mapEx.Message)
        Catch ex As MappingUnknownException
            e.Result = New PathInfo(path, ServerStatus.Bad, ex.Message) 'Or could use ServerStatus.Unknown
        End Try
    End Sub

    'Corrects the width so that the status panel doesn't invoke the horizontal scroll bar
    Private Sub SetStatusPanelWidth()
        StatusPanel.Width = ScrollPanel.Width - SystemInformation.VerticalScrollBarWidth
    End Sub

    Private Sub ServerStatusForm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        For Each status As KeyValuePair(Of String, StatusRecord) In _statusRecords
            status.Value.Dispose()
        Next
    End Sub
End Class


Public Class PathInfo
    Public Path As String
    Public Status As ServerStatus
    Public FreeBytes As Long
    Public FreePercent As Integer
    Public ErrorMessage As String

    Public Sub New()
    End Sub

    Public Sub New(ByVal newPath As String, ByVal newFreeBytes As Long, ByVal newFreePercent As Integer)
        Path = newPath
        Status = ServerStatus.Good
        FreeBytes = newFreeBytes
        FreePercent = newFreePercent
        ErrorMessage = Nothing
    End Sub

    Public Sub New(ByVal newPath As String, ByVal newStatus As ServerStatus, ByVal newErrorMessage As String)
        Path = newPath
        Status = newStatus
        ErrorMessage = newErrorMessage
    End Sub
End Class


