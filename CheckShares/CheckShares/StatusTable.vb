Imports System.Math
Imports System.ComponentModel

Public Enum ServerStatus
    Good
    Bad
    Unknown
End Enum

Public Class StatusRecord
    Implements IDisposable

    Private _embeddedResources As EmbeddedResources
    Private _freeSpaceTable As TableLayoutPanel

    Private _serverLabel, _
            _freeSpaceLabel, _
            _errorMessageLabel As Label

    Private _barGraph, _
            _barGraphBar As Panel

    Private _statusIcon As PictureBox
    Private _parentForm As Form

    Public MainTable As TableLayoutPanel

    Public Sub New(ByVal path As String)
        _embeddedResources = New EmbeddedResources

        'MainTable = statusTable
        MainTable = New TableLayoutPanel
        MainTable.Width = 598
        MainTable.Height = 20
        MainTable.Margin = New Padding(0)
        MainTable.ColumnCount = 3
        MainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 20))
        MainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
        MainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 400))
        'MainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 200))
        MainTable.RowCount = 1
        MainTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 20))

        '----- Status icon -----
        _statusIcon = New PictureBox
        _statusIcon.Margin = New Padding(3)
        _statusIcon.Image = _embeddedResources.Bitmaps("grey.png")
        'StatusToolTip.SetToolTip(_statusIcon, "Getting status...")
        MainTable.Controls.Add(_statusIcon, 0, 0)

        '----- Server name -----
        _serverLabel = New Label()
        _serverLabel.Text = path
        _serverLabel.Margin = New Padding(3)
        _serverLabel.AutoSize = False
        _serverLabel.AutoEllipsis = True
        _serverLabel.Width = CInt(MainTable.ColumnStyles(1).Width _
                                 - _serverLabel.Margin.Left - _serverLabel.Margin.Right)
        MainTable.Controls.Add(_serverLabel, 1, 0)

        MainTable.ResumeLayout()
        Application.DoEvents()
    End Sub

    Public Sub SetError(ByVal errorDescription As String, ByVal status As ServerStatus)

        If _errorMessageLabel Is Nothing Then
            _errorMessageLabel = New Label
            _errorMessageLabel.AutoSize = True
            _errorMessageLabel.MaximumSize = New Size(400, 20)
            _errorMessageLabel.Padding = New Padding(0)
            _errorMessageLabel.Margin = New Padding(0)
            _errorMessageLabel.AutoEllipsis = True
            _errorMessageLabel.ForeColor = Color.DarkRed
            _errorMessageLabel.TextAlign = ContentAlignment.TopLeft
            _errorMessageLabel.Width = CInt(MainTable.ColumnStyles(2).Width _
                                            - _errorMessageLabel.Margin.Left - _errorMessageLabel.Margin.Right)


            MainTable.Controls.Add(_errorMessageLabel, 2, 0)

            Select Case status
                Case ServerStatus.Bad
                    Dim redImage As Image = _embeddedResources.Bitmaps("red.png")
                    _statusIcon.Image = redImage
                Case ServerStatus.Unknown
                    Dim greyImage As Image = _embeddedResources.Bitmaps("grey.png")
                    _statusIcon.Image = greyImage
            End Select
        End If
        _errorMessageLabel.Text = errorDescription

        Application.DoEvents()
    End Sub


    Public Sub SetFreeSpace(ByVal freeSpace As Long, ByVal percentFree As Integer)

        Dim greenImage As Image = _embeddedResources.Bitmaps("green.png")
        _statusIcon.Image = greenImage

        If _freeSpaceTable Is Nothing Then
            _freeSpaceTable = New TableLayoutPanel
            _freeSpaceTable.ColumnCount = 2
            _freeSpaceTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 54))
            _freeSpaceTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            _freeSpaceTable.RowCount = 1
            _freeSpaceTable.Margin = New Padding(0)
            _freeSpaceTable.Padding = New Padding(0)
        End If

        '----- Bar graph background -----
        If _barGraph Is Nothing Then
            _barGraph = New Panel
            'barGraph.BackColor = Color.SteelBlue
            _barGraph.Width = 54
            _barGraph.Height = 14
            _barGraph.Top = 0
            _barGraph.Left = 0
            _barGraph.Margin = New Padding(0)
            _barGraph.BorderStyle = BorderStyle.Fixed3D
            _barGraph.BackgroundImageLayout = ImageLayout.Stretch
            _barGraph.BackgroundImage = _embeddedResources.Bitmaps("progress-free.png")
            'barGraph.BorderStyle = BorderStyle.FixedSingle
        End If

        '----- Bar graph data -----
        If _barGraphBar Is Nothing Then
            _barGraphBar = New Panel
            'barGraphBar.BackColor = Color.LightSteelBlue
            _barGraphBar.Height = 16
            _barGraphBar.Top = 0
            _barGraphBar.Left = 0
            _barGraph.Margin = New Padding(0)
            _barGraphBar.BackgroundImageLayout = ImageLayout.Stretch
            _barGraphBar.BackgroundImage = _embeddedResources.Bitmaps("progress-used.png")
            _barGraph.Controls.Add(_barGraphBar)
            _freeSpaceTable.Controls.Add(_barGraph, 0, 0)
        End If
        _barGraphBar.Width = CInt((100 - percentFree) / 2)


        '----- Free space description -----
        If _freeSpaceLabel Is Nothing Then
            _freeSpaceLabel = New Label
            _freeSpaceLabel.Margin = New Padding(0)
            _freeSpaceLabel.TextAlign = ContentAlignment.TopLeft
            _freeSpaceTable.Controls.Add(_freeSpaceLabel, 1, 0)
        End If
        _freeSpaceLabel.Text = String.Format("{0} free ({1}%)", GetSizeDescription(freeSpace), percentFree)

        MainTable.Controls.Add(_freeSpaceTable, 2, 0)
    End Sub


    Private Shared Function GetSizeDescription(ByVal size As Long) As String
        Const PB As Long = 1125899906842624
        Const TB As Long = 1099511627776
        Const GB As Integer = 1073741824
        Const MB As Integer = 1048576
        Const KB As Integer = 1024

        Select Case size
            Case Is > PB
                Return String.Format("{0} PB", Round(size / PB, 1))
            Case Is > TB
                Return String.Format("{0} TB", Round(size / TB, 1))
            Case Is > GB
                Return String.Format("{0} GB", Round(size / GB, 1))
            Case Is > MB
                Return String.Format("{0} MB", Round(size / MB, 1))
            Case Is > KB
                Return String.Format("{0} KB", Round(size / KB, 1))
            Case Else
                Return String.Format("{0} Bytes", size)
        End Select

    End Function

#Region " IDisposable Support "
    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                _embeddedResources.Dispose()
            End If
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
