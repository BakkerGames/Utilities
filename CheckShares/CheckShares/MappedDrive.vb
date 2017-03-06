'Large portions are based on Eric Dalnas's publicly published code at http://www.mredkj.com/vbnet/vbnetmapdrive.html

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.ComponentModel

Public Class OutOfLettersException
    Inherits Exception
End Class

Public Class MappedDrive
    Protected _rwLock As ReaderWriterLock
    Protected _driveLetter As Nullable(Of Char)
    Protected _uncPath, _
              _username, _
              _password As String

    Private Shared AvailableLetters As New AvailableDriveLetters
    Private Const WAIT_LIMIT As Integer = 60000 'milliseconds

    Public ReadOnly Property DriveLetter() As Nullable(Of Char)
        Get
            Return _driveLetter
        End Get
    End Property
    Public ReadOnly Property UncPath() As String
        Get
            Return _uncPath
        End Get
    End Property

#Region "WinAPI References"
    Protected Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
        (ByRef lpNetResource As NetResource, ByVal lpPassword As String, _
         ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Protected Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" _
        (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

    <DllImport("Kernel32.dll", EntryPoint:="FormatMessageW", SetLastError:=True, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Protected Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, _
                                            ByRef lpBuffer As [String], ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer
    End Function

#End Region

    <StructLayout(LayoutKind.Sequential)> _
    Protected Structure NetResource
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String

        Public Sub New(ByVal localName As String, ByVal remoteName As String, ByVal type As Integer)
            lpLocalName = localName
            lpRemoteName = remoteName
            dwType = type
        End Sub
    End Structure

#Region "Constants"
    Protected Const FORCE_DISCONNECT As Integer = 1
    Protected Const RESOURCETYPE_DISK As Long = &H1

    Protected Const OK As Integer = 0
    Protected Const ERROR_ACCESS_DENIED As Integer = 5
    'Protected Const ERROR_DUPLICATE_PATH As Integer = 52 '"You were not connected because a duplicate name exists on the network. If joining a domain, go to System in Control Panel to change the computer name and try again. If joining a workgroup, choose another workgroup name."
    'Protected Const ERROR_BAD_NETPATH As Integer = 53
    'Protected Const ERROR_ALREADY_ASSIGNED As Integer = 85
    Protected Const ERROR_WRONG_TARGET_NAME As Integer = 1396
    Protected Const ERROR_NOT_CONNECTED As Integer = 2250

#End Region

    Public Sub New(ByVal uncPathToUse As String)
        _rwLock = New ReaderWriterLock
        _uncPath = uncPathToUse
        Dim timeWaiting As Integer = 0

        'Get the first available drive letter
        _rwLock.AcquireWriterLock(WAIT_LIMIT)
        Try
            'If needed, wait until a drive letter is available
            While AvailableLetters.Count < 1
                _rwLock.ReleaseWriterLock()

                If timeWaiting >= WAIT_LIMIT Then
                    Throw New OutOfLettersException()
                End If

                Thread.Sleep(500)
                timeWaiting += 500

                _rwLock.AcquireWriterLock(WAIT_LIMIT)
            End While

            _driveLetter = AvailableLetters.Pop

        Finally
            If _rwLock.IsWriterLockHeld Then
                _rwLock.ReleaseWriterLock()
            End If
        End Try

        If _driveLetter.HasValue Then
            Try
                MapDrive(_driveLetter.Value, _uncPath, Nothing, Nothing)
            Catch ex As Exception
                Disconnect()
                Throw
            End Try
        End If
    End Sub

    Private Sub MapDrive(ByVal driveLetterToUse As Char, ByVal uncPathToUse As String, _
                         ByVal usernameToUse As String, ByVal passwordToUse As String)

        Dim nr As New NetResource(DriveLetter & ":", UncPath, RESOURCETYPE_DISK)

        Dim result As Integer = WNetAddConnection2(nr, passwordToUse, usernameToUse, 0)

        Select Case result
            Case ERROR_ACCESS_DENIED, ERROR_WRONG_TARGET_NAME
                Throw New MappingRejectedException("Access denied")
                'Case ERROR_BAD_NETPATH
                '    Throw New MappingUnknownException("Network path not found")
                'Case ERROR_ALREADY_ASSIGNED
                '    Throw New MappingUnknownException("Drive letter """ & driveLetterToUse & ":"" is already in use.")
                'Case ERROR_DUPLICATE_PATH
                '    Throw New MappingUnknownException("Duplicate path exists on network")
            Case Is > 0
                Dim errorMessage As String = New Win32Exception(Err.LastDllError).Message
                Throw New MappingUnknownException(errorMessage)

                'Throw New MappingUnknownException(FormatMessage(result))
                'Throw New MappingUnknownException("Windows system error " & result & " occurred while mapping """ & uncPathToUse & """." & vbCrLf & _
                '                                  "You can look up that error number at http://msdn.microsoft.com/en-us/library/ms681381(VS.85).aspx")

        End Select

    End Sub

    Public Function Disconnect() As Boolean

        'Un-map the drive
        Dim returnCode As Integer = _
            WNetCancelConnection2(DriveLetter & ":", 0, FORCE_DISCONNECT)

        'Check if drive un-mapped successfully
        If returnCode = OK _
        OrElse returnCode = ERROR_NOT_CONNECTED Then
            'Drive was successfully disconnected (or not connected to begin with), so return this drive letter to the pool of available letters
            ReAddLetter(DriveLetter.Value)
            Return True
        End If

        'If another, unknown returnCode resulted, manually check if drive is still connected
        Dim info As DriveInfo = My.Computer.FileSystem.GetDriveInfo(DriveLetter.ToString & ":\")
        If info Is Nothing Then 'If drive is disconnected, so make letter available again
            ReAddLetter(DriveLetter.Value)
            Return True
        End If

        Return False
    End Function

    Private Sub ReAddLetter(ByVal letter As Char)
        _rwLock.AcquireWriterLock(WAIT_LIMIT)
        Try
            If Not AvailableLetters.Contains(DriveLetter.Value) Then
                AvailableLetters.Push(DriveLetter.Value)
            End If
        Finally
            _rwLock.ReleaseWriterLock()
        End Try
    End Sub


    ''' <summary>
    ''' Class to keep track of which drive letters are not in use.
    ''' </summary>
    ''' <remarks>To be used as a single-instance class, it must be declared Shared and instantiated in declaration, like this:
    ''' Private Shared AvailableLetters As New AvailableDriveLetters</remarks>
    Private NotInheritable Class AvailableDriveLetters
        Inherits Stack(Of Char)

        'Friend Shared Instance As New AvailableDriveLetters
        'Friend Available As Stack(Of Char)

        Public Shadows Sub Push(ByVal letter As Char)
            If Not Me.Contains(letter) Then 'Force uniqueness of each drive letter
                MyBase.Push(letter)
            End If
        End Sub

        Public Sub New()
            'Available = New Stack(Of Char)

            'Get a temporary list of all drive letters being used now
            Dim lettersInUse As New List(Of Char)
            For Each drive As DriveInfo In My.Computer.FileSystem.Drives
                lettersInUse.Add(drive.Name.ToUpper.Chars(0))
            Next

            'Create the list of available letters by looping through all letters and adding the ones that aren't in use now
            Dim letter As Char
            For i As Integer = 0 To 25 'For each of the 26 letters in the alphabet
                letter = Convert.ToChar(i + 65) 'In the ASCII character list, capital letters start at index 65
                If Not lettersInUse.Contains(letter) Then
                    Me.Push(letter)
                End If
            Next
        End Sub
    End Class
End Class



''' <summary>
''' Throw when the given path exists and clearly responded that the current user is denied access
''' </summary>
''' <remarks></remarks>
Public Class MappingRejectedException
    Inherits Exception

    Public Sub New(ByVal description As String)
        MyBase.New(description)
    End Sub
End Class

''' <summary>
''' Throw when we can't determine whether the current user may map the given path
''' </summary>
''' <remarks></remarks>
Public Class MappingUnknownException
    Inherits Exception

    Public Sub New(ByVal description As String)
        MyBase.New(description)
    End Sub
End Class
