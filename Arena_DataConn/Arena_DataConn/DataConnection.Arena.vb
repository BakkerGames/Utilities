' --------------------------------------------
' --- DataConnection.Arena.vb - 04/20/2015 ---
' --------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_Arena for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out Arena DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _ArenaDCC As ConnectionInfo

    Private WithEvents _CN_Arena As SqlConnection = Nothing
    Private _Transaction_Arena As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_Arena()
        _ArenaDCC = ArenaConfigInfo.GetConnectionInfo("Arena")
    End Sub

    Private Sub EndTransaction_Arena()
        If Not _Transaction_Arena Is Nothing Then
            _Transaction_Arena.Commit()
            _Transaction_Arena.Dispose()
            _Transaction_Arena = Nothing
        End If
    End Sub

    Private Sub CloseConnection_Arena()
        If Not _CN_Arena Is Nothing Then
            _CN_Arena.Close()
            _CN_Arena.Dispose()
            _CN_Arena = Nothing
        End If
    End Sub

    Private Sub Rollback_Arena()
        ' --- Roll back any Arena changes ---
        Try
            If Not _Transaction_Arena Is Nothing Then
                _Transaction_Arena.Rollback()
                _Transaction_Arena.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_Arena = Nothing
        Try
            If Not _CN_Arena Is Nothing Then
                _CN_Arena.Close()
                _CN_Arena.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_Arena = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_Arena() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _ArenaDCC Is Nothing OrElse _ArenaDCC.Server = "" OrElse _ArenaDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for Arena ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _ArenaDCC.Server
        Result.InitialCatalog = _ArenaDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_Arena() As SqlConnection
        If _InTransaction Then
            ' --- create a new Arena connection ---
            If _CN_Arena Is Nothing Then
                Try
                    _CN_Arena = New SqlConnection(ConnectionString_Arena)
                    _CN_Arena.Open()
                    _Transaction_Arena = _CN_Arena.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_Arena
        Else
            Dim Result As New SqlConnection(ConnectionString_Arena)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_Arena() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_Arena
    End Function

    Public Function ExecuteCommand_Arena(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnArena As SqlClient.SqlConnection = GetConnection_Arena()
                Using cmd As New SqlCommand(SQLQuery, cnArena)
                    cmd.CommandType = CommandType.Text
                    Result = cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            Throw ' re-throw the exception
        End Try
        Return Result
    End Function

#End Region

End Class
