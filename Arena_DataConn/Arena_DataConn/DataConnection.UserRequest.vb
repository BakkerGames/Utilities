' --------------------------------------------------
' --- DataConnection.UserRequest.vb - 04/20/2015 ---
' --------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_UserRequest for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out UserRequest DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _UserRequestDCC As ConnectionInfo

    Private WithEvents _CN_UserRequest As SqlConnection = Nothing
    Private _Transaction_UserRequest As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_UserRequest()
        _UserRequestDCC = ArenaConfigInfo.GetConnectionInfo("UserRequest")
    End Sub

    Private Sub EndTransaction_UserRequest()
        If Not _Transaction_UserRequest Is Nothing Then
            _Transaction_UserRequest.Commit()
            _Transaction_UserRequest.Dispose()
            _Transaction_UserRequest = Nothing
        End If
    End Sub

    Private Sub CloseConnection_UserRequest()
        If Not _CN_UserRequest Is Nothing Then
            _CN_UserRequest.Close()
            _CN_UserRequest.Dispose()
            _CN_UserRequest = Nothing
        End If
    End Sub

    Private Sub Rollback_UserRequest()
        ' --- Roll back any UserRequest changes ---
        Try
            If Not _Transaction_UserRequest Is Nothing Then
                _Transaction_UserRequest.Rollback()
                _Transaction_UserRequest.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_UserRequest = Nothing
        Try
            If Not _CN_UserRequest Is Nothing Then
                _CN_UserRequest.Close()
                _CN_UserRequest.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_UserRequest = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_UserRequest() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _UserRequestDCC Is Nothing OrElse _UserRequestDCC.Server = "" OrElse _UserRequestDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for UserRequest ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _UserRequestDCC.Server
        Result.InitialCatalog = _UserRequestDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_UserRequest() As SqlConnection
        If _InTransaction Then
            ' --- create a new UserRequest connection ---
            If _CN_UserRequest Is Nothing Then
                Try
                    _CN_UserRequest = New SqlConnection(ConnectionString_UserRequest)
                    _CN_UserRequest.Open()
                    _Transaction_UserRequest = _CN_UserRequest.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_UserRequest
        Else
            Dim Result As New SqlConnection(ConnectionString_UserRequest)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_UserRequest() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_UserRequest
    End Function

    Public Function ExecuteCommand_UserRequest(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnUserRequest As SqlClient.SqlConnection = GetConnection_UserRequest()
                Using cmd As New SqlCommand(SQLQuery, cnUserRequest)
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
