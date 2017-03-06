' -----------------------------------------------
' --- DataConnection.Security.vb - 04/20/2015 ---
' -----------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_Security for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out Security DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _SecurityDCC As ConnectionInfo

    Private WithEvents _CN_Security As SqlConnection = Nothing
    Private _Transaction_Security As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_Security()
        _SecurityDCC = ArenaConfigInfo.GetConnectionInfo("Security")
    End Sub

    Private Sub EndTransaction_Security()
        If Not _Transaction_Security Is Nothing Then
            _Transaction_Security.Commit()
            _Transaction_Security.Dispose()
            _Transaction_Security = Nothing
        End If
    End Sub

    Private Sub CloseConnection_Security()
        If Not _CN_Security Is Nothing Then
            _CN_Security.Close()
            _CN_Security.Dispose()
            _CN_Security = Nothing
        End If
    End Sub

    Private Sub Rollback_Security()
        ' --- Roll back any Security changes ---
        Try
            If Not _Transaction_Security Is Nothing Then
                _Transaction_Security.Rollback()
                _Transaction_Security.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_Security = Nothing
        Try
            If Not _CN_Security Is Nothing Then
                _CN_Security.Close()
                _CN_Security.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_Security = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_Security() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _SecurityDCC Is Nothing OrElse _SecurityDCC.Server = "" OrElse _SecurityDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for Security ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _SecurityDCC.Server
        Result.InitialCatalog = _SecurityDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_Security() As SqlConnection
        If _InTransaction Then
            ' --- create a new Security connection ---
            If _CN_Security Is Nothing Then
                Try
                    _CN_Security = New SqlConnection(ConnectionString_Security)
                    _CN_Security.Open()
                    _Transaction_Security = _CN_Security.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_Security
        Else
            Dim Result As New SqlConnection(ConnectionString_Security)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_Security() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_Security
    End Function

    Public Function ExecuteCommand_Security(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnSecurity As SqlClient.SqlConnection = GetConnection_Security()
                Using cmd As New SqlCommand(SQLQuery, cnSecurity)
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
