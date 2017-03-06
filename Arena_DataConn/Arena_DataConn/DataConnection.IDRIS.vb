' --------------------------------------------
' --- DataConnection.IDRIS.vb - 04/20/2015 ---
' --------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_IDRIS for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out IDRIS DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _IDRISDCC As ConnectionInfo

    Private WithEvents _CN_IDRIS As SqlConnection = Nothing
    Private _Transaction_IDRIS As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_IDRIS()
        _IDRISDCC = ArenaConfigInfo.GetConnectionInfo("IDRIS")
    End Sub

    Private Sub EndTransaction_IDRIS()
        If Not _Transaction_IDRIS Is Nothing Then
            _Transaction_IDRIS.Commit()
            _Transaction_IDRIS.Dispose()
            _Transaction_IDRIS = Nothing
        End If
    End Sub

    Private Sub CloseConnection_IDRIS()
        If Not _CN_IDRIS Is Nothing Then
            _CN_IDRIS.Close()
            _CN_IDRIS.Dispose()
            _CN_IDRIS = Nothing
        End If
    End Sub

    Private Sub Rollback_IDRIS()
        ' --- Roll back any IDRIS changes ---
        Try
            If Not _Transaction_IDRIS Is Nothing Then
                _Transaction_IDRIS.Rollback()
                _Transaction_IDRIS.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_IDRIS = Nothing
        Try
            If Not _CN_IDRIS Is Nothing Then
                _CN_IDRIS.Close()
                _CN_IDRIS.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_IDRIS = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_IDRIS() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _IDRISDCC Is Nothing OrElse _IDRISDCC.Server = "" OrElse _IDRISDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for IDRIS ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _IDRISDCC.Server
        Result.InitialCatalog = _IDRISDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_IDRIS() As SqlConnection
        If _InTransaction Then
            ' --- create a new IDRIS connection ---
            If _CN_IDRIS Is Nothing Then
                Try
                    _CN_IDRIS = New SqlConnection(ConnectionString_IDRIS)
                    _CN_IDRIS.Open()
                    _Transaction_IDRIS = _CN_IDRIS.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_IDRIS
        Else
            Dim Result As New SqlConnection(ConnectionString_IDRIS)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_IDRIS() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_IDRIS
    End Function

    Public Function ExecuteCommand_IDRIS(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnIDRIS As SqlClient.SqlConnection = GetConnection_IDRIS()
                Using cmd As New SqlCommand(SQLQuery, cnIDRIS)
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
