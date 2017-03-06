' ---------------------------------------------------
' --- DataConnection.DN_DataStore.vb - 04/20/2015 ---
' ---------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_DN_DataStore for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out DN_DataStore DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _DN_DataStoreDCC As ConnectionInfo

    Private WithEvents _CN_DN_DataStore As SqlConnection = Nothing
    Private _Transaction_DN_DataStore As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_DN_DataStore()
        _DN_DataStoreDCC = ArenaConfigInfo.GetConnectionInfo("DN_DataStore")
    End Sub

    Private Sub EndTransaction_DN_DataStore()
        If Not _Transaction_DN_DataStore Is Nothing Then
            _Transaction_DN_DataStore.Commit()
            _Transaction_DN_DataStore.Dispose()
            _Transaction_DN_DataStore = Nothing
        End If
    End Sub

    Private Sub CloseConnection_DN_DataStore()
        If Not _CN_DN_DataStore Is Nothing Then
            _CN_DN_DataStore.Close()
            _CN_DN_DataStore.Dispose()
            _CN_DN_DataStore = Nothing
        End If
    End Sub

    Private Sub Rollback_DN_DataStore()
        ' --- Roll back any DN_DataStore changes ---
        Try
            If Not _Transaction_DN_DataStore Is Nothing Then
                _Transaction_DN_DataStore.Rollback()
                _Transaction_DN_DataStore.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_DN_DataStore = Nothing
        Try
            If Not _CN_DN_DataStore Is Nothing Then
                _CN_DN_DataStore.Close()
                _CN_DN_DataStore.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_DN_DataStore = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_DN_DataStore() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _DN_DataStoreDCC Is Nothing OrElse _DN_DataStoreDCC.Server = "" OrElse _DN_DataStoreDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for DN_DataStore ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _DN_DataStoreDCC.Server
        Result.InitialCatalog = _DN_DataStoreDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_DN_DataStore() As SqlConnection
        If _InTransaction Then
            ' --- create a new DN_DataStore connection ---
            If _CN_DN_DataStore Is Nothing Then
                Try
                    _CN_DN_DataStore = New SqlConnection(ConnectionString_DN_DataStore)
                    _CN_DN_DataStore.Open()
                    _Transaction_DN_DataStore = _CN_DN_DataStore.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_DN_DataStore
        Else
            Dim Result As New SqlConnection(ConnectionString_DN_DataStore)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_DN_DataStore() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_DN_DataStore
    End Function

    Public Function ExecuteCommand_DN_DataStore(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnDN_DataStore As SqlClient.SqlConnection = GetConnection_DN_DataStore()
                Using cmd As New SqlCommand(SQLQuery, cnDN_DataStore)
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
