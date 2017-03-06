' ------------------------------------------
' --- DataConnection.IDR.vb - 04/20/2015 ---
' ------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_IDR for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out IDR DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _IDRDCC As ConnectionInfo

    Private WithEvents _CN_IDR As SqlConnection = Nothing
    Private _Transaction_IDR As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_IDR()
        _IDRDCC = ArenaConfigInfo.GetConnectionInfo("IDR")
    End Sub

    Private Sub EndTransaction_IDR()
        If Not _Transaction_IDR Is Nothing Then
            _Transaction_IDR.Commit()
            _Transaction_IDR.Dispose()
            _Transaction_IDR = Nothing
        End If
    End Sub

    Private Sub CloseConnection_IDR()
        If Not _CN_IDR Is Nothing Then
            _CN_IDR.Close()
            _CN_IDR.Dispose()
            _CN_IDR = Nothing
        End If
    End Sub

    Private Sub Rollback_IDR()
        ' --- Roll back any IDR changes ---
        Try
            If Not _Transaction_IDR Is Nothing Then
                _Transaction_IDR.Rollback()
                _Transaction_IDR.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_IDR = Nothing
        Try
            If Not _CN_IDR Is Nothing Then
                _CN_IDR.Close()
                _CN_IDR.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_IDR = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_IDR() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _IDRDCC Is Nothing OrElse _IDRDCC.Server = "" OrElse _IDRDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for IDR ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _IDRDCC.Server
        Result.InitialCatalog = _IDRDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_IDR() As SqlConnection
        If _InTransaction Then
            ' --- create a new IDR connection ---
            If _CN_IDR Is Nothing Then
                Try
                    _CN_IDR = New SqlConnection(ConnectionString_IDR)
                    _CN_IDR.Open()
                    _Transaction_IDR = _CN_IDR.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_IDR
        Else
            Dim Result As New SqlConnection(ConnectionString_IDR)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_IDR() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_IDR
    End Function

    Public Function ExecuteCommand_IDR(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnIDR As SqlClient.SqlConnection = GetConnection_IDR()
                Using cmd As New SqlCommand(SQLQuery, cnIDR)
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
