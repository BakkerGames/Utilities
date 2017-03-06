' -----------------------------------------------
' --- DataConnection.TempData.vb - 04/20/2015 ---
' -----------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_TempData for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out TempData DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _TempDataDCC As ConnectionInfo

    Private WithEvents _CN_TempData As SqlConnection = Nothing
    Private _Transaction_TempData As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_TempData()
        _TempDataDCC = ArenaConfigInfo.GetConnectionInfo("TempData")
    End Sub

    Private Sub EndTransaction_TempData()
        If Not _Transaction_TempData Is Nothing Then
            _Transaction_TempData.Commit()
            _Transaction_TempData.Dispose()
            _Transaction_TempData = Nothing
        End If
    End Sub

    Private Sub CloseConnection_TempData()
        If Not _CN_TempData Is Nothing Then
            _CN_TempData.Close()
            _CN_TempData.Dispose()
            _CN_TempData = Nothing
        End If
    End Sub

    Private Sub Rollback_TempData()
        ' --- Roll back any TempData changes ---
        Try
            If Not _Transaction_TempData Is Nothing Then
                _Transaction_TempData.Rollback()
                _Transaction_TempData.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_TempData = Nothing
        Try
            If Not _CN_TempData Is Nothing Then
                _CN_TempData.Close()
                _CN_TempData.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_TempData = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_TempData() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _TempDataDCC Is Nothing OrElse _TempDataDCC.Server = "" OrElse _TempDataDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for TempData ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _TempDataDCC.Server
        Result.InitialCatalog = _TempDataDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_TempData() As SqlConnection
        If _InTransaction Then
            ' --- create a new TempData connection ---
            If _CN_TempData Is Nothing Then
                Try
                    _CN_TempData = New SqlConnection(ConnectionString_TempData)
                    _CN_TempData.Open()
                    _Transaction_TempData = _CN_TempData.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_TempData
        Else
            Dim Result As New SqlConnection(ConnectionString_TempData)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_TempData() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_TempData
    End Function

    Public Function ExecuteCommand_TempData(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnTempData As SqlClient.SqlConnection = GetConnection_TempData()
                Using cmd As New SqlCommand(SQLQuery, cnTempData)
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
