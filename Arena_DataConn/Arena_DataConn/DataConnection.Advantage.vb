' ------------------------------------------------
' --- DataConnection.Advantage.vb - 04/20/2015 ---
' ------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_Advantage for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out Advantage DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _AdvantageDCC As ConnectionInfo

    Private WithEvents _CN_Advantage As SqlConnection = Nothing
    Private _Transaction_Advantage As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_Advantage()
        _AdvantageDCC = ArenaConfigInfo.GetConnectionInfo("Advantage")
    End Sub

    Private Sub EndTransaction_Advantage()
        If Not _Transaction_Advantage Is Nothing Then
            _Transaction_Advantage.Commit()
            _Transaction_Advantage.Dispose()
            _Transaction_Advantage = Nothing
        End If
    End Sub

    Private Sub CloseConnection_Advantage()
        If Not _CN_Advantage Is Nothing Then
            _CN_Advantage.Close()
            _CN_Advantage.Dispose()
            _CN_Advantage = Nothing
        End If
    End Sub

    Private Sub Rollback_Advantage()
        ' --- Roll back any Advantage changes ---
        Try
            If Not _Transaction_Advantage Is Nothing Then
                _Transaction_Advantage.Rollback()
                _Transaction_Advantage.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_Advantage = Nothing
        Try
            If Not _CN_Advantage Is Nothing Then
                _CN_Advantage.Close()
                _CN_Advantage.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_Advantage = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_Advantage() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _AdvantageDCC Is Nothing OrElse _AdvantageDCC.Server = "" OrElse _AdvantageDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for Advantage ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _AdvantageDCC.Server
        Result.InitialCatalog = _AdvantageDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_Advantage() As SqlConnection
        If _InTransaction Then
            ' --- create a new Advantage connection ---
            If _CN_Advantage Is Nothing Then
                Try
                    _CN_Advantage = New SqlConnection(ConnectionString_Advantage)
                    _CN_Advantage.Open()
                    _Transaction_Advantage = _CN_Advantage.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_Advantage
        Else
            Dim Result As New SqlConnection(ConnectionString_Advantage)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_Advantage() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_Advantage
    End Function

    Public Function ExecuteCommand_Advantage(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnAdvantage As SqlClient.SqlConnection = GetConnection_Advantage()
                Using cmd As New SqlCommand(SQLQuery, cnAdvantage)
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
