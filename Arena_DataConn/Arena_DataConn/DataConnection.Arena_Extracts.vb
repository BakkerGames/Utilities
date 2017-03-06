' -----------------------------------------------------
' --- DataConnection.Arena_Extracts.vb - 04/20/2015 ---
' -----------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_Arena_Extracts for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 03/09/2015 - SBakker
'            - Added Arena_Extracts DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _Arena_ExtractsDCC As ConnectionInfo

    Private WithEvents _CN_Arena_Extracts As SqlConnection = Nothing
    Private _Transaction_Arena_Extracts As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_Arena_Extracts()
        _Arena_ExtractsDCC = ArenaConfigInfo.GetConnectionInfo("Arena_Extracts")
    End Sub

    Private Sub EndTransaction_Arena_Extracts()
        If Not _Transaction_Arena_Extracts Is Nothing Then
            _Transaction_Arena_Extracts.Commit()
            _Transaction_Arena_Extracts.Dispose()
            _Transaction_Arena_Extracts = Nothing
        End If
    End Sub

    Private Sub CloseConnection_Arena_Extracts()
        If Not _CN_Arena_Extracts Is Nothing Then
            _CN_Arena_Extracts.Close()
            _CN_Arena_Extracts.Dispose()
            _CN_Arena_Extracts = Nothing
        End If
    End Sub

    Private Sub Rollback_Arena_Extracts()
        ' --- Roll back any Arena_Extracts changes ---
        Try
            If Not _Transaction_Arena_Extracts Is Nothing Then
                _Transaction_Arena_Extracts.Rollback()
                _Transaction_Arena_Extracts.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_Arena_Extracts = Nothing
        Try
            If Not _CN_Arena_Extracts Is Nothing Then
                _CN_Arena_Extracts.Close()
                _CN_Arena_Extracts.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_Arena_Extracts = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_Arena_Extracts() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _Arena_ExtractsDCC Is Nothing OrElse _Arena_ExtractsDCC.Server = "" OrElse _Arena_ExtractsDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for Arena_Extracts ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _Arena_ExtractsDCC.Server
        Result.InitialCatalog = _Arena_ExtractsDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_Arena_Extracts() As SqlConnection
        If _InTransaction Then
            ' --- create a new Arena_Extracts connection ---
            If _CN_Arena_Extracts Is Nothing Then
                Try
                    _CN_Arena_Extracts = New SqlConnection(ConnectionString_Arena_Extracts)
                    _CN_Arena_Extracts.Open()
                    _Transaction_Arena_Extracts = _CN_Arena_Extracts.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_Arena_Extracts
        Else
            Dim Result As New SqlConnection(ConnectionString_Arena_Extracts)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_Arena_Extracts() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_Arena_Extracts
    End Function

    Public Function ExecuteCommand_Arena_Extracts(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnArena_Extracts As SqlClient.SqlConnection = GetConnection_Arena_Extracts()
                Using cmd As New SqlCommand(SQLQuery, cnArena_Extracts)
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
