' ----------------------------------------------------
' --- DataConnection.Arena_Imports.vb - 02/27/2015 ---
' ----------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_Arena_Imports for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 02/27/2015 - SBakker
'            - Added Arena_Imports DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _Arena_ImportsDCC As ConnectionInfo

    Private WithEvents _CN_Arena_Imports As SqlConnection = Nothing
    Private _Transaction_Arena_Imports As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_Arena_Imports()
        _Arena_ImportsDCC = ArenaConfigInfo.GetConnectionInfo("Arena_Imports")
    End Sub

    Private Sub EndTransaction_Arena_Imports()
        If Not _Transaction_Arena_Imports Is Nothing Then
            _Transaction_Arena_Imports.Commit()
            _Transaction_Arena_Imports.Dispose()
            _Transaction_Arena_Imports = Nothing
        End If
    End Sub

    Private Sub CloseConnection_Arena_Imports()
        If Not _CN_Arena_Imports Is Nothing Then
            _CN_Arena_Imports.Close()
            _CN_Arena_Imports.Dispose()
            _CN_Arena_Imports = Nothing
        End If
    End Sub

    Private Sub Rollback_Arena_Imports()
        ' --- Roll back any Arena_Imports changes ---
        Try
            If Not _Transaction_Arena_Imports Is Nothing Then
                _Transaction_Arena_Imports.Rollback()
                _Transaction_Arena_Imports.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_Arena_Imports = Nothing
        Try
            If Not _CN_Arena_Imports Is Nothing Then
                _CN_Arena_Imports.Close()
                _CN_Arena_Imports.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_Arena_Imports = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_Arena_Imports() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _Arena_ImportsDCC Is Nothing OrElse _Arena_ImportsDCC.Server = "" OrElse _Arena_ImportsDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for Arena_Imports ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _Arena_ImportsDCC.Server
        Result.InitialCatalog = _Arena_ImportsDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_Arena_Imports() As SqlConnection
        If _InTransaction Then
            ' --- create a new Arena_Imports connection ---
            If _CN_Arena_Imports Is Nothing Then
                Try
                    _CN_Arena_Imports = New SqlConnection(ConnectionString_Arena_Imports)
                    _CN_Arena_Imports.Open()
                    _Transaction_Arena_Imports = _CN_Arena_Imports.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_Arena_Imports
        Else
            Dim Result As New SqlConnection(ConnectionString_Arena_Imports)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_Arena_Imports() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_Arena_Imports
    End Function

    Public Function ExecuteCommand_Arena_Imports(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnArena_Imports As SqlClient.SqlConnection = GetConnection_Arena_Imports()
                Using cmd As New SqlCommand(SQLQuery, cnArena_Imports)
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
