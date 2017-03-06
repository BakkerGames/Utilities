' -----------------------------------------------------
' --- DataConnection.IDRIS_Extracts.vb - 04/20/2015 ---
' -----------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_IDRIS_Extracts for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 08/22/2013 - SBakker
'            - Added IDRIS_Extracts DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _IDRIS_ExtractsDCC As ConnectionInfo

    Private WithEvents _CN_IDRIS_Extracts As SqlConnection = Nothing
    Private _Transaction_IDRIS_Extracts As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_IDRIS_Extracts()
        _IDRIS_ExtractsDCC = ArenaConfigInfo.GetConnectionInfo("IDRIS_Extracts")
    End Sub

    Private Sub EndTransaction_IDRIS_Extracts()
        If Not _Transaction_IDRIS_Extracts Is Nothing Then
            _Transaction_IDRIS_Extracts.Commit()
            _Transaction_IDRIS_Extracts.Dispose()
            _Transaction_IDRIS_Extracts = Nothing
        End If
    End Sub

    Private Sub CloseConnection_IDRIS_Extracts()
        If Not _CN_IDRIS_Extracts Is Nothing Then
            _CN_IDRIS_Extracts.Close()
            _CN_IDRIS_Extracts.Dispose()
            _CN_IDRIS_Extracts = Nothing
        End If
    End Sub

    Private Sub Rollback_IDRIS_Extracts()
        ' --- Roll back any IDRIS_Extracts changes ---
        Try
            If Not _Transaction_IDRIS_Extracts Is Nothing Then
                _Transaction_IDRIS_Extracts.Rollback()
                _Transaction_IDRIS_Extracts.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_IDRIS_Extracts = Nothing
        Try
            If Not _CN_IDRIS_Extracts Is Nothing Then
                _CN_IDRIS_Extracts.Close()
                _CN_IDRIS_Extracts.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_IDRIS_Extracts = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_IDRIS_Extracts() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _IDRIS_ExtractsDCC Is Nothing OrElse _IDRIS_ExtractsDCC.Server = "" OrElse _IDRIS_ExtractsDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for IDRIS_Extracts ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _IDRIS_ExtractsDCC.Server
        Result.InitialCatalog = _IDRIS_ExtractsDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_IDRIS_Extracts() As SqlConnection
        If _InTransaction Then
            ' --- create a new IDRIS_Extracts connection ---
            If _CN_IDRIS_Extracts Is Nothing Then
                Try
                    _CN_IDRIS_Extracts = New SqlConnection(ConnectionString_IDRIS_Extracts)
                    _CN_IDRIS_Extracts.Open()
                    _Transaction_IDRIS_Extracts = _CN_IDRIS_Extracts.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_IDRIS_Extracts
        Else
            Dim Result As New SqlConnection(ConnectionString_IDRIS_Extracts)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_IDRIS_Extracts() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_IDRIS_Extracts
    End Function

    Public Function ExecuteCommand_IDRIS_Extracts(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnIDRIS_Extracts As SqlClient.SqlConnection = GetConnection_IDRIS_Extracts()
                Using cmd As New SqlCommand(SQLQuery, cnIDRIS_Extracts)
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
