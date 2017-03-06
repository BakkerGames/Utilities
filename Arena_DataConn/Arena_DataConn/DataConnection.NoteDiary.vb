' ------------------------------------------------
' --- DataConnection.NoteDiary.vb - 04/20/2015 ---
' ------------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 04/20/2015 - SBakker
'            - Added routine ExecuteCommand_NoteDiary for simpler starting of Stored Procedures or
'              direct Add/Update/Delete queries. Note: the SQLQuery is NOT sent through StringToSQL!
' 05/31/2013 - SBakker
'            - Separated out NoteDiary DataConnection information.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Partial Public Class DataConnection

#Region " Private Variables "

    Private _NoteDiaryDCC As ConnectionInfo

    Private WithEvents _CN_NoteDiary As SqlConnection = Nothing
    Private _Transaction_NoteDiary As SqlTransaction = Nothing

#End Region

#Region " Private Routines "

    Private Sub LoadSettings_NoteDiary()
        _NoteDiaryDCC = ArenaConfigInfo.GetConnectionInfo("NoteDiary")
    End Sub

    Private Sub EndTransaction_NoteDiary()
        If Not _Transaction_NoteDiary Is Nothing Then
            _Transaction_NoteDiary.Commit()
            _Transaction_NoteDiary.Dispose()
            _Transaction_NoteDiary = Nothing
        End If
    End Sub

    Private Sub CloseConnection_NoteDiary()
        If Not _CN_NoteDiary Is Nothing Then
            _CN_NoteDiary.Close()
            _CN_NoteDiary.Dispose()
            _CN_NoteDiary = Nothing
        End If
    End Sub

    Private Sub Rollback_NoteDiary()
        ' --- Roll back any NoteDiary changes ---
        Try
            If Not _Transaction_NoteDiary Is Nothing Then
                _Transaction_NoteDiary.Rollback()
                _Transaction_NoteDiary.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _Transaction_NoteDiary = Nothing
        Try
            If Not _CN_NoteDiary Is Nothing Then
                _CN_NoteDiary.Close()
                _CN_NoteDiary.Dispose()
            End If
        Catch ex As Exception
            ' --- No errors should be thrown in Rollback ---
        End Try
        _CN_NoteDiary = Nothing
    End Sub

#End Region

#Region " Public Routines "

    Public Function ConnectionString_NoteDiary() As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If _NoteDiaryDCC Is Nothing OrElse _NoteDiaryDCC.Server = "" OrElse _NoteDiaryDCC.Database = "" Then
            Throw New SystemException(FuncName + vbCrLf + "Server/Database not specified in configuration file")
        End If
        ' --- return the connection string for NoteDiary ---
        Dim Result As New SqlConnectionStringBuilder
        Result.IntegratedSecurity = True
        Result.PersistSecurityInfo = False
        Result.DataSource = _NoteDiaryDCC.Server
        Result.InitialCatalog = _NoteDiaryDCC.Database
        Result.Encrypt = False
        Result.ConnectTimeout = SQLTimeoutSeconds
        Return Result.ToString
    End Function

    Public Function GetConnection_NoteDiary() As SqlConnection
        If _InTransaction Then
            ' --- create a new NoteDiary connection ---
            If _CN_NoteDiary Is Nothing Then
                Try
                    _CN_NoteDiary = New SqlConnection(ConnectionString_NoteDiary)
                    _CN_NoteDiary.Open()
                    _Transaction_NoteDiary = _CN_NoteDiary.BeginTransaction()
                Catch ex As Exception
                    Me.Rollback()
                    Throw ' re-throw the exception
                End Try
            End If
            ' --- return the connection ---
            Return _CN_NoteDiary
        Else
            Dim Result As New SqlConnection(ConnectionString_NoteDiary)
            Result.Open()
            Return Result
        End If
    End Function

    Public Function GetTransaction_NoteDiary() As SqlTransaction
        ' --- This is needed by a SQLCommand object when running inside a transaction ---
        Return _Transaction_NoteDiary
    End Function

    Public Function ExecuteCommand_NoteDiary(ByVal SQLQuery As String) As Integer
        Dim Result As Integer
        ' -------------------
        Try
            Using cnNoteDiary As SqlClient.SqlConnection = GetConnection_NoteDiary()
                Using cmd As New SqlCommand(SQLQuery, cnNoteDiary)
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
