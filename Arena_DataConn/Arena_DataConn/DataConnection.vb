' --------------------------------------
' --- DataConnection.vb - 10/27/2016 ---
' --------------------------------------

' ----------------------------------------------------------------------------------------------------
' 10/27/2016 - SBakker
'            - Moved timeout information into Arena_ConfigInfo, so it can be externally adjusted.
'            - Made Environment, SQLTimeoutSeconds, MaxSQLRetryCount readonly properties.
' 10/09/2015 - SBakker
'            - Replaced property Timeout with constant SQLTimeoutSeconds.
' 03/09/2015 - SBakker
'            - Added connection information for the Arena_Extracts database.
' 02/27/2015 - SBakker
'            - Added connection information for the Arena_Imports database.
' 04/24/2014 - SBakker
'            - Implement IDisposable so that a Rollback will occur if this object is destroyed outside
'              of the normal program handling.
' 01/08/2014 - SBakker
'            - Allow multiple classes to be part of the same transaction by just incrementing and
'              decrementing the new _TransactionLevel counter. This allows a class to add/edit/delete
'              other records, but not need to keep track if the calling program started a transaction
'              or not. They would just call dc.BeginTransaction and dc.EndTransaction with no other
'              checking neccessary.
' 08/22/2013 - SBakker
'            - Added connection information for the IDRIS_Extracts database.
' 05/31/2013 - SBakker
'            - Added connection information for the NoteDiary database.
'            - Split all database-specific info into separate partial classes for easier
'              management. The code was getting too large in this single class file.
' 03/25/2013 - SBakker
'            - Added connection information for the DN_DataStore database.
' 04/30/2012 - SBakker
'            - Added connection information for the UserRequest database.
' 04/13/2012 - SBakker
'            - Added connection information for the IDR database.
'            - Fixed bug in CloseConnection, where it was closing Security twice, instead of
'              Security once and TempData once.
' 03/18/2011 - SBakker
'            - Added CloseConnection() so that connections could be property closed and
'              disposed of, unless currently in a transaction.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 09/03/2009 - SBakker
'            - Use new Arena_ConfigInfo to read the Arena.xml file, rather than
'              duplicating the code to do it everywhere.
'            - Changed TimeOut to be Integer instead of UInteger.
' 07/22/2009 - SBakker
'            - Added TempData connections for any data which is transitory and
'              doesn't need to be backed up.
' 01/29/2009 - SBakker
'            - Added a routine TestConnections to see if all servers are up and
'              working. Usually would be used by an executable just as it is
'              starting, not by a lower data class.
' 01/28/2009 - SBakker
'            - Added a Timeout property to the DataConnection object, so it can
'              be set from a calling program if needed.
'            - Increased the default timeout from 15 to 60 seconds.
' 12/31/2008 - SBakker - Arena
'            - Changed My.Settings to only hold the path to the new "Arena.xml"
'              file. This file will contain all the Server and Database info
'              needed for connection strings, and for the environment name. Any
'              other info in the file is ignored, so it can hold anything else.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports System.Data.SqlClient

Public Class DataConnection

    Implements IDisposable

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Public Sub New()
        LoadSettings()
    End Sub

    Public Sub LoadSettings()
        LoadSettings_Advantage()
        LoadSettings_Arena()
        LoadSettings_Arena_Extracts()
        LoadSettings_Arena_Imports()
        LoadSettings_DN_DataStore()
        LoadSettings_IDR()
        LoadSettings_IDRIS()
        LoadSettings_IDRIS_Extracts()
        LoadSettings_NoteDiary()
        LoadSettings_Security()
        LoadSettings_TempData()
        LoadSettings_UserRequest()
    End Sub

    Public Shared ReadOnly Property Environment As String
        Get
            Return ArenaConfigInfo.Environment
        End Get
    End Property

    Public Shared ReadOnly Property SQLTimeoutSeconds As Integer
        Get
            Return ArenaConfigInfo.TimeoutSeconds
        End Get
    End Property

    Public Shared ReadOnly Property MaxSQLRetryCount As Integer
        Get
            Return ArenaConfigInfo.TimeoutRetry
        End Get
    End Property

    ' ---------------------------------------
    ' --- Transaction Processing Routines ---
    ' ---------------------------------------

    Private _InTransaction As Boolean = False
    Public ReadOnly Property InTransaction() As Boolean
        Get
            Return _InTransaction
        End Get
    End Property

    Private _TransactionLevel As Integer = 0

    Public Sub BeginTransaction()
        ' --- Use this if multiple objects must be saved all-or-nothing to preserve database integrity ---
        ' --- Allow multiple classes to all be part of the same transaction by just calling BeginTransaction ---
        _TransactionLevel += 1
        If _InTransaction Then
            Exit Sub
        End If
        _InTransaction = True
    End Sub

    Public Sub EndTransaction()
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Check for programming errors ---
        If Not _InTransaction Then
            Throw New SystemException(FuncName + vbCrLf + "Not currently in a transaction.")
        End If
        ' --- Allow multiple classes to all be part of the same transaction by just calling EndTransaction ---
        _TransactionLevel -= 1
        If _TransactionLevel > 0 Then
            Exit Sub
        End If
        ' --- Commit the transaction ---
        Try
            ' -- Commit Transactions ---
            EndTransaction_Advantage()
            EndTransaction_Arena()
            EndTransaction_Arena_Extracts()
            EndTransaction_Arena_Imports()
            EndTransaction_DN_DataStore()
            EndTransaction_IDR()
            EndTransaction_IDRIS()
            EndTransaction_IDRIS_Extracts()
            EndTransaction_NoteDiary()
            EndTransaction_Security()
            EndTransaction_TempData()
            EndTransaction_UserRequest()
            ' --- Close Connections ---
            CloseConnection_Advantage()
            CloseConnection_Arena()
            CloseConnection_Arena_Extracts()
            CloseConnection_Arena_Imports()
            CloseConnection_DN_DataStore()
            CloseConnection_IDR()
            CloseConnection_IDRIS()
            CloseConnection_IDRIS_Extracts()
            CloseConnection_NoteDiary()
            CloseConnection_Security()
            CloseConnection_TempData()
            CloseConnection_UserRequest()
        Catch ex As Exception
            Me.Rollback()
            Throw ' Re-throw the exception to the calling program
        End Try
        _TransactionLevel = 0
        _InTransaction = False
    End Sub

    Public Sub Rollback()
        Rollback_Advantage()
        Rollback_Arena()
        Rollback_Arena_Extracts()
        Rollback_Arena_Imports()
        Rollback_DN_DataStore()
        Rollback_IDR()
        Rollback_IDRIS()
        Rollback_IDRIS_Extracts()
        Rollback_NoteDiary()
        Rollback_Security()
        Rollback_TempData()
        Rollback_UserRequest()
        ' --- Done ---
        _TransactionLevel = 0
        _InTransaction = False
    End Sub

    ' --------------------------------
    ' --- CloseConnection Routines ---
    ' --------------------------------

    Public Sub CloseConnection(ByRef CurrConn As SqlConnection)
        If CurrConn Is Nothing Then Exit Sub
        If _InTransaction Then Exit Sub
        CurrConn.Close()
        CurrConn.Dispose()
        CurrConn = Nothing
    End Sub

#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If _InTransaction Then
                    Rollback()
                End If
            End If
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
