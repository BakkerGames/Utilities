' -----------------------------------
' --- SystemUtils.vb - 07/22/2016 ---
' -----------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/22/2016 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrEmpty().
' 06/30/2016 - SBakker
'            - Changed NormalizePath() to use hard-coded settings for UNC paths instead of calling
'              "NET USE". Neither is good, but "NET USE" could have issues.
'            - Fixed NormalizePath() to always end with "\". Also handle Nothing/Empty better.
'            - Make sure that there are no text case issues in NormalizePath().
' 06/29/2015 - SBakker - URD 12548
'            - Added error checking in NormalizePath() to return the original path if errors occur.
'            - Lowercase the results in NormalizePath() to avoid text case issues.
' 04/29/2015 - SBakker
'            - Added CMDAutomate() to execute a command in a CMD window and return the results. Used
'              in NormalizePath().
'            - Added NormalizePath() to return a path normalized, stripped of trailing slashes, and
'              converted to UNC if applicable. The text case of the result is dependent on the actual
'              folder/server/share names on the system and should not be assumed.
' 09/30/2014 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrWhiteSpace().
' 03/28/2014 - SBakker
'            - Don't try to read from Active Directory if Arena_ConfigInfo says it isn't available.
' 09/17/2013 - SBakker
'            - Added additional error information.
' 01/22/2013 - SBakker
'            - Fix logic for InEOMPeriod to test for drifting into or out of EOM period.
' 11/20/2012 - SBakker
'            - Added use of IDRIS [%EOMDates] and [%EOMUsers] in IsAfterMonthEndCutoff to
'              get the same answer in Arena as in IDRIS.
' 05/22/2012 - SBakker - Bug 2-85
'            - Added IsAfterMonthEndCutoff function and the settings MonthEndCutoffHour and
'              MonthEndCutoffMinute, for checking against the current Month End Cutoff time.
'              Background tasks will be allowed to always run, however.
' 01/25/2011 - SBakker
'            - Added more error handling for GetUserName, throwing an error if everything
'              fails. Might help with SQL jobs in the background.
' 09/30/2011 - SBakker
'            - Added version of GetUserName which can be used to inpersonate a test user. To
'              use, call GetUserName("zzztest") at the very beginning of a frontend program.
'              It will then remember that name for the life of that one application.
'              *** ONLY WORKS WITH NAMES ENDING IN "test" FOR SECURITY REASONS!!! ***
' 11/15/2010 - SBakker
'            - Added comments to each function so it describes exactly what it returns.
'            - Made all functions in this module use static variables to hold the answers.
'              It's unlikely these answers could change during the running of a program.
' 04/14/2010 - SBakker
'            - Added error handling around trying to access Active Directory info.
' 11/03/2009 - SBakker
'            - Added GetUserFullName, for showing the "Firstname Lastname" of
'              the user currently logged in. Since this takes about 1-2 seconds
'              to get the answer, it will only ask Active Directory once and
'              store the result for any subsequent calls.
' 09/03/2009 - SBakker
'            - Added GetComputerName, for standardized use in other routines.
' 12/05/2008 - SBAKKER - Arena
'            - Added GetUserName() so it can be standardized in this one place.
'              Will be used anywhere the current username is required.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo
Imports Arena_Utilities.DateUtils
Imports Arena_Utilities.StringUtils
Imports System.Data
Imports System.Data.SqlClient
Imports System.DirectoryServices
Imports System.IO
Imports System.Text

Public Class SystemUtils

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

#Region " User Login ID Routines "

    ''' <summary>
    ''' Returns the user's Login ID in lowercase, discarding any domain.
    ''' </summary>
    Public Shared Function GetUserName() As String
        Return GetUserName("")
    End Function

    ''' <summary>
    ''' Returns the user's Login ID in lowercase, discarding any domain.
    ''' </summary>
    Public Shared Function GetUserName(ByVal TestUserName As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Static LastUserName As String = ""
        ' --------------------------------
        If Not String.IsNullOrEmpty(TestUserName) AndAlso TestUserName.ToLower.EndsWith("test") Then
            Dim Result() As String
            Result = TestUserName.ToLower.Split("\"c)
            LastUserName = Result(Result.GetUpperBound(0)).ToLower
        End If
        If String.IsNullOrEmpty(LastUserName) Then
            Try
                Dim Result() As String
                Result = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString.Split("\"c)
                LastUserName = Result(Result.GetUpperBound(0)).ToLower
            Catch ex As Exception
            End Try
        End If
        If String.IsNullOrEmpty(LastUserName) Then
            Try
                Dim Result() As String
                Result = My.User.Name.ToString.Split("\"c)
                LastUserName = Result(Result.GetUpperBound(0)).ToLower
            Catch ex As Exception
            End Try
        End If
        If String.IsNullOrEmpty(LastUserName) Then
            Throw New SystemException(FuncName + vbCrLf + "Unable to determine LoginID for this user")
        End If
        Return LastUserName
    End Function

    ''' <summary>
    ''' Return the user's full name from the Active Directory listing.
    ''' If not found, will return the user's Login ID from GetUserName().
    ''' </summary>
    Public Shared Function GetUserFullName() As String
        Static LastUserFullName As String = ""
        ' ------------------------------------
        If String.IsNullOrEmpty(LastUserFullName) Then
            LastUserFullName = GetUserName() ' default to LoginID
            If ArenaConfigInfo.UseActiveDirectory Then
                Try
                    ' --- Ask Active Directory for the User Full Name ---
                    Dim CurrDE As DirectoryEntry = New DirectoryEntry("LDAP://" + My.Settings.Domain)
                    Dim CurrDS As DirectorySearcher = New DirectorySearcher(CurrDE, "(&(objectClass=user)(objectCategory=Person))")
                    For Each CurrSR As SearchResult In CurrDS.FindAll
                        Dim UserEntry As DirectoryEntry = CurrSR.GetDirectoryEntry()
                        Dim Props As System.DirectoryServices.PropertyCollection = UserEntry.Properties
                        If Props("samaccountname").Value.ToString.ToLower = GetUserName() Then
                            If String.IsNullOrEmpty(Props("givenname").Value.ToString) Then Exit For
                            If String.IsNullOrEmpty(Props("sn").Value.ToString) Then Exit For
                            LastUserFullName = Props("givenname").Value.ToString + " " + Props("sn").Value.ToString
                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    ' --- Just use the LoginID ---
                End Try
            End If
        End If
        ' --- Done ---
        Return LastUserFullName
    End Function

#End Region

#Region " Computer Routines "

    ''' <summary>
    ''' Gets the computer's name in lowercase.
    ''' </summary>
    Public Shared Function GetComputerName() As String
        Static LastComputerName As String = My.Computer.Name.ToLower
        ' ----------------------------------------------------------
        Return LastComputerName
    End Function

#End Region

#Region " End Of Month (EOM) Routines "

    Public Shared Function IsAfterMonthEndCutoff() As Boolean
        If Not InEOMPeriod() Then Return False
        If IsEOMUser() Then Return False
        Return True
    End Function

    Private Shared Property EOMStartDate As Nullable(Of DateTime) = Nothing
    Private Shared Property EOMEndDate As Nullable(Of DateTime) = Nothing
    Private Shared Property WasInEOMPeriod As Boolean = False
    Private Shared Property CurrUserIsEOMUser As Nullable(Of Boolean) = Nothing

    Private Shared Function IsEOMUser() As Boolean
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Load values if first time ---
        If Not CurrUserIsEOMUser.HasValue Then
            CurrUserIsEOMUser = False
            Dim SQLQuery As New StringBuilder
            Dim dc As New Arena_DataConn.DataConnection
            ' --- Check if user is EOM User ---
            With SQLQuery
                .AppendLine("SELECT [LoginID] FROM [IDRIS].[dbo].[%EOMUsers]")
                .Append("WHERE [LoginID] = '")
                .Append(StringToSQL(GetUserName))
                .AppendLine("'")
            End With
            Try
                Using cnIDRIS As New SqlClient.SqlConnection(dc.ConnectionString_IDRIS)
                    cnIDRIS.Open()
                    Using cmd As New SqlCommand(SQLQuery.ToString, cnIDRIS)
                        cmd.CommandType = CommandType.Text
                        Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SingleRow)
                        If dr.Read Then
                            CurrUserIsEOMUser = True
                        End If
                        dr.Close()
                    End Using
                End Using
            Catch ex As Exception
                Throw New SystemException(FuncName + vbCrLf + "Error validating EOM User" + vbCrLf + ex.Message)
            End Try
        End If
        ' --- Check if this user is allowed to save during EOM time ---
        Return CurrUserIsEOMUser.Value
    End Function

    Private Shared Function InEOMPeriod() As Boolean
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Check for time shifting in or out of EOM time ---
        If EOMStartDate.HasValue AndAlso EOMEndDate.HasValue Then
            If WasInEOMPeriod Then
                If EOMEndDate <= Now Then ' Drifted past EOM period
                    EOMStartDate = Nothing
                    EOMEndDate = Nothing
                End If
            Else
                If EOMStartDate <= Now Then ' Drifted into EOM period
                    EOMStartDate = Nothing
                    EOMEndDate = Nothing
                End If
            End If
        End If
        ' --- Load values if nothing ---
        If (Not EOMStartDate.HasValue) OrElse (Not EOMEndDate.HasValue) Then
            Dim SQLQuery As New StringBuilder
            Dim dc As New Arena_DataConn.DataConnection
            ' --- Get last EOM Start Date and Time ---
            SQLQuery.Clear()
            With SQLQuery
                .AppendLine("SELECT TOP 1 [StartDate], [EndDate] FROM [IDRIS].[dbo].[%EOMDates]")
                .AppendLine("WHERE [StartDate] <= GETDATE()")
                .AppendLine("AND [EndDate] >= GETDATE()")
                .AppendLine("ORDER BY [StartDate] DESC")
            End With
            Try
                Using cnIDRIS As New SqlClient.SqlConnection(dc.ConnectionString_IDRIS)
                    cnIDRIS.Open()
                    Using cmd As New SqlCommand(SQLQuery.ToString, cnIDRIS)
                        cmd.CommandType = CommandType.Text
                        Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SingleRow)
                        If dr.Read Then
                            If Not dr.IsDBNull(0) Then
                                EOMStartDate = dr.GetDateTime(0)
                                EOMEndDate = dr.GetDateTime(1)
                            End If
                        End If
                        dr.Close()
                    End Using
                End Using
            Catch ex As Exception
                Throw New SystemException(FuncName + vbCrLf + "Error getting next EOM Start Date" + vbCrLf + ex.Message)
            End Try
            If (Not EOMStartDate.HasValue) OrElse
               (Not EOMEndDate.HasValue) Then
                ' --- Get last EOM Start Date and Time ---
                SQLQuery.Clear()
                With SQLQuery
                    .AppendLine("SELECT TOP 1 [StartDate], [EndDate] FROM [IDRIS].[dbo].[%EOMDates]")
                    .AppendLine("WHERE [StartDate] >= GETDATE()")
                    .AppendLine("ORDER BY [StartDate]")
                End With
                Try
                    Using cnIDRIS As New SqlClient.SqlConnection(dc.ConnectionString_IDRIS)
                        cnIDRIS.Open()
                        Using cmd As New SqlCommand(SQLQuery.ToString, cnIDRIS)
                            cmd.CommandType = CommandType.Text
                            Dim dr As SqlDataReader = cmd.ExecuteReader(CommandBehavior.SingleRow)
                            If dr.Read Then
                                If Not dr.IsDBNull(0) Then
                                    EOMStartDate = dr.GetDateTime(0)
                                    EOMEndDate = dr.GetDateTime(1)
                                End If
                            End If
                            dr.Close()
                        End Using
                    End Using
                Catch ex As Exception
                    Throw New SystemException(FuncName + vbCrLf + "Error getting next EOM Start Date" + vbCrLf + ex.Message)
                End Try
                If (Not EOMStartDate.HasValue) OrElse
                   (Not EOMEndDate.HasValue) Then
                    EOMStartDate = ArenaMaxDate
                    EOMEndDate = ArenaMaxDate
                End If
            End If
        End If
        ' --- Check if in EOM period ---
        WasInEOMPeriod = False
        If EOMStartDate <= Now AndAlso EOMEndDate >= Now Then
            WasInEOMPeriod = True
        End If
        Return WasInEOMPeriod
    End Function

#End Region

#Region " Path Routines "

    ''' <summary>
    ''' Returns the normalized path lowercase ending with "\". Converts known mapped drives to UNC paths.
    ''' </summary>
    Public Shared Function NormalizePath(ByVal Pathname As String) As String
        Static UNCPathDictionary As Dictionary(Of String, String) = Nothing
        ' -----------------------------------------------------------------
        If String.IsNullOrEmpty(Pathname) Then
            Return ""
        End If
        If Not Pathname.EndsWith("\"c) Then
            Pathname += "\"c
        End If
        Try
            Dim TempPathname As String = Pathname
            If UNCPathDictionary Is Nothing Then
                UNCPathDictionary = New Dictionary(Of String, String)
                Dim MappingList As String() = My.Settings.UNCMapping.ToUpper.Split(";"c)
                For Each MappingItem As String In MappingList
                    If Not String.IsNullOrEmpty(MappingItem) Then
                        Dim MappingItems As String() = MappingItem.Split("="c)
                        UNCPathDictionary.Add(MappingItems(0), MappingItems(1).Replace("$USER$", GetUserName.ToUpper))
                    End If
                Next
            End If
            ' --- Change "Z:" to "\\SERVER\SHARE", for example, if in My.Settings.UNCMapping ---
            If UNCPathDictionary.ContainsKey(TempPathname.ToUpper.Substring(0, 2)) Then
                TempPathname = UNCPathDictionary.Item(TempPathname.ToUpper.Substring(0, 2)) + TempPathname.Substring(2)
            End If
            TempPathname = Path.GetFullPath(New Uri(TempPathname).LocalPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            ' --- Always end paths with "\" ---
            If Not TempPathname.EndsWith("\"c) Then
                TempPathname += "\"c
            End If
            Return TempPathname.ToLower
        Catch ex As Exception
            Return Pathname.ToLower
        End Try
    End Function

#End Region

#Region " CMD Routines "

    ''' <summary>
    ''' Runs the specified command in a hiden CMD window and returns the results
    ''' </summary>
    Public Shared Function CMDAutomate(ByVal cmdString As String) As String
        Dim Result As String = ""
        Dim MyProcess As New Process
        Dim StartInfo As New System.Diagnostics.ProcessStartInfo
        ' ------------------------------------------------------
        StartInfo.FileName = "cmd" 'starts cmd window        
        StartInfo.RedirectStandardInput = True
        StartInfo.RedirectStandardOutput = True
        StartInfo.UseShellExecute = False 'required to redirect    
        StartInfo.CreateNoWindow = True 'creates no cmd window         
        MyProcess.StartInfo = StartInfo
        MyProcess.Start()
        Dim SW As System.IO.StreamWriter = MyProcess.StandardInput
        Dim SR As System.IO.StreamReader = MyProcess.StandardOutput
        SW.WriteLine(cmdString) 'the commands you wish to run.....      
        SW.WriteLine("exit") 'exits command prompt window      
        Result = SR.ReadToEnd 'returns results of the command window
        SW.Close()
        SR.Close()
        Return Result
    End Function

#End Region

End Class
