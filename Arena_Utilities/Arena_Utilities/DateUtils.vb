' ---------------------------------
' --- DateUtils.vb - 08/12/2016 ---
' ---------------------------------

' ----------------------------------------------------------------------------------------------------
' 08/12/2016 - SBakker
'            - Added FormatDateTime() function for showing LastChanged and other datetimes.
' 07/22/2016 - SBakker
'            - Added IsValidDate(Integer) and UnformatDate(Integer). Apparently I missed them before.
'            - Switch all blank string comparisons to use String.IsNullOrEmpty().
' 05/04/2015 - SBakker
'            - Added comment to AddMonthsDays() that the special 10-day and 15-day handling is not
'              symetrical. Adding a day then subtracting a day won't always give the starting answer.
'              Not fixing now.
' 04/30/2015 - SBakker
'            - Added handling for dates of "0" and "00000000" so no error is thrown.
' 11/14/2014 - SBakker
'            - Changed "CDate(value)" to "CDate(FormatDate(value))" wherever "value" is a string. This
'              prevents bad conversions.
' 09/30/2014 - SBakker
'            - Switch all blank string comparisons to use String.IsNullOrWhiteSpace().
' 07/07/2014 - SBakker
'            - Removed version of CalcDurationWeeklyDays() without a NonWorkDays parameter. It is now
'              required. Same for deprecated IDRIS version.
' 07/02/2014 - SBakker
'            - Added parameter NonWorkDays to CalcDurationWeeklyDays(). If the remaining days are ones
'              the claimant doesn't work, they must be subtracted from the result. Each day would be
'              pro-rated based on the work week length, typically 5, so each day is 1/5 of the amount.              
' 03/19/2014 - SBakker
'            - Added back the "IDRIS..." functions as Deprecated routines, so no existing programs
'              will break. Marked them with the "Obsolete" attribute so they could still be used, but
'              will throw a warning.
' 03/17/2014 - SBakker
'            - Renamed all duration months/weeks/days functions to have the prefix "Calc" instead of
'              the prefix "IDRIS". Then the fields can be just "DurationMonths", "DurationWeeks", and
'              "DurationDays" with no conflicts.
' 03/11/2014 - SBakker
'            - Added new functions CalcDurationWeeks() and CalcDurationWeeklyDays() for handling STD.
' 03/10/2014 - SBakker
'            - Added the reversing routine UnformatDateToYYYYMMDD to change a date into a YYYYMMDD
'              integer. Needed for ATPTRANS/ATPFILES.
' 03/03/2014 - SBakker
'            - Added a new form of FormatDate that takes a YYYYMMDD integer and returns a "MM/DD/YYYY"
'              string. Needed for ATPTRANS/ATPFILES.
' 05/24/2013 - SBakker - URD 12013
'            - Removed default AddMonthsDays. Now must specify EOMCalc = True or False
'              (FrozenDay) everywhere.
' 05/16/2013 - SBakker
'            - Added DateTimeMilliFormatPattern to be able to access datetime including
'              milliseconds.
' 05/07/2013 - SBakker
'            - Added an EOMCalc parameter to AddMonthsDays. If it is False, it does a Frozen
'              Day calculation where adding one month to Feb 28 = March 28. If it is True,
'              adding one months goes from Feb 28 to March 31.
'            - Overloaded AddMonthsDays so the default is EOMDays = True, to use an End of
'              Month calc. Sent a note to Kim D and Trina to see if they want the company to
'              change to a Frozen Day calc. Switching True to False would do that.
' 01/22/2013 - SBakker
'            - Added new constants ArenaMinDate and ArenaMaxDate.
' 01/05/2012 - SBakker
'            - Added the ability to handle input of "YYYY-MM-DD" or "YYYYMMDD" dates, just
'              in case. This is SQL's default date display format, so dates can be pasted
'              from queries more easily.
' 07/19/2011 - SBakker
'            - Added DateTime.MinValue and DateTime.MaxValue to MinDate() and MaxDate() so
'              those values act just the same as Nothing, and the other date is returned.
' 07/15/2011 - SBakker
'            - Added MinDate() and MaxDate() functions, to be able to easily send two dates
'              and have it return the desired answer. Need to have Nothing be the wrong
'              answer for both MinDate and MaxDate if the other date has a value.
' 07/13/2011 - SBakker
'            - Added three new functions, GetNextWeekday, GetNextBusinessDay, and
'              Get15thOfNextMonth, so they can be used everywhere consistently.
' 06/29/2011 - SBakker
'            - Added constants for TimeFormatPattern and DateTimeFormatPattern.
' 05/06/2011 - SBakker
'            - Added CalcDurationMonthsDays for backfilling IDRIS payment history.
'            - Added separate CalcDurationMonths and IDRISDurationDays to get only the
'              pieces needed and not a combined value.
' 12/23/2010 - SBakker
'            - Added Special Date Routines, for handling dates as needed for various Arena
'              calculations. AddMonthsDays() will function the same as the IDRIS version.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 08/23/2010 - SBakker
'            - Added MinDateYYMMDD and MaxDateYYMMDD constants for comparisons
'              during IDRIS backfilling.
'            - Added AdjustDateToYYMMDD to handle the backfilling conversion.
' 04/29/2010 - SBakker
'            - Fixed so date strings with a time of midnight in them won't throw
'              an error - the time will just be removed.
'            - Added FormatDate(Object), so it can handle Nothing before trying to
'              CType it to String or DateTime or DateTime?.
'            - Added logic back so PastOnly parameter works. Allows Today, but not
'              a future date. Typically used for Date of Birth fields.
' 04/06/2010 - SBakker
'            - Removed automatic handling of 2-digit years as requested by Claims.
'              All years must be entered now as 4 digits!
'            - Added handling for single-digit months and days which are valid,
'              such as "892010" = "08/09/2010".
'            - Changed MinYear to 1900 and MaxYear to 2099. Will help avoid bad
'              dates. Can be bumped forward in later years. (Hi!)
' 01/21/2010 - SBakker
'            - Moved DateFormatSQLCode to here so it can be used by all classes.
' 08/10/2009 - SBakker
'            - Added "UnformatDatePartial" functions to handle searching for
'              partial dates.
' 04/23/2009 - SBakker
'            - Added UnformatDate, to turn a String back into a DateTime?.
' 03/09/2009 - SBakker
'            - For a DateTime, DateTime.MinValue is the same as Nothing, and
'              should return "".
'            - For a nullable DateTime?, use the corresponding DateTime function
'              on it's .Value, if it has one, rather than duplicating code.
' 03/03/2009 - SBakker
'            - Added FormatDatePast and IsValidDatePast, for dates that must be
'              today or earlier. (No time travellers from the future allowed!)
'            - Moved the guts of FormatDate to a private function DoFormatDate.
' 01/26/2009 - SBakker
'            - Added DateUtils to handle formatting and validating of dates.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.FormatUtils

Public Class DateUtils

#Region " Constants "

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Public Const DateFormatPattern As String = "MM/dd/yyyy" ' USA Date Format with leading zeros
    Public Const DateFormatSQLCode As Integer = 101 ' USA Date Format
    Public Const TwoDigitCutoffYear As Integer = 55 ' 1955 to 2054, used for IDRIS data only
    Public Const MinDateYYMMDD As Date = #1/1/1955# ' Must match 1900 + TwoDigitCutoffYear
    Public Const MaxDateYYMMDD As Date = #12/31/2054# ' Must match 2000 + TwoDigitCutoffYear - 1
    Public Const MinYear As Integer = 1900
    Public Const MaxYear As Integer = 2099 ' Can be changed to 9999 when all YYMMDD fields are gone
    Public Const ArenaMinDate As Date = #1/1/1900#
    Public Const ArenaMaxDate As Date = #12/31/2199 11:59:59 PM#

    Public Const TimeFormatPattern As String = "HH:mm:ss" ' 24-hour local time, down to the second

    Public Const DateTimeFormatPattern As String = DateFormatPattern + " " + TimeFormatPattern

    Public Const DateTimeMilliFormatPattern As String = DateFormatPattern + " " + TimeFormatPattern + ".fff"

#End Region

#Region " FormatDate Routines "

    Public Shared Function FormatDate(ByVal Value As Object) As String
        If Value Is Nothing Then
            Return ""
        End If
        If TypeOf Value Is System.String Then
            Return FormatDate(CStr(Value))
        End If
        If TypeOf Value Is System.DateTime Then
            Return FormatDate(CDate(Value))
        End If
        Return FormatDate(CType(Value, DateTime?))
    End Function

    Public Shared Function FormatDate(ByVal Value As DateTime?) As String
        If Value.HasValue Then
            Return FormatDate(Value.Value)
        End If
        Return ""
    End Function

    Public Shared Function FormatDate(ByVal Value As DateTime) As String
        If Value <> DateTime.MinValue Then
            Return Value.ToString(DateFormatPattern)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Converts integer dates in the format YYYYMMDD to "MM/DD/YYYY"
    ''' </summary>
    Public Shared Function FormatDate(ByVal Value As Integer) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If Value = 0 Then Return ""
        If Value < 0 Or Value > 99991231 Then
            Throw New SystemException(FuncName + vbCrLf + "Invalid Date: """ + Value.ToString + """")
        End If
        Try
            Dim TempDateValue As String = Right("00000000" + Value.ToString, 8)
            Dim NewTempDate As String = TempDateValue.Substring(4, 2) + "/" + TempDateValue.Substring(6, 2) + "/" + TempDateValue.Substring(0, 4)
            Return DoFormatDate(NewTempDate, False)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Date: """ + Value.ToString + """")
        End Try
    End Function

    Public Shared Function FormatDate(ByVal Value As String) As String
        Return DoFormatDate(Value, False)
    End Function

#End Region

#Region " FormatDatePast Routines "

    Public Shared Function FormatDatePast(ByVal Value As DateTime?) As String
        If Value.HasValue Then
            Return FormatDatePast(Value.Value)
        End If
        Return ""
    End Function

    Public Shared Function FormatDatePast(ByVal Value As DateTime) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        If Value <> DateTime.MinValue Then
            If Value > Today Then
                Throw New SystemException(FuncName + vbCrLf + "Date isn't today or earlier: """ + Value.ToString(DateFormatPattern) + """")
            End If
            Return Value.ToString(DateFormatPattern)
        End If
        Return ""
    End Function

    Public Shared Function FormatDatePast(ByVal Value As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- This date must be today or earlier ---
        Dim Result As String = DoFormatDate(Value, True) ' round century down, not up
        If Not String.IsNullOrEmpty(Result) Then
            If CDate(Result) > Today Then
                Throw New SystemException(FuncName + vbCrLf + "Date isn't today or earlier: """ + Value + """")
            End If
        End If
        Return Result
    End Function

#End Region

#Region " FormatDateTime Routines "

    Public Shared Function FormatDateTime(ByVal Value As DateTime?) As String
        If Not Value.HasValue Then
            Return ""
        ElseIf Value.Value.Hour + Value.Value.Minute + Value.Value.Second = 0 Then
            ' --- Has no time value ---
            Return Value.Value.ToString(DateFormatPattern)
        Else
            Return Value.Value.ToString(DateTimeFormatPattern)
        End If
    End Function

#End Region

#Region " UnformatDate Routines "

    Public Shared Function UnformatDate(ByVal Value As String) As DateTime?
        If String.IsNullOrEmpty(Value) OrElse Value.Trim = "0" OrElse Value = "00000000" Then
            Return Nothing
        ElseIf IsValidDate(Value) Then
            Return CDate(FormatDate(Value))
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function UnformatDate(ByVal Value As Integer) As DateTime?
        If Value = 0 Then
            Return Nothing
        ElseIf IsValidDate(Value) Then
            Return CDate(FormatDate(Value))
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function UnformatDateToYYYYMMDD(ByVal Value As DateTime?) As Integer
        If Not Value.HasValue Then
            Return 0
        Else
            Return CInt(Value.Value.ToString("yyyyMMdd"))
        End If
    End Function

    Public Shared Function UnformatDateToYYYYMMDD(ByVal Value As String) As Integer
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Try
            If String.IsNullOrEmpty(Value) Then
                Return 0
            Else
                Return UnformatDateToYYYYMMDD(CDate(FormatDate(Value)))
            End If
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + "Invalid Date: """ + Value + """")
        End Try
    End Function

    Public Shared Function UnformatDatePartial(ByVal Value As String) As String
        Return UnformatDatePartial(Value, False)
    End Function

    Public Shared Function UnformatDatePartialPast(ByVal Value As String) As String
        Return UnformatDatePartial(Value, True)
    End Function

    Public Shared Function UnformatDatePartial(ByVal Value As String, ByVal PastOnly As Boolean) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim Result As String = ""
        ' --- Empty string is valid. Calling program must trap it first, if it shouldn't be valid. ---
        If String.IsNullOrEmpty(Value) Then
            Return ""
        End If
        ' --- Check for wildcards ---
        If StringUtils.HasWildcards(Value) Then
            Return Value
        End If
        ' --- Parse the value ---
        Try
            Dim SectionLen() As Integer = {2, 2, 4}
            Dim MinValues() As Integer = {0, 0, 0}
            Result = FormatBySection(Value, 3, SectionLen, MinValues, "/", True)
            If IsValidDatePartial(Result) Then Return DoFormatDate(Result, PastOnly)
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {2, 2}
            Dim MinValues() As Integer = {0, 0}
            Result = FormatBySection(Value, 2, SectionLen, MinValues, "/", True)
            If IsValidDatePartial(Result) Then Return Result
        Catch ex As Exception
        End Try
        Try
            Dim SectionLen() As Integer = {2}
            Dim MinValues() As Integer = {0}
            Result = FormatBySection(Value, 1, SectionLen, MinValues, "/", True)
            If IsValidDatePartial(Result) Then Return Result
        Catch ex As Exception
        End Try
        Throw New SystemException(FuncName + vbCrLf + "Invalid Partial Date: """ + Value + """")
    End Function

#End Region

#Region " IsValidDate Routines "

    Public Shared Function IsValidDate(ByVal Value As String) As Boolean
        Try
            Dim TempDate As String = FormatDate(Value)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function IsValidDate(ByVal Value As Integer) As Boolean
        Try
            Dim TempDate As String = FormatDate(Value)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function IsValidDatePast(ByVal Value As String) As Boolean
        Try
            Dim TempDate As String = FormatDatePast(Value)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function IsValidDatePartial(ByVal Value As String) As Boolean
        Dim TempMonth As Integer = -1
        Dim TempDay As Integer = -1
        Dim TempYear As Integer = -1
        ' ---------------------------
        Try
            If String.IsNullOrEmpty(Value) Then Return True
            Dim Sections() As String = Value.Split("/"c)
            If Sections(0).Length <> 2 Then Return False
            TempMonth = CInt(Sections(0))
            If TempMonth < 1 OrElse TempMonth > 12 Then Return False
            If Sections.GetUpperBound(0) = 0 Then Return True
            If Sections(1).Length <> 2 Then Return False
            TempDay = CInt(Sections(1))
            If TempDay < 1 OrElse TempDay > 31 Then Return False
            If TempDay = 31 Then
                If TempMonth = 2 Then Return False
                If TempMonth = 4 Then Return False
                If TempMonth = 6 Then Return False
                If TempMonth = 9 Then Return False
                If TempMonth = 11 Then Return False
            End If
            If TempDay = 30 Then
                If TempMonth = 2 Then Return False
            End If
            If Sections.GetUpperBound(0) = 1 Then Return True
            TempYear = CInt(Sections(2))
            ' --- Removed 04/05/2010 as requested by Claims - 4-digit years only! ---
            'If TempYear >= 0 AndAlso TempYear <= 99 Then
            '    If TempYear < TwoDigitCutoffYear Then
            '        TempYear += 2000
            '    Else
            '        TempYear += 1900
            '    End If
            'End If
            If TempYear < MinYear OrElse TempYear > MaxYear Then Return False
            If TempMonth = 2 AndAlso TempDay = 29 Then
                If TempYear Mod 4 <> 0 Then Return False
                If TempYear Mod 100 = 0 AndAlso TempYear Mod 400 <> 0 Then Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region " YYMMDD Routines "

    Public Shared Function AdjustDateToYYMMDD(ByVal Value As Nullable(Of DateTime)) As Nullable(Of DateTime)
        If Not Value.HasValue Then
            Return Value
        End If
        If Value.Value < MinDateYYMMDD Then
            Return MinDateYYMMDD
        End If
        If Value.Value > MaxDateYYMMDD Then
            Return MaxDateYYMMDD
        End If
        Return Value
    End Function

#End Region

#Region " Special Date Routines "

    '' --- Removed default AddMonthsDays. Now must specify EOMCalc = True or False (FrozenDay) everywhere ---
    ''Public Shared Function AddMonthsDays(ByVal Value As DateTime, ByVal NumMonths As Integer, ByVal NumDays As Integer) As DateTime
    ''    ' --- Use End of Month calc by default to match old IDRIS ---
    ''    Return AddMonthsDays(Value, NumMonths, NumDays, False)
    ''End Function

    ''' <summary>
    ''' This will add months and days to the specified date using special date logic.
    ''' Ends of months are handled uniformly, so adding one month to Mar 31 returns Apr 30.
    ''' Adding 10 or 15 days will return the end of the month if the answer is the 30th.
    ''' This also properly handles the adding of negative months and/or days.
    ''' </summary>
    Public Shared Function AddMonthsDays(ByVal Value As DateTime, ByVal NumMonths As Integer, ByVal NumDays As Integer, ByVal EOMCalc As Boolean) As DateTime
        Dim CurrYear As Integer = Value.Year
        Dim CurrMonth As Integer = Value.Month
        Dim CurrDay As Integer = Value.Day
        ' ------------------------------------
        If EOMCalc Then
            ' --- Find if currently at the end of a month ---
            If CurrDay = DaysInMonth(CurrMonth, CurrYear) Then
                CurrDay = 31 ' force to highest end of month
            End If
        End If
        ' --- Add months ---
        CurrMonth += NumMonths
        ' --- Normalize months and years ---
        Do While CurrMonth > 12
            CurrYear += 1
            CurrMonth -= 12
        Loop
        Do While CurrMonth <= 0
            CurrYear -= 1
            CurrMonth += 12
        Loop
        ' --- Adjust back to end of new month ---
        If CurrDay > DaysInMonth(CurrMonth, CurrYear) Then
            CurrDay = DaysInMonth(CurrMonth, CurrYear)
        End If
        ' --- Add days ---
        CurrDay += NumDays
        ' --- Check for special 10-day or 15-day logic ---
        If (NumDays = 10 OrElse NumDays = 15) AndAlso CurrDay = 30 Then
            ' --- Not symetrical! This causes "+ 1, - 1" to not always give the same starting value. ---
            CurrDay = DaysInMonth(CurrMonth, CurrYear)
        End If
        ' --- Normalize days ---
        Do While CurrDay > DaysInMonth(CurrMonth, CurrYear)
            ' --- Subtract first ---
            CurrDay -= DaysInMonth(CurrMonth, CurrYear)
            ' --- Adjust second ---
            CurrMonth += 1
            If CurrMonth > 12 Then
                CurrYear += 1
                CurrMonth -= 12
            End If
        Loop
        Do While CurrDay <= 0
            ' --- Adjust first ---
            CurrMonth -= 1
            If CurrMonth <= 0 Then
                CurrYear -= 1
                CurrMonth += 12
            End If
            ' --- Add second ---
            CurrDay += DaysInMonth(CurrMonth, CurrYear)
        Loop
        ' --- Return result ---
        Return New DateTime(CurrYear, CurrMonth, CurrDay)
    End Function

    ''' <summary>
    ''' This returns the duration in months and days as a single value in MMMMDD format.
    ''' </summary>
    Public Shared Function CalcDurationMonthsDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        ' --- Check for errors ---
        If FromDate > ToDate Then
            Throw New SystemException(FuncName + vbCrLf + "FromDate cannot be after ToDate")
        End If

        Dim ResultMonths As Integer = 0
        Dim ResultDays As Integer = 0

        Dim DaysInMonthFrom As Integer = DaysInMonth(FromDate.Month, FromDate.Year)
        Dim DaysInMonthTo As Integer = DaysInMonth(ToDate.Month, ToDate.Year)
        Dim TempFromDay As Integer = FromDate.Day
        Dim TempToDay As Integer = ToDate.Day
        Dim DoNormalCalc As Boolean = False
        Dim DoneCalc As Boolean = False

        ' --- Calculate Months ---
        ResultMonths = (ToDate.Year - FromDate.Year) * 12
        ResultMonths += ToDate.Month - FromDate.Month

        ' --- Check for special situations ---
        If ResultMonths = 0 Then DoNormalCalc = True ' same month
        If TempFromDay = TempToDay Then DoNormalCalc = True ' same day of another month

        If Not DoNormalCalc Then

            If TempFromDay = DaysInMonthFrom Then ' end of first month
                If TempToDay = DaysInMonthTo Then DoneCalc = True ' end of second month, no days
                If TempToDay > TempFromDay Then DoneCalc = True ' 2nd month has more days
            End If

            If TempToDay = DaysInMonthTo Then ' end of second month
                If TempToDay < TempFromDay Then DoneCalc = True ' 2nd month has less days
            End If

            If Not DoneCalc Then

                If ResultMonths <> 1 Then  ' not next month
                    If TempToDay <= TempFromDay Then ' don't count as actual days
                        If TempFromDay > 30 Then TempFromDay = 30 ' no more than 30 days
                        If TempToDay > 30 Then TempToDay = 30 ' no more than 30 days
                        If DaysInMonthFrom < 30 AndAlso TempFromDay = DaysInMonthFrom Then TempFromDay = 30 ' end of february, make 30 days
                        If DaysInMonthTo < 30 AndAlso TempToDay = DaysInMonthTo Then TempToDay = 30 ' end of february, make 30 days
                    End If
                Else
                    If TempFromDay > TempToDay Then ' not one full month
                        ResultDays = DaysInMonthFrom - TempFromDay + TempToDay ' actual days between dates
                        ResultMonths = 0 ' no months
                        DoneCalc = True
                    End If
                End If

            End If

            If Not DoneCalc Then DoNormalCalc = True

        End If

        If DoNormalCalc Then
            If TempFromDay <= TempToDay Then ' from days before to days
                ResultDays = TempToDay - TempFromDay ' number of days
            Else
                ResultMonths -= 1 ' subtract one month
                ResultDays = TempToDay - TempFromDay + 30 ' always 30 days per month
            End If
        End If

        ' --- Done ---
        Return (ResultMonths * 100) + ResultDays

    End Function

    ''' <summary>
    ''' This returns the duration in IDRIS 30-day months between the two dates.
    ''' </summary>
    Public Shared Function CalcDurationMonths(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationMonthsDays(FromDate, ToDate) \ 100
    End Function

    ''' <summary>
    ''' This returns the remaining days (from IDRIS 30-day months) between the two dates.
    ''' </summary>
    Public Shared Function CalcDurationMonthlyDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationMonthsDays(FromDate, ToDate) Mod 100
    End Function

    ''' <summary>
    ''' This returns the duration in whole weeks between the two dates.
    ''' </summary>
    Public Shared Function CalcDurationWeeks(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- Check for errors ---
        If FromDate > ToDate Then
            Throw New SystemException(FuncName + vbCrLf + "FromDate cannot be after ToDate")
        End If
        ' --- Return whole weeks ---
        Return CInt(DateDiff(DateInterval.Day, FromDate, ToDate)) \ 7
    End Function

    ''' <summary>
    ''' This returns the duration in days after whole weeks between the two dates. NonWorkDays are subtracted.
    ''' </summary>
    Public Shared Function CalcDurationWeeklyDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime, ByVal NonWorkDays As String) As Integer
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim Result As Integer
        ' --- Check for errors ---
        If FromDate > ToDate Then
            Throw New SystemException(FuncName + vbCrLf + "FromDate cannot be after ToDate")
        End If
        ' --- Return remaining days ---
        Result = CInt(DateDiff(DateInterval.Day, FromDate, ToDate)) Mod 7
        ' --- Subtract out NonWorkDays ---
        If Result > 0 AndAlso Not String.IsNullOrEmpty(NonWorkDays) Then
            Dim TempDays As Integer = Result ' Must save, because subtracting from Result
            Dim TempDate As DateTime
            For PriorDay As Integer = 1 To TempDays
                TempDate = DateAdd(DateInterval.Day, -PriorDay, ToDate)
                For Each CurrNonWorkDay As Char In NonWorkDays
                    Select Case CurrNonWorkDay ' MTWRFSU
                        Case "M"c
                            If TempDate.DayOfWeek = DayOfWeek.Monday Then
                                Result -= 1
                            End If
                        Case "T"c
                            If TempDate.DayOfWeek = DayOfWeek.Tuesday Then
                                Result -= 1
                            End If
                        Case "W"c
                            If TempDate.DayOfWeek = DayOfWeek.Wednesday Then
                                Result -= 1
                            End If
                        Case "R"c
                            If TempDate.DayOfWeek = DayOfWeek.Thursday Then
                                Result -= 1
                            End If
                        Case "F"c
                            If TempDate.DayOfWeek = DayOfWeek.Friday Then
                                Result -= 1
                            End If
                        Case "S"c
                            If TempDate.DayOfWeek = DayOfWeek.Saturday Then
                                Result -= 1
                            End If
                        Case "U"c
                            If TempDate.DayOfWeek = DayOfWeek.Sunday Then
                                Result -= 1
                            End If
                    End Select
                Next
            Next
        End If
        ' --- Done ---
        Return Result
    End Function

    ''' <summary>
    ''' This returns the last day of the specified month and year.
    ''' </summary>
    Public Shared Function DaysInMonth(ByVal Month As Integer, ByVal Year As Integer) As Integer
        Dim Result As Integer
        If Month = 4 OrElse Month = 6 OrElse Month = 9 OrElse Month = 11 Then ' april, june, september, november
            Result = 30
        ElseIf Month = 2 Then ' february
            Result = 28
            If Year Mod 4 = 0 Then Result = 29 ' divisible by 4...
            If Year Mod 100 = 0 Then Result = 28 ' ...but not 100...
            If Year Mod 400 = 0 Then Result = 29 ' ...or by 400
        Else ' all others
            Result = 31
        End If
        Return Result
    End Function

    Public Shared Function GetNextWeekday(ByVal CurrDate As DateTime) As DateTime
        Dim IsWeekday As Boolean
        Dim Result As DateTime = CurrDate
        ' -------------------------------
        Do
            Result = DateAdd(DateInterval.Day, 1, Result)
            IsWeekday = True
            ' --- Check for weekends ---
            If Result.DayOfWeek = DayOfWeek.Saturday Then IsWeekday = False
            If Result.DayOfWeek = DayOfWeek.Sunday Then IsWeekday = False
        Loop Until IsWeekday
        Return Result
    End Function

    Public Shared Function GetNextBusinessDay(ByVal CurrDate As DateTime) As DateTime
        Dim IsBusinessDay As Boolean
        Dim Result As DateTime = CurrDate
        ' -------------------------------
        Do
            Result = DateAdd(DateInterval.Day, 1, Result)
            IsBusinessDay = True
            ' --- Check for weekends ---
            If Result.DayOfWeek = DayOfWeek.Saturday Then IsBusinessDay = False
            If Result.DayOfWeek = DayOfWeek.Sunday Then IsBusinessDay = False
            ' --- Check for exact date holidays ---
            If Result.Month = 1 And Result.Day = 1 Then IsBusinessDay = False ' January 1
            If Result.Month = 7 And Result.Day = 4 Then IsBusinessDay = False ' July 4
            If Result.Month = 12 And Result.Day = 25 Then IsBusinessDay = False ' December 25
            ' --- Check for Monday Federal holidays ---
            If Result.DayOfWeek = DayOfWeek.Monday Then
                If Result.Month = 1 Then ' January
                    If Result.Day >= 15 AndAlso Result.Day <= 21 Then IsBusinessDay = False ' Martin Luther King Jr's Birthday
                End If
                If Result.Month = 2 Then ' February
                    If Result.Day >= 15 AndAlso Result.Day <= 21 Then IsBusinessDay = False ' President's Day
                End If
                If Result.Month = 5 Then ' May
                    If Result.Day + 7 > 31 Then IsBusinessDay = False ' Memorial Day
                End If
                If Result.Month = 9 Then ' September
                    If Result.Day <= 7 Then IsBusinessDay = False ' Labor Day
                End If
            End If
            ' --- Check for Thanksgiving ---
            If Result.Month = 11 AndAlso Result.DayOfWeek = DayOfWeek.Thursday Then
                If Result.Day >= 22 AndAlso Result.Day <= 28 Then IsBusinessDay = False ' Thanksgiving
            End If
        Loop Until IsBusinessDay
        Return Result
    End Function

    Public Shared Function Get15thOfNextMonth(ByVal CurrDate As DateTime) As DateTime
        Dim TempMonth As Integer = CurrDate.Month + 1
        Dim TempYear As Integer = CurrDate.Year
        If TempMonth > 12 Then
            TempYear += 1
            TempMonth -= 12
        End If
        Return New DateTime(TempYear, TempMonth, 15)
    End Function

#End Region

#Region " Compare Routines "

    Public Shared Function MinDate(ByVal Date1 As DateTime?, ByVal Date2 As DateTime?) As DateTime?
        If Not Date1.HasValue Then Return Date2
        If Not Date2.HasValue Then Return Date1
        Return MinDate(Date1.Value, Date2.Value)
    End Function

    Public Shared Function MinDate(ByVal Date1 As DateTime, ByVal Date2 As DateTime) As DateTime
        If Date1 = DateTime.MinValue OrElse Date1 = DateTime.MaxValue Then Return Date2
        If Date2 = DateTime.MinValue OrElse Date2 = DateTime.MaxValue Then Return Date1
        If Date1 < Date2 Then
            Return Date1
        Else
            Return Date2
        End If
    End Function

    Public Shared Function MaxDate(ByVal Date1 As DateTime?, ByVal Date2 As DateTime?) As DateTime?
        If Not Date1.HasValue Then Return Date2
        If Not Date2.HasValue Then Return Date1
        Return MaxDate(Date1.Value, Date2.Value)
    End Function

    Public Shared Function MaxDate(ByVal Date1 As DateTime, ByVal Date2 As DateTime) As DateTime
        If Date1 = DateTime.MinValue OrElse Date1 = DateTime.MaxValue Then Return Date2
        If Date2 = DateTime.MinValue OrElse Date2 = DateTime.MaxValue Then Return Date1
        If Date1 > Date2 Then
            Return Date1
        Else
            Return Date2
        End If
    End Function

#End Region

#Region " Internal Routines "

    Private Shared Function DoFormatDate(ByVal Value As String, ByVal PastOnly As Boolean) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        ' --- This will return a valid date, a null string, or throws an error ---
        ' --- PastOnly will only control the adding of a century to a 2-digit year, but the date may still be in the future. ---
        If String.IsNullOrEmpty(Value) OrElse Value.Trim = "0" OrElse Value = "00000000" Then
            Return ""
        End If
        Dim TempValue As String = Value
        ' --- Ignore midnight. Other times with throw an error. ---
        If TempValue.ToUpper.EndsWith("12:00:00 AM") Then
            TempValue = TempValue.Substring(0, Len(TempValue) - Len("12:00:00 AM"))
        End If
        If TempValue.EndsWith("00:00:00") Then
            TempValue = TempValue.Substring(0, Len(TempValue) - Len("00:00:00"))
        End If
        If TempValue.EndsWith(" 0:00:00") Then
            TempValue = TempValue.Substring(0, Len(TempValue) - Len(" 0:00:00"))
        End If
        ' --- Trim to have fewer chars to process ---
        TempValue = TempValue.Trim
        ' --- Check for empty string ---
        If String.IsNullOrEmpty(TempValue) Then
            Return ""
        End If
        ' --- Parse the date string into pieces ---
        Dim _Month As Integer = 0
        Dim _Day As Integer = 0
        Dim _Year As Integer = 0
        Dim _Section As Integer = 0
        Dim _InSection As Boolean = False
        Dim _CharCount As Integer = 0
        ' --- Check for YYYY-MM-DD or YYYYMMDD dates first, just in case ---
        If StringUtils.DigitsOnly(Left(TempValue, 4)) AndAlso CInt(Left(TempValue, 4)) >= MinYear AndAlso CInt(Left(TempValue, 4)) <= MaxYear Then
            ' --- Check each character an put into a section of the date ---
            For Each c As Char In TempValue
                If c >= "0"c AndAlso c <= "9"c Then
                    If Not _InSection Then
                        _Section += 1
                        _InSection = True
                        _CharCount = 0
                    End If
                    _CharCount += 1
                    ' --- If four digit year, jump to next section ---
                    If _Section < 2 AndAlso _CharCount > 4 Then
                        _Section += 1
                        _CharCount = 1
                    ElseIf _Section >= 2 AndAlso _CharCount > 2 Then ' --- Also check two-digit month ---
                        _Section += 1
                        _CharCount = 1
                    End If
                    Select Case _Section
                        Case 1
                            _Year = (_Year * 10) + Val(c)
                            If _Year >= MinYear Then
                                _InSection = False
                            End If
                        Case 2
                            _Month = (_Month * 10) + Val(c)
                            If _Month > 1 Then
                                _InSection = False
                            End If
                        Case 3
                            _Day = (_Day * 10) + Val(c)
                            If _Day > 3 Then
                                _InSection = False
                            End If
                        Case Else
                            GoTo FoundError
                    End Select
                Else
                    _InSection = False
                End If
            Next
        Else
            ' --- Check each character an put into a section of the date ---
            For Each c As Char In TempValue
                If c >= "0"c AndAlso c <= "9"c Then
                    If Not _InSection Then
                        _Section += 1
                        _InSection = True
                        _CharCount = 0
                    End If
                    _CharCount += 1
                    ' --- If two digit month/day, jump to next section ---
                    If _Section < 3 AndAlso _CharCount > 2 Then
                        _Section += 1
                        _CharCount = 1
                    End If
                    Select Case _Section
                        Case 1
                            _Month = (_Month * 10) + Val(c)
                            If _Month > 1 Then
                                _InSection = False
                            End If
                        Case 2
                            _Day = (_Day * 10) + Val(c)
                            If _Day > 3 Then
                                _InSection = False
                            End If
                        Case 3
                            _Year = (_Year * 10) + Val(c)
                        Case Else
                            GoTo FoundError
                    End Select
                Else
                    _InSection = False
                End If
            Next
        End If
        ' --- Check for all zeros ---
        If _Section >= 3 AndAlso _Month = 0 AndAlso _Day = 0 AndAlso _Year = 0 Then
            Return ""
        End If
        ' --- Check for invalid values ---
        If _Month < 1 OrElse _Month > 12 OrElse _Day < 1 OrElse _Day > 31 Then
            GoTo FoundError
        End If
        ' --- Handle 2-digit years ---
        ' --- Removed 04/05/2010 as requested by Claims - 4-digit years only! ---
        'If PastOnly AndAlso _Year < TwoDigitCutoffYear AndAlso _Year > (Year(Today) Mod 100) Then
        '    _Year += 1900
        'ElseIf _Year < TwoDigitCutoffYear Then
        '    _Year += 2000
        'ElseIf _Year < 100 Then
        '    _Year += 1900
        'End If
        ' --- Check for invalid year ---
        If _Year < MinYear OrElse _Year > MaxYear Then
            GoTo FoundError
        End If
        ' --- Check for invalid days in month ---
        Select Case _Month
            Case 9, 4, 6, 11 ' September, April, June, November
                If _Day > 30 Then
                    GoTo FoundError
                End If
            Case 2 ' February
                If _Day > 29 Then
                    GoTo FoundError
                End If
                If _Day > 28 Then
                    If _Year Mod 4 <> 0 Then
                        GoTo FoundError
                    ElseIf _Year Mod 100 = 0 AndAlso _Year Mod 400 <> 0 Then
                        GoTo FoundError
                    End If
                End If
        End Select
        ' --- Combine the pieces into a date ---
        Dim Result As DateTime = CDate(_Month.ToString + "/" + _Day.ToString + "/" + _Year.ToString)
        ' --- Don't allow future dates ---
        If PastOnly Then
            If Result > Today Then GoTo FoundError
        End If
        ' --- Return the formatted date ---
        Return Result.ToString(DateFormatPattern)
        ' --- Error found while parsing ---
FoundError:
        Throw New SystemException(FuncName + vbCrLf + "Invalid Date: """ + Value + """")
    End Function

#End Region

#Region " DEPRECATED Routines "

    ''' <summary>
    ''' DEPRECATED: Use CalcDurationMonthsDays.
    ''' </summary>
    <Obsolete("Please switch to CalcDurationMonthsDays()", False)>
    Public Shared Function IDRISDurationMonthsDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationMonthsDays(FromDate, ToDate)
    End Function

    ''' <summary>
    ''' DEPRECATED: Use CalcDurationMonths.
    ''' </summary>
    <Obsolete("Please switch to CalcDurationMonths()", False)>
    Public Shared Function IDRISDurationMonths(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationMonths(FromDate, ToDate)
    End Function

    ''' <summary>
    ''' DEPRECATED: Use CalcDurationMonthlyDays.
    ''' </summary>
    <Obsolete("Please switch to CalcDurationMonthlyDays()", False)>
    Public Shared Function IDRISDurationDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationMonthlyDays(FromDate, ToDate)
    End Function

    ''' <summary>
    ''' DEPRECATED: Use CalcDurationWeeks.
    ''' </summary>
    <Obsolete("Please switch to CalcDurationWeeks()", False)>
    Public Shared Function DurationWeeks(ByVal FromDate As DateTime, ByVal ToDate As DateTime) As Integer
        Return CalcDurationWeeks(FromDate, ToDate)
    End Function

    ''' <summary>
    ''' DEPRECATED: Use CalcDurationWeeklyDays.
    ''' </summary>
    <Obsolete("Please switch to CalcDurationWeeklyDays()", False)>
    Public Shared Function DurationWeeklyDays(ByVal FromDate As DateTime, ByVal ToDate As DateTime, ByVal NonWorkDays As String) As Integer
        Return CalcDurationWeeklyDays(FromDate, ToDate, NonWorkDays)
    End Function

#End Region

End Class
