' --------------------------------------
' --- ArenaClassBuilder - 08/03/2016 ---
' --------------------------------------

' ----------------------------------------------------------------------------------------------------
' 08/03/2016 - SBakker
'            - Added handling for "uniqueidentifier", "sysname", and "text" data types.
' 04/29/2016 - SBakker
'            - Send the RowsChanged parameter to the RecordMissingOrChangedErrorMsg() routine so it
'              can be displayed in the error.
' 02/02/2016 - SBakker
'            - Added soft errors for Path Not Found.
'            - Switched from complicated LastFromPath, LastToPath to simple calculated fields.
' 01/13/2016 - SBakker
'            - Removed Try...Catch around filling fields. If the field name lists don't match, throw
'              an error instead of skipping the field. The field name lists should always match!
' 01/12/2016 - SBakker
'            - Ignore calculated fields with the format "[fieldname] AS ...".
' 10/28/2015 - SBakker
'            - Removed ExactClone. It causes major errors if used to save and restore objects.
' 10/14/2015 - SBakker - URD 12527
'            - Added code for setting "cmd.CommandTimeout = DataConnection.SQLTimeoutSeconds"
'              everywhere that a SQLCommand object is created. This might help with the infrequent
'              timeout errors.
' 05/28/2015 - SBakker
'            - Exclude vbCr from non-multiline string properties also.
'            - Wrap entire body of Save() command in Try...Catch statement to prevent timeouts.
'            - Fixed IDisposable support to properly Dispose of internal cmd and dc objects.
' 03/31/2015 - SBakker
'            - Renamed "My.Settings.Applications" to "My.Settings.ApplicationList". Apparently you
'              can't change a User setting to an Application setting and have it work properly without
'              renaming it.
' 03/11/2015 - SBakker
'            - Added Arena_Imports application to settings.
'            - Changed "My.Settings.Applications" setting to be an application level setting.
' 01/23/2015 - SBakker
'            - Check for either vbCr or vbLf to determine Multiline. ### Will implement later ###
' 06/16/2014 - SBakker
'            - Use "##ORIG##" for original field name before fixing. This allows filling and saving
'              using the SQL field name, instead of the fixed property name.
' 04/08/2014 - SBakker
'            - Added code to call RaiseFieldChangedEvent() whenever a field changes, if not filling
'              fields and UseDebugMode is true. The two boolean fields are checked first, to avoid the
'              overhead of making strings and calling the routine when it isn't needed or wanted.
'              Allows field-level debugging without loading the data layer project in the solution. :)
' 02/24/2014 - SBakker
'            - Added Bootstrap loading all programs to another location, and then running from there.
' 01/10/2014 - SBakker
'            - Must delete record in Save() routine if applicable before calling FixRecordPreSave(),
'              or the number of BeginTransactions/EndTransactions could become mismatched.
' 01/02/2014 - SBakker
'            - Added "ValidatePreDelete()" in case there are constraints which need checking to be
'              sure that the record may be deleted.
' 12/23/2013 - SBakker
'            - Also split out "datetime" in the Value scripts, used when creating new records. Doh!
' 12/20/2013 - SBakker
'            - Split out the "datetime" Update scripts from "date" and "smalldatetime". The "datetime"
'              format needs to be able to save milliseconds, while the others don't.
'            - Only include "Imports Arena_Utilities.DateUtils" if there is a DateTime field in the
'              class.
' 12/13/2013 - SBakker
'            - Added a ChangedCount to show how many classes were updated.
' 10/22/2013 - SBakker
'            - Allow any CHAR/VARCHAR fields with MultilineMinLen or greater length to be multiline.
' 10/04/2013 - SBakker
'            - Strip the Drive from all FromPath/ToPath strings upon loading. Makes it easier so you
'              don't have to do it yourself.
'            - Only enable ButtonBuildAll when something is selected.
'            - Add the Drive to the filepath when checking for ".INFO" files.
' 10/01/2013 - SBakker
'            - Added Drive Combo Box.
' 09/17/2013 - SBakker
'            - Added "Partial Private Shared Sub ValidateRecordLevel()". This allows Record level
'              validation to be added in a partial class, as needed.
'            - Added additional error information.
' 08/19/2013 - SBakker
'            - Added "smalldatetime" support.
' 07/17/2013 - SBakker
'            - Added FixFieldName() to handle any invalid field names, such as "100PctRate".
'            - Added FixClassName() to fix NoteDiary (from below), and also IDRIS class names that
'              would conflict with Arena class names.
' 05/31/2013 - SBakker
'            - Some NoteDiary class names are reserved words, so add "NoteDiary_" at their
'              beginning.
' 05/28/2013 - SBakker
'            - Added the ability to handle "text" fields, although "varchar(max)" is more
'              appropriate for newer SQL versions.
' 08/17/2012 - SBakker
'            - Changed "Dim cmd..." to "Using cmd..." so it will be properly disposed.
' 02/09/2012 - SBakker - Bug 2-37
'            - Added ExactClone(Obj) routine, to create an exact copy of all fields, "safe"
'              and "unsafe".
'            - Added comment to Clone and ExactClone, to identify which direction the copy
'              is going. (I kept forgetting.)
' 05/18/2011 - SBakker
'            - Remove spaces after comments in generated classes.
' 05/05/2011 - SBakker
'            - Added Partial Private Sub Changed_###() so that other fields can be updated
'              as needed when a value changes.
' 04/26/2011 - SBakker
'            - Added IsFillingFields flag to prevent IsChanged = True while filling fields.
' 04/15/2011 - SBakker
'            - Properly handle Integer Null and NotNull.
' 04/11/2011 - SBakker
'            - Only use one vbCrLf in error messages.
' 03/18/2011 - SBakker
'            - Added dc.CloseConnection(cn$ConnName$) after every dc.GetConnection_$$$, when
'              the connection is just about to fall out of scope. dc.CloseConnection() does
'              a cn$ConnName$.Dispose(), to avoid eating up memory in the application.
'            - Use ".GetValueOrDefault" instead of "Not .HasValue OrElse .Value" clauses.
'              Much easier to read!
' 12/06/2010 - SBakker
'            - If a string = "", don't save it as NULL.
' 11/18/2010 - SBakker
'            - Standardized error messages for easier debugging.
'            - Changed ObjName/FuncName to get the values from System.Reflection.MethodBase
'              instead of hardcoding them.
' 11/17/2010 - SBakker
'            - Removed Clipboard use - started throwing errors. Not sure why.
' 10/18/2010 - SBakker
'            - Added partial sub FixDefault_###(), so default values can be adjusted if they
'              should be different than the default for the data type. Example: Set default
'              Taxable Percent to 100, not 0.
' 09/08/2010 - SBakker
'            - Added a new Clone routine. This allows a separate object of this or a derived
'              class to become an exact duplicate of this object, for all fields in common.
' 09/01/2010 - SBakker
'            - Made "Claim_Master" class be a public class again. Need direct access from
'              other places beyond LTDClaimMaster.
' 07/19/2010 - SBakker
'            - Check if the files are the same before checking if the output file
'              is read-only.
' 07/14/2010 - SBakker
'            - Fixed Strings and Nullable(Of Boolean) to have "If ^^^...", so that
'              the "^^^" will be replaced with the "Is Nothing" logic.
' 06/29/2010 - SBakker
'            - Don't write out a new file if it is exactly the same as the old
'              file. The time/date stamps shouldn't be updated if not changed.
' 06/28/2010 - SBakker
'            - Fixed code for strings. Was checking for IsNot Nothing OrElse "",
'              and it should have just been using AndAlso. Thanks, Muthu!
' 06/25/2010 - SBakker
'            - Fixed the code for Boolean Null. Wasn't ever used until today, and
'              it created code that gave compile errors.
' 05/21/2010 - SBakker
'            - Stop trimming string properties by default. They can be trimmed in
'              their FixValue routines if needed.
' 05/20/2010 - SBakker
'            - Make NotNull "int" fields to have a default of zero, not Nothing.
'              Changed "smallint" and "tinyint" to use the same logic.
' 05/10/2010 - SBakker
'            - Added in extra "Nothing" checks for any string properties that can
'              be "Null". Otherwise it wasn't swapping the property between "" and
'              Nothing.
' 03/31/2010 - SBakker
'            - For VARCHAR(MAX) fields, remove the error, so they can be multiline
'              properties. This will allow them to be whole documents.
' 01/14/2010 - SBakker
'            - Made "Claim_Master" class be a friend class, not a public class.
'            - Added MenuStripMain, File/Exit, and Help/About.
' 01/12/2010 - SBakker
'            - Removed special handling for "Code" and "CodeType" Lookup classes.
' 01/06/2010 - SBakker
'            - Changed internal variable definitions from Private to Protected so
'              they can be overridden in derived classes.
'            - Don't build classes for History tables.
' 09/25/2009 - SBakker
'            - Changed SQLTableName, BaseQuery, FirstConj, and DeleteQuery to be
'              properties with private variables. Also they have partial subs so
'              they can be modified in a *.Part1.vb file, without having to chg
'              the original file.
' 09/23/2009 - SBakker
'            - Added code to check for #BASECLASS# and #IGNORE# directives in
'              the .INFO file associated with the table. This is for classes
'              which are inherited from other classes, not BaseClass.
' 09/04/2009 - SBakker
'            - Added non-shared Save and Delete properties, for simple objects.
'            - Added handling of VARCHAR(MAX) fields.
' 05/05/2009 - SBakker
'            - Removed object name from "Not a multiline property" error so all
'              fields can use the same string.
' 03/18/2009 - SBakker
'            - Added proper processing for "real" data types.
' 03/13/2009 - SBakker
'            - Added partial routines FixValue. This allows some adjusting
'              of values before checking and setting them. A common one would be
'              to always uppercase a string. Strings are already being trimmed.
'            - Added partial routine FixRecordPreSave. This allows fields to be
'              adjusted before the record is saved.
' 03/06/2009 - SBakker
'            - Added support for Null and NotNull strings, and some extra checks
'              for Nothing.
'            - Added CheckValueMore and ValidateMore partial routines, so extra
'              validation can be put into external partial classes.
' 02/24/2009 - SBakker
'            - Added support for calculated fields.
' 02/19/2009 - SBakker
'            - Changed FillFields to use Try/Catch to find out if the field was
'              in the data reader. If not, no error is thrown. The value is
'              either filled with Nothing for Nullable fields, or is not set
'              to any value (i.e. just has the default value for the data type).
'            - Added special handling for "Code" and "CodeType" Lookup classes.
'            - Added read-only check on existing files with the option for them
'              to be fixed and re-saved.
' 02/13/2009 - SBakker
'            - Changed FillFields to use ordinal field numbers rather than
'              string names. These are faster for multiple references.
' 12/17/2008 - SBAKKER - Arena
'            - Removed "Partial" modifier. This will be the one non-partial
'              class, but other partial class definitions can be added as well.
'            - Removed dates in heading comment. This allows easier comparisons.
' 12/04/2008 - SBAKKER - Arena
'            - Add Regions, new Imports, trim strings, and enhance error msgs.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Reflection

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private HasDeletedField As Boolean = False
    Private HasDateTimeField As Boolean = False

    Private AdjustedClassName As String = ""

    Private _BaseClassName As String = "BaseClass"
    Private _ShadowText As String = ""
    Private _TableInfo As String = ""

    Private BaseClassName As String = _BaseClassName
    Private ShadowText As String = _ShadowText
    Private TableInfo As String = _TableInfo
    Private IgnoreFieldList As New List(Of String)

    Private Const MultilineMinLen As Integer = 256

#Region " BlankProperties "

    Private Const BlankIDProp As String =
            "#Region "" Property ### (Int NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Integer) = Nothing" + vbCrLf +
            "    Protected _### As Nullable(Of Integer) = _###_Default" + vbCrLf +
            "    Protected _###_Min As Integer = 1" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Nullable(Of Integer)" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Nullable(Of Integer))" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of Integer))" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        If Value.HasValue Then" + vbCrLf +
            "            If Value.Value < _###_Min Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Value out of range: "" + Value.Value.ToString)" + vbCrLf +
            "            End If" + vbCrLf +
            "        End If" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankIDPropNotNull As String =
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankIDPropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Nullable(Of Integer))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Nullable(Of Integer))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankIntPropHeadNotNull As String =
            "#Region "" Property ### (Int NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Integer = 0" + vbCrLf
    Private Const BlankIntPropHeadNull As String =
            "#Region "" Property ### (Int Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Integer) = Nothing" + vbCrLf
    Private Const BlankIntProp As String =
            "    Protected _### As Integer = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Integer" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Integer)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf
    Private Const BlankIntPropNull As String =
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf
    Private Const BlankIntPropNotNull As String =
            "            If ### <> value Then" + vbCrLf
    Private Const BlankIntPropEnd As String =
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Integer)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf +
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Integer)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Integer)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankStringPropHeadNotNull As String =
            "#Region "" Property ### (String ??? NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As String = """"" + vbCrLf

    Private Const BlankStringPropHeadNull As String =
            "#Region "" Property ### (String ??? Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As String = Nothing" + vbCrLf

    Private Const BlankStringPropNotNullPart1 As String =
            "    Protected _### As String = _###_Default" + vbCrLf +
            "    Protected _###_Max As Integer = ???" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As String" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As String)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If _### <> value Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###, value)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As String)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Default must be valid ---" + vbCrLf +
            "        If Value = _###_Default Then Exit Sub" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        If Value IsNot Nothing Then" + vbCrLf +
            "            If Value.Length > _###_Max Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Invalid length: "" + Value.Length.ToString)" + vbCrLf +
            "            End If" + vbCrLf

    '' ### Implement later ###
    ''Private Const BlankStringPropNotMultiline As String = _
    ''        "            If Value.IndexOf(vbCr) >= 0 OrElse Value.IndexOf(vbLf) >= 0 Then" + vbCrLf +
    ''        "                Throw New SystemException(FuncName + vbCrLf + ""Not a multiline property"")" + vbCrLf +
    ''        "            End If" + vbCrLf

    Private Const BlankStringPropNotMultiline As String =
            "            If Value.IndexOf(vbLf) >= 0 OrElse Value.IndexOf(vbCr) >= 0 Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Not a multiline property"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringPropNotNullPart2 As String =
            "        End If" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf +
            "        If Obj.### Is Nothing Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf

    Private Const BlankStringPropNullPart1 As String =
            "    Protected _### As String = _###_Default" + vbCrLf +
            "    Protected _###_Max As Integer = ???" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As String" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As String)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If EmptyStringToNothing(_###) <> EmptyStringToNothing(value) Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###, value)" + vbCrLf +
            "                _### = EmptyStringToNothing(value)" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As String)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Default must be valid ---" + vbCrLf +
            "        If Value = _###_Default Then Exit Sub" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        If Value IsNot Nothing Then" + vbCrLf +
            "            If Value.Length > _###_Max Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Invalid length: "" + Value.Length.ToString)" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringPropNullPart2 As String =
            "        End If" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf

    Private Const BlankStringMaxProp As String =
            "    Protected _### As String = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As String" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As String)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            If value IsNot Nothing Then" + vbCrLf +
            "                If value.StartsWith("" "") OrElse value.EndsWith("" "") Then" + vbCrLf +
            "                    value = value.Trim" + vbCrLf +
            "                End If" + vbCrLf +
            "            End If" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If ^^^_### <> value Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###, value)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As String)" + vbCrLf +
            "        ' --- Default must be valid ---" + vbCrLf +
            "        If Value = _###_Default Then Exit Sub" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf

    Private Const BlankStringPropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As String)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As String)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankBooleanPropHeadNotNull As String =
            "#Region "" Property ### (Boolean NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Boolean = False" + vbCrLf

    Private Const BlankBooleanPropHeadNull As String =
            "#Region "" Property ### (Boolean Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Boolean) = Nothing" + vbCrLf

    Private Const BlankBooleanProp As String =
            "    Protected _### As Boolean = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Boolean" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Boolean)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If ^^^_### <> value Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Boolean)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf +
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Boolean)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Boolean)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDecimalPropHeadNull As String =
            "#Region "" Property ### (Decimal Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Decimal) = Nothing" + vbCrLf +
            "    Protected _### As Nullable(Of Decimal) = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Nullable(Of Decimal)" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Nullable(Of Decimal))" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf
    Private Const BlankDecimalPropHeadNotNull As String =
            "#Region "" Property ### (Decimal NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Decimal = 0" + vbCrLf +
            "    Protected _### As Decimal = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Decimal" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Decimal)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If _### <> value Then" + vbCrLf
    Private Const BlankDecimalProp As String =
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Decimal)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf +
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Decimal)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Decimal)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDoubleProp As String =
            "#Region "" Property ### (Double NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Double) = Nothing" + vbCrLf +
            "    Protected _### As Nullable(Of Double) = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Nullable(Of Double)" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Nullable(Of Double))" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of Double))" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDoublePropNotNull As String =
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankDoublePropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Nullable(Of Double))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Nullable(Of Double))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankSingleProp As String =
            "#Region "" Property ### (Single NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Single) = Nothing" + vbCrLf +
            "    Protected _### As Nullable(Of Single) = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Nullable(Of Single)" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Nullable(Of Single))" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of Single))" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankSinglePropNotNull As String =
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankSinglePropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Nullable(Of Single))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Nullable(Of Single))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDateProp As String =
            "#Region "" Property ### (DateTime NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of DateTime) = Nothing" + vbCrLf +
            "    Protected _### As Nullable(Of DateTime) = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    Public Property ###() As Nullable(Of DateTime)" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Nullable(Of DateTime))" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If (_###.HasValue <> value.HasValue) OrElse _" + vbCrLf +
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf +
            "                If UseDebugMode AndAlso Not IsFillingFields Then RaiseFieldChangedEvent(""###"", _###.ToString, value.ToString)" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "                Changed_###()" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of DateTime))" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDatePropNotNull As String =
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankDatePropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Nullable(Of DateTime))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Nullable(Of DateTime))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub Changed_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankBytesPropHeadNotNull As String =
            "#Region "" Property ### (Byte() ??? NotNull) """ + vbCrLf +
            "" + vbCrLf

    Private Const BlankBytesPropHeadNull As String =
            "#Region "" Property ### (Byte() ??? Null) """ + vbCrLf +
            "" + vbCrLf

    Private Const BlankBytesPropEnd As String =
            "    End Sub" + vbCrLf +
            "" + vbCrLf

#End Region

#Region " Blank Templates "

    Private Const BlankFillFieldsDeclareOrdinal As String =
            "        Static _FieldNum_### As Integer" + vbCrLf

    Private Const BlankFillFieldsNull As String =
            "        ' --- ### ---" + vbCrLf +
            "        Static _FieldNum_### As Integer = -1 ' not set yet" + vbCrLf +
            "        If _FieldNum_### = -1 Then" + vbCrLf +
            "            _FieldNum_### = dr.GetOrdinal(""##ORIG##"")" + vbCrLf +
            "        End If" + vbCrLf +
            "        If dr.IsDBNull(_FieldNum_###) Then" + vbCrLf +
            "            Obj.### = Nothing" + vbCrLf +
            "        Else" + vbCrLf +
            "            Obj.### = dr.@@@(_FieldNum_###)" + vbCrLf +
            "        End If" + vbCrLf

    Private Const BlankFillFieldsNotNull As String =
            "        ' --- ### ---" + vbCrLf +
            "        Static _FieldNum_### As Integer = -1 ' not set yet" + vbCrLf +
            "        If _FieldNum_### = -1 Then" + vbCrLf +
            "            _FieldNum_### = dr.GetOrdinal(""##ORIG##"")" + vbCrLf +
            "        End If" + vbCrLf +
            "        Obj.### = dr.@@@(_FieldNum_###)" + vbCrLf

    Private Const BlankValidate As String = "        %%%.Validate_###(Obj)" + vbCrLf

    Private Const BlankFieldList As String = "            .Append("",[##ORIG##]"")" + vbCrLf

    ' --- Value constants ---

    Private Const BlankNumericValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("","") : .Append(Me.###.Value.ToString)" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankNumericValueNotNull As String =
            "            .Append("","") : .Append(Me.###.ToString)" + vbCrLf

    Private Const BlankStringValueNull As String =
            "            If Me.### IsNot Nothing Then" + vbCrLf +
            "                .Append("",'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringValueNotNull As String =
            "            .Append("",'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf

    Private Const BlankStringValueV2 As String =
            "            .Append("","") : .Append(StringToSQLNull(Me.###))" + vbCrLf

    Private Const BlankBooleanValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankBooleanValueNotNull As String =
            "            .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankBooleanValueV2 As String =
            "            .Append("","") : .Append(StringToSQLNull(Me.###.ToString))" + vbCrLf

    Private Const BlankDateValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateValueNotNull As String =
            "            .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateValueV2 As String =
            "            .Append("","") : .Append(StringToSQLNull(Me.###.Value.ToString))" + vbCrLf

    Private Const BlankDateTimeValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.Value.ToString(DateTimeMilliFormatPattern)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateTimeValueNotNull As String =
            "            .Append("",'"") : .Append(Me.###.Value.ToString(DateTimeMilliFormatPattern)) : .Append(""'"")" + vbCrLf

    ' --- Update constants ---

    Private Const BlankNumericUpdateNull As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(Me.###.Value.ToString)" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankNumericUpdateNotNull As String =
            "            .Append("",[##ORIG##] = "") : .Append(Me.###.ToString)" + vbCrLf

    Private Const BlankStringUpdateNull As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.### IsNot Nothing Then" + vbCrLf +
            "                .Append(""'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringUpdateNotNull As String =
            "            .Append("",[##ORIG##] = '"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf

    Private Const BlankBooleanUpdateNull As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(""'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankBooleanUpdateNotNull As String =
            "            .Append("",[##ORIG##] = '"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateUpdateNull As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(""'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateUpdateNotNull As String =
            "            .Append("",[##ORIG##] = '"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateTimeUpdateNull As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(""'"") : .Append(Me.###.Value.ToString(DateTimeMilliFormatPattern)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateTimeUpdateNotNull As String =
            "            .Append("",[##ORIG##] = '"") : .Append(Me.###.Value.ToString(DateTimeMilliFormatPattern)) : .Append(""'"")" + vbCrLf

    ' --- SQL constants ---

    Private Const BlankBaseQuery As String = """SELECT * FROM "" + _SQLTableName"
    Private Const BlankFirstConj As String = """ WHERE"""
    Private Const BlankDeleteQuery As String = """DELETE FROM "" + _SQLTableName"

#End Region

#Region " Header/Trailer "

    Private Const HeaderFillFields As String =
            "Public Overloads Shared Sub FillFields(ByVal Obj As %%%, ByVal dr As SqlDataReader)" + vbCrLf +
            "If Obj Is Nothing Then Exit Sub" + vbCrLf
    Private Const TrailerFillFields As String =
            "$BaseClass$.FillFields(Obj, dr)" + vbCrLf +
            "End Sub" + vbCrLf + vbCrLf

    Private Const HeaderValidate As String =
            "Public Overloads Shared Sub Validate(ByVal Obj As %%%)" + vbCrLf +
            "If Obj Is Nothing Then Exit Sub" + vbCrLf
    Private Const TrailerValidate As String =
            "$BaseClass$.Validate(Obj)" + vbCrLf +
            "End Sub" + vbCrLf + vbCrLf

    Private Const HeaderFieldList As String =
            "Protected Overloads Sub FieldList(ByVal sb As StringBuilder)" + vbCrLf +
            "' --- used in the list of fields in an INSERT INTO command ---" + vbCrLf +
            "With sb" + vbCrLf +
            "MyBase.FieldList(sb)" + vbCrLf
    Private Const TrailerFieldList As String =
            "End With" + vbCrLf +
            "End Sub" + vbCrLf + vbCrLf

    Private Const HeaderValueList As String =
            "Protected Overloads Sub ValueList(ByVal sb As StringBuilder)" + vbCrLf +
            "' --- used in the list of values in an INSERT INTO command ---" + vbCrLf +
            "With sb" + vbCrLf +
            "MyBase.ValueList(sb)" + vbCrLf
    Private Const TrailerValueList As String =
            "End With" + vbCrLf +
            "End Sub" + vbCrLf + vbCrLf

    Private Const HeaderUpdateList As String =
            "Protected Overloads Sub UpdateList(ByVal sb As StringBuilder)" + vbCrLf +
            "' --- used in the SET section of an UPDATE command ---" + vbCrLf +
            "With sb" + vbCrLf +
            "MyBase.UpdateList(sb)" + vbCrLf
    Private Const TrailerUpdateList As String =
            "End With" + vbCrLf +
            "End Sub" + vbCrLf + vbCrLf

#End Region

#Region " Form Routines "

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try
            If Arena_Bootstrap.BootstrapClass.CopyProgramsToLaunchPath Then
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Me.Close()
            Exit Sub
        End Try

        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
            My.Settings.Save()
        End If

        ToolStripComboBoxApp.Items.Clear()
        ToolStripComboBoxApp.Items.AddRange(My.Settings.ApplicationList.Split(";"c))

        If ToolStripComboBoxApp.SelectedIndex < 0 Then
            TextFromPath.Text = ""
            TextToPath.Text = ""
        Else
            TextFromPath.Text = CalcFromDirName(ToolStripComboBoxApp.Text)
            TextToPath.Text = CalcToDirName(ToolStripComboBoxApp.Text)
        End If

        TextDate.Text = Format(Now, "MM/dd/yyyy")

        ToolStripComboBoxApp.SelectedItem = My.Settings.LastApp

        For TempDriveIndex As Integer = ToolStripComboBoxDrive.Items.Count - 1 To 0 Step -1
            If Not Directory.Exists(CStr(ToolStripComboBoxDrive.Items(TempDriveIndex)) + "\Arena_Scripts") Then
                ToolStripComboBoxDrive.Items.RemoveAt(TempDriveIndex)
            End If
        Next

        ToolStripComboBoxDrive.SelectedItem = My.Settings.LastDrive

    End Sub

#End Region

    Private Sub LoadTableInfo(ByVal Filename As String)
        ' --- Clear out all .INFO variables ---
        BaseClassName = _BaseClassName
        ShadowText = _ShadowText
        TableInfo = _TableInfo
        IgnoreFieldList.Clear()
        ' --- Load .INFO info ---
        Dim InfoFilename As String = Filename.ToUpper.Replace(".SQL", ".INFO")
        If Not File.Exists(InfoFilename) Then Exit Sub
        Try
            Dim sr As New StreamReader(InfoFilename)
            TableInfo = sr.ReadToEnd
            sr.Close()
        Catch ex As Exception
            Exit Sub
        End Try
        Dim Lines() As String = TableInfo.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each TempLine As String In Lines
            If TempLine.StartsWith("#BASECLASS#") Then
                BaseClassName = TempLine.Substring(11).Trim
                If Not String.Equals(BaseClassName, "BaseClass", StringComparison.OrdinalIgnoreCase) Then
                    ShadowText = " Shadows"
                End If
            ElseIf TempLine.StartsWith("#IGNORE#") Then
                TempLine = TempLine.Substring(8).Trim
                IgnoreFieldList.Add(TempLine)
            End If
        Next
    End Sub

    Private Function CreateProperties() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim CurrType As String
        Dim CurrLen As String
        Dim OutLine As String
        Dim NotNull As Boolean
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If String.IsNullOrEmpty(CurrField) Then Continue For
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated field
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            OutLine = ""
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If CurrType.ToLower = "int" AndAlso CurrField.EndsWith("ID") Then
                        OutLine = BlankIDProp
                        If NotNull Then OutLine += BlankIDPropNotNull
                        OutLine += BlankIDPropEnd
                    ElseIf NotNull Then
                        OutLine = BlankIntPropHeadNotNull
                        OutLine += BlankIntProp
                        OutLine += BlankIntPropNotNull
                        OutLine += BlankIntPropEnd
                    Else
                        OutLine = BlankIntPropHeadNull
                        OutLine += BlankIntProp.Replace("Integer", "Nullable(Of Integer)")
                        OutLine += BlankIntPropNull
                        OutLine += BlankIntPropEnd.Replace("Integer", "Nullable(Of Integer)")
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "sysname"
                    If CurrType.ToLower = "text" Then
                        CurrLen = "MAX"
                    Else
                        CurrLen = CurrLine.Substring(CurrLine.IndexOf("(") + 1)
                        CurrLen = CurrLen.Substring(0, CurrLen.IndexOf(")")).ToUpper ' might be MAX
                    End If
                    If NotNull Then
                        OutLine = BlankStringPropHeadNotNull.Replace("???", CurrLen)
                    Else
                        OutLine = BlankStringPropHeadNull.Replace("???", CurrLen)
                    End If
                    If CurrLen = "MAX" Then
                        OutLine += BlankStringMaxProp.Replace("???", CurrLen)
                    ElseIf NotNull Then
                        OutLine += BlankStringPropNotNullPart1.Replace("???", CurrLen)
                        If Not IsNumeric(CurrLen) OrElse CInt(CurrLen) < MultilineMinLen Then
                            OutLine += BlankStringPropNotMultiline.Replace("???", CurrLen)
                        End If
                        OutLine += BlankStringPropNotNullPart2.Replace("???", CurrLen)
                    Else
                        OutLine += BlankStringPropNullPart1.Replace("???", CurrLen)
                        If Not IsNumeric(CurrLen) OrElse CInt(CurrLen) < MultilineMinLen Then
                            OutLine += BlankStringPropNotMultiline.Replace("???", CurrLen)
                        End If
                        OutLine += BlankStringPropNullPart2.Replace("???", CurrLen)
                    End If
                    OutLine += BlankStringPropEnd
                Case "bit"
                    If NotNull Then
                        OutLine = BlankBooleanPropHeadNotNull
                        OutLine += BlankBooleanProp
                    Else
                        OutLine = BlankBooleanPropHeadNull
                        OutLine += BlankBooleanProp.Replace("Boolean", "Nullable(Of Boolean)")
                    End If
                Case "decimal", "money", "smallmoney"
                    If NotNull Then
                        OutLine = BlankDecimalPropHeadNotNull
                        OutLine += BlankDecimalProp
                    Else
                        OutLine = BlankDecimalPropHeadNull
                        OutLine += BlankDecimalProp.Replace("Decimal", "Nullable(Of Decimal)")
                    End If
                Case "float"
                    OutLine = BlankDoubleProp
                    If NotNull Then OutLine += BlankDoublePropNotNull
                    OutLine += BlankDoublePropEnd
                Case "real"
                    OutLine = BlankSingleProp
                    If NotNull Then OutLine += BlankSinglePropNotNull
                    OutLine += BlankSinglePropEnd
                Case "date", "datetime", "smalldatetime"
                    OutLine = BlankDateProp
                    If NotNull Then OutLine += BlankDatePropNotNull
                    OutLine += BlankDatePropEnd
                Case "uniqueidentifier"
                    CurrLen = "36" ' Length of a uniqueidentifier string
                    If NotNull Then
                        OutLine = BlankStringPropHeadNotNull.Replace("???", CurrLen)
                    Else
                        OutLine = BlankStringPropHeadNull.Replace("???", CurrLen)
                    End If
                    If NotNull Then
                        OutLine += BlankStringPropNotNullPart1.Replace("???", CurrLen)
                        OutLine += BlankStringPropNotMultiline.Replace("???", CurrLen)
                        OutLine += BlankStringPropNotNullPart2.Replace("???", CurrLen)
                    Else
                        OutLine += BlankStringPropNullPart1.Replace("???", CurrLen)
                        OutLine += BlankStringPropNotMultiline.Replace("???", CurrLen)
                        OutLine += BlankStringPropNullPart2.Replace("???", CurrLen)
                    End If
                    OutLine += BlankStringPropEnd
                Case Else
                    MessageBox.Show("Unknown Property Type: " + CurrType)
                    OutLine = "--- Unknown Property Type: " + CurrType + " ---" + vbCrLf + vbCrLf
            End Select
            If Not NotNull Then
                OutLine = OutLine.Replace("NotNull", "Null")
                OutLine = OutLine.Replace("^^^", "(_### Is Nothing) <> (value Is Nothing) OrElse ")
            Else
                OutLine = OutLine.Replace("^^^", "")
            End If
            If Result.Length > 0 Then
                Result.Append(vbCrLf)
            End If
            Result.Append(OutLine.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        Return Result.ToString
    End Function

    Private Function CreateFillFields() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim CurrType As String
        Dim CurrCvt As String
        Dim NotNull As Boolean
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If String.IsNullOrEmpty(CurrField) Then Continue For
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            CurrCvt = "???"
            Select Case CurrType.ToLower
                Case "byte"
                    CurrCvt = "GetByte"
                Case "bigint"
                    CurrCvt = "GetInt64"
                Case "int"
                    CurrCvt = "GetInt32"
                Case "smallint"
                    CurrCvt = "GetInt16"
                Case "tinyint"
                    CurrCvt = "GetByte"
                Case "char", "varchar", "nchar", "nvarchar", "text", "sysname"
                    CurrCvt = "GetString"
                Case "bit"
                    CurrCvt = "GetBoolean"
                Case "decimal", "money", "smallmoney"
                    CurrCvt = "GetDecimal"
                Case "float"
                    CurrCvt = "GetDouble"
                Case "real"
                    CurrCvt = "GetFloat"
                Case "date", "datetime", "smalldatetime"
                    CurrCvt = "GetDateTime"
                Case "uniqueidentifier"
                    CurrCvt = "GetSQLGuid"
            End Select
            If CurrCvt = "???" Then
                Result.Append("### Unknown Type ###" + vbCrLf)
            ElseIf NotNull Then
                Result.Append(BlankFillFieldsNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("@@@", CurrCvt))
            Else
                Result.Append(BlankFillFieldsNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("@@@", CurrCvt))
            End If
        Next
        Return Result.ToString
    End Function

    Private Function CreateValidateList() As String
        Dim Lines() As String
        Dim OrigField As String
        Dim CurrField As String
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrField In Lines
            CurrField = CurrField.Trim
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            Result.Append(BlankValidate.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        Return Result.ToString
    End Function

    Private Function CreateFieldList() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            Result.Append(BlankFieldList.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        Return Result.ToString
    End Function

    Private Function CreateValueList() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim CurrType As String
        Dim NotNull As Boolean
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If String.IsNullOrEmpty(CurrField) Then Continue For
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "sysname"
                    If NotNull Then
                        Result.Append(BlankStringValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankBooleanValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "decimal", "money", "smallmoney"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "date", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "datetime"
                    If NotNull Then
                        Result.Append(BlankDateTimeValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateTimeValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                    HasDateTimeField = True ' Needs "Imports Arena_Utilities.DateUtils"
                Case "uniqueidentifier"
                    If NotNull Then
                        Result.Append(BlankStringValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
        Next
        Return Result.ToString
    End Function

    Private Function CreateUpdateList() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim CurrType As String
        Dim NotNull As Boolean
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If String.IsNullOrEmpty(CurrField) Then Continue For
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "sysname"
                    If NotNull Then
                        Result.Append(BlankStringUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankBooleanUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "decimal", "money", "smallmoney"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "date", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "datetime"
                    If NotNull Then
                        Result.Append(BlankDateTimeUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateTimeUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                    HasDateTimeField = True ' Needs "Imports Arena_Utilities.DateUtils"
                Case "uniqueidentifier"
                    If NotNull Then
                        Result.Append(BlankStringUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
        Next
        Return Result.ToString
    End Function

    Private Function CreateCloneList() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim OrigField As String
        Dim CurrField As String
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
            If String.IsNullOrEmpty(CurrField) Then Continue For
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            ' --- Note that these are copying the internal values directly ---
            Result.Append("        Obj._")
            Result.Append(CurrField)
            Result.Append(" = Me._")
            Result.Append(CurrField)
            Result.Append(vbCrLf)
        Next
        Return Result.ToString
    End Function

    Private Sub DoClearScreen()
        TextInput.Text = ""
        TextDatabaseName.Text = ""
        TextClassName.Text = ""
        TextFields.Text = ""
        TextOutput.Text = ""
        TextInput.Focus()
    End Sub

    Private Function DoBuildClass() As Boolean
        Dim A As Assembly
        Dim SR As StreamReader
        Dim TempResult As String
        Dim sb As New StringBuilder
        Dim Lines() As String
        Dim InFields As Boolean = False
        Dim SchemaName As String = ""
        Dim SchemaLen As Integer = 0
        Dim ClassShort As String = ""
        ' -----------------------------
        HasDeletedField = False
        HasDateTimeField = False
        ' --- split the SQL table definition into Database, TableName, and Fields ---
        If TextClassName.Text = "" Then
            Lines = TextInput.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
            For Each TempLine As String In Lines
                TempLine = TempLine.Replace(vbTab, " ")
                Do While TempLine.Contains("  ")
                    TempLine = TempLine.Replace("  ", " ")
                Loop
                TempLine = TempLine.Trim
                If TempLine.StartsWith("USE [") Then
                    TempLine = TempLine.Substring(5, TempLine.IndexOf("]") - 5)
                    TextDatabaseName.Text = TempLine
                ElseIf TempLine.StartsWith("CREATE TABLE [") Then
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1) ' remove CREATE TABLE
                    SchemaName = TempLine.Substring(0, TempLine.IndexOf("]")) ' get schema
                    If SchemaName.ToLower = "dbo" Then SchemaName = ""
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1) ' remove schema
                    TempLine = TempLine.Substring(0, TempLine.IndexOf("]"))
                    ClassShort = TempLine
                    If SchemaName <> "" Then
                        TextClassName.Text = SchemaName + "_" + ClassShort
                    Else
                        TextClassName.Text = ClassShort
                    End If
                    InFields = True
                ElseIf TempLine.StartsWith("CONSTRAINT") OrElse TempLine.StartsWith(")") Then
                    InFields = False
                ElseIf InFields AndAlso TempLine.StartsWith("[") AndAlso Not TempLine.ToUpper.Contains("] AS ") Then
                    sb.Append(TempLine)
                    sb.Append(vbCrLf)
                End If
            Next
            TextFields.Text = sb.ToString
        End If
        ' --- Fix class name if necessary ---
        AdjustedClassName = FixClassName(TextClassName.Text)
        ' --- Build the class from the known information ---
        A = Assembly.GetExecutingAssembly
        SR = New StreamReader(A.GetManifestResourceStream("ArenaClassBuilder.BlankArenaDataClass.txt"))
        TempResult = SR.ReadToEnd
        SR.Close()
        If TextDatabaseName.Text <> "" Then
            TempResult = TempResult.Replace("$Database$", TextDatabaseName.Text)
        End If
        If SchemaName <> "" Then
            SchemaLen = SchemaName.Length + 1
            TempResult = TempResult.Replace("$SchemaName$", SchemaName + ".")
            TempResult = TempResult.Replace("$SchemaFull$", "[" + SchemaName + "].")
        Else
            SchemaLen = 0
            TempResult = TempResult.Replace("$SchemaName$", "")
            TempResult = TempResult.Replace("$SchemaFull$", "")
        End If
        TempResult = TempResult.Replace("$BaseClass$", BaseClassName)
        TempResult = TempResult.Replace("$Shadows$", ShadowText)
        TempResult = TempResult.Replace("$ClassShort$", ClassShort)
        TempResult = TempResult.Replace("$ClassName$", AdjustedClassName)
        TempResult = TempResult.Replace("$---$", StrDup(22 - 13 + TextDatabaseName.Text.Length + TextClassName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$---1---$", StrDup(21 + TextConnName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$MMDDYYYY$", TextDate.Text)
        If TextConnName.Text <> "" Then
            TempResult = TempResult.Replace("$ConnName$", TextConnName.Text)
        End If
        TempResult = TempResult.Replace("$Properties$" + vbCrLf, CreateProperties)
        If HasDeletedField Then
            TempResult = TempResult.Replace("$BaseQuery$", BlankBaseQuery + " + "" WHERE [IsDeleted] = 'False'""")
            TempResult = TempResult.Replace("$FirstConj$", """ AND""")
            TempResult = TempResult.Replace("$DeleteQuery$", """UPDATE "" + SQLTableName + "" SET [IsDeleted] = 'True'""")
        Else
            TempResult = TempResult.Replace("$BaseQuery$", BlankBaseQuery)
            TempResult = TempResult.Replace("$FirstConj$", BlankFirstConj)
            TempResult = TempResult.Replace("$DeleteQuery$", BlankDeleteQuery)
        End If
        TempResult = TempResult.Replace("$FillFields$" + vbCrLf, CreateFillFields)
        TempResult = TempResult.Replace("$ValidateList$" + vbCrLf, CreateValidateList)
        TempResult = TempResult.Replace("$FieldList$" + vbCrLf, CreateFieldList)
        TempResult = TempResult.Replace("$ValueList$" + vbCrLf, CreateValueList)
        TempResult = TempResult.Replace("$UpdateList$" + vbCrLf, CreateUpdateList)
        TempResult = TempResult.Replace("$CloneList$" + vbCrLf, CreateCloneList)
        If Not HasDateTimeField Then
            ' --- Doesn't need this import ---
            TempResult = TempResult.Replace("Imports Arena_Utilities.DateUtils" + vbCrLf, "")
        End If
        ' --- Display on the output screen ---
        TextOutput.Text = TempResult
        TextOutput.Focus()
        TextOutput.SelectAll()
        Return True
    End Function

    Private Sub TextInput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextInput.KeyPress
        If e.KeyChar = Chr(1) Then ' ctrl-a
            e.Handled = True
            TextInput.SelectAll()
        End If
    End Sub

    Private Sub TextDatabaseName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextDatabaseName.KeyPress
        If e.KeyChar = Chr(1) Then ' ctrl-a
            e.Handled = True
            TextDatabaseName.SelectAll()
        End If
    End Sub

    Private Sub TextClassName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextClassName.KeyPress
        If e.KeyChar = Chr(1) Then ' ctrl-a
            e.Handled = True
            TextClassName.SelectAll()
        End If
    End Sub

    Private Sub TextFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextFields.KeyPress
        If e.KeyChar = Chr(1) Then ' ctrl-a
            e.Handled = True
            TextFields.SelectAll()
        End If
    End Sub

    Private Sub TextOutput_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextOutput.KeyPress
        If e.KeyChar = Chr(1) Then ' ctrl-a
            e.Handled = True
            TextOutput.SelectAll()
        End If
    End Sub

    Private Sub ButtonBuildAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBuildAll.Click
        Dim CurrInfo As FileInfo
        Dim Answer As DialogResult
        Dim FullFilePath As String = ""
        Dim FileCount As Integer = 0
        Dim ChangedCount As Integer = 0
        Const FileCountMsg As String = "Files Found: "
        Const ChangedCountMsg As String = " - Files Changed: "
        ' ----------------------------------------------------
        If TextFromPath.Text = "" Then Exit Sub
        If TextToPath.Text = "" Then Exit Sub
        If TextDatabase.Text = "" Then Exit Sub
        If TextConnName.Text = "" Then Exit Sub
        If ToolStripComboBoxApp.SelectedIndex < 0 Then Exit Sub
        If ToolStripComboBoxDrive.SelectedIndex < 0 Then Exit Sub
        If Not Directory.Exists(ToolStripComboBoxDrive.Text + TextFromPath.Text) Then
            MessageBox.Show("FromPath not found: " + ToolStripComboBoxDrive.Text + TextFromPath.Text)
            Exit Sub
        End If
        If Not Directory.Exists(ToolStripComboBoxDrive.Text + TextToPath.Text) Then
            MessageBox.Show("ToPath not found: " + ToolStripComboBoxDrive.Text + TextToPath.Text)
            Exit Sub
        End If
        My.Settings.LastApp = ToolStripComboBoxApp.Text
        My.Settings.LastDrive = ToolStripComboBoxDrive.Text
        My.Settings.Save()
        FileCount = 0
        ChangedCount = 0
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
        Dim FromDir As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(ToolStripComboBoxDrive.Text + TextFromPath.Text)
        Dim FromFiles() As FileInfo = FromDir.GetFiles("*.sql")
        For Each CurrFile As FileInfo In FromFiles
            If CurrFile.Name.ToUpper.IndexOf(".TABLE.") < 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf("_HIST.TABLE.") >= 0 Then Continue For
            FileCount += 1
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            My.Application.DoEvents()
            DoClearScreen()
            TextDatabaseName.Text = TextDatabase.Text
            Dim sr As StreamReader = CurrFile.OpenText
            TextInput.Text = sr.ReadToEnd
            sr.Close()
            LoadTableInfo(ToolStripComboBoxDrive.Text + TextFromPath.Text + "\" + CurrFile.Name)
            If DoBuildClass() Then
                FullFilePath = ToolStripComboBoxDrive.Text + TextToPath.Text + "\" + CurrFile.Name.Replace("dbo.", "").Replace(".Table.sql", ".vb")
                ' --- check if the file exists ---
                If File.Exists(FullFilePath) Then
                    ' --- check if the file has changed ---
                    Dim OldSR As New StreamReader(FullFilePath)
                    Dim OldFile As String = OldSR.ReadToEnd
                    OldSR.Close()
                    If OldFile = TextOutput.Text Then
                        Continue For
                    End If
                    ' --- check if the file read-only ---
                    CurrInfo = My.Computer.FileSystem.GetFileInfo(FullFilePath)
                    Do While CurrInfo.IsReadOnly
                        Answer = MessageBox.Show("""" + FullFilePath + """ is Read-Only",
                                                 "File is Read-Only", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                        ' --- abort ---
                        If Answer = DialogResult.Abort Then
                            Exit Sub
                        End If
                        ' --- ignore ---
                        If Answer = DialogResult.Ignore Then
                            Continue For
                        End If
                        ' --- retry ---
                        CurrInfo = My.Computer.FileSystem.GetFileInfo(FullFilePath)
                    Loop
                End If
                ' --- output the result ---
                Dim sw As New StreamWriter(FullFilePath)
                sw.Write(TextOutput.Text)
                sw.Close()
                ChangedCount += 1
                ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            End If
        Next
        ToolStripStatusLabelMain.Text += " - Done"
    End Sub

    Private Function IgnoreField(ByVal CurrField As String) As Boolean
        If String.Equals(CurrField, "ID", StringComparison.OrdinalIgnoreCase) Then Return True
        If String.Equals(CurrField, "ROWVERSION", StringComparison.OrdinalIgnoreCase) Then Return True
        If String.Equals(CurrField, "LASTCHANGED", StringComparison.OrdinalIgnoreCase) Then Return True
        If String.Equals(CurrField, "CHANGEDBY", StringComparison.OrdinalIgnoreCase) Then Return True
        If String.Equals(CurrField, "DELETED", StringComparison.OrdinalIgnoreCase) Then
            HasDeletedField = True
            Return True
        End If
        If String.Equals(CurrField, "ISDELETED", StringComparison.OrdinalIgnoreCase) Then
            HasDeletedField = True
            Return True
        End If
        For Each TempIgnoreField As String In IgnoreFieldList
            If String.Equals(CurrField, TempIgnoreField, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
    End Sub

    Private Sub ToolStripComboBoxApp_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBoxApp.SelectedIndexChanged
        With ToolStripComboBoxApp
            If .SelectedIndex < 0 Then Exit Sub
            If ToolStripComboBoxApp.SelectedIndex < 0 Then
                TextFromPath.Text = ""
                TextToPath.Text = ""
                TextDatabase.Text = ""
                TextConnName.Text = ""
                ButtonBuildAll.Enabled = False
            Else
                TextFromPath.Text = CalcFromDirName(ToolStripComboBoxApp.Text)
                TextToPath.Text = CalcToDirName(ToolStripComboBoxApp.Text)
                TextDatabase.Text = CStr(.Items(.SelectedIndex))
                TextConnName.Text = CStr(.Items(.SelectedIndex))
                ButtonBuildAll.Enabled = True
            End If
            My.Settings.LastApp = CStr(.Items(.SelectedIndex))
            My.Settings.Save()
        End With
        TextInput.Text = ""
        TextDatabaseName.Text = ""
        TextClassName.Text = ""
        TextFields.Text = ""
        TextOutput.Text = ""
    End Sub

    Private Function FixFieldName(ByVal FieldName As String) As String
        ' --- Check for spaces in property name ---
        If FieldName.Contains(" ") Then
            FieldName = FieldName.Replace(" ", "_")
        End If
        ' --- Check for invalid field names ---
        If FieldName.StartsWith("100") Then
            FieldName = FieldName.Replace("100", "Hundred")
        End If
        Return FieldName
    End Function

    Private Function FixClassName(ByVal ClassName As String) As String
        ' --- Remove spaces from class names ---
        If ClassName.Contains(" ") Then
            ClassName = ClassName.Replace(" ", "_")
        End If
        ' --- Adjust Advantage class names ---
        If TextDatabase.Text.ToUpper = "ADVANTAGE" Then
            If ClassName.ToUpper = "CODES" Then
                ClassName = "Advantage_" + ClassName
            End If
        End If
        ' --- Some IDR/NoteDiary class names are reserved words ---
        If TextDatabase.Text.ToUpper = "IDR" Then
            ClassName = "IDR_" + ClassName
        End If
        If TextDatabase.Text.ToUpper = "NOTEDIARY" Then
            ClassName = "NoteDiary_" + ClassName
        End If
        ' --- Some IDRIS class names conflict with Arena class names ---
        If TextDatabase.Text.ToUpper = "IDRIS" Then
            If ClassName.ToUpper.StartsWith("POLICY") Then
                ClassName = "IDRIS_" + ClassName
            End If
        End If
        Return ClassName
    End Function

    Private Function CalcFromDirName(ByVal AppName As String) As String
        If AppName = "IDRIS" Then
            Return My.Settings.BaseFromPath + " " + ToolStripComboBoxApp.Text + " Arena"
        Else
            Return My.Settings.BaseFromPath + " " + ToolStripComboBoxApp.Text
        End If
    End Function

    Private Function CalcToDirName(ByVal AppName As String) As String
        If AppName = "IDRIS" Then
            Return My.Settings.BaseToPath + " " + ToolStripComboBoxApp.Text + " Arena"
        Else
            Return My.Settings.BaseToPath + " " + ToolStripComboBoxApp.Text
        End If
    End Function

End Class
