' ------------------------------------------
' --- AdvantageClassBuilder - 02/02/2016 ---
' ------------------------------------------

' ----------------------------------------------------------------------------------------------------
' 02/02/2016 - SBakker
'            - Added soft errors for Path Not Found.
' 10/28/2015 - SBakker - URD 12527
'            - Added code for setting "cmd.CommandTimeout = DataConnection.SQLTimeoutSeconds"
'              everywhere that a SQLCommand object is created. This might help with the infrequent
'              timeout errors.
' 03/31/2015 - SBakker
'            - Renamed "My.Settings.Applications" to "My.Settings.ApplicationList". Apparently you
'              can't change a User setting to an Application setting and have it work properly without
'              renaming it.
' 03/11/2015 - SBakker
'            - Changed "My.Settings.Applications" setting to be an application level setting.
' 01/23/2015 - SBakker
'            - Check for either vbCr or vbLf to determine Multiline. ### Will implement later ###
' 06/16/2014 - SBakker
'            - Fixed handling of IDRIS "%" or "_" tables.
'            - Use "##ORIG##" for original field name before fixing. This allows filling and saving
'              using the SQL field name, instead of the fixed property name.
' 02/25/2014 - SBakker
'            - Added handling of "ntext".
'            - Added UserRequest to list of applications.
' 02/24/2014 - SBakker
'            - Added Bootstrap loading all programs to another location, and then running from there.
' 01/10/2014 - SBakker
'            - Added routines "Public Sub Save(ByVal dc As DataConnection)" so an object can do a
'              "MyObj.Save(dc)". Same for "Delete()". Arena already had this, but Advantage didn't.
'              Great for transactions!
'            - Must delete record in Save() routine if applicable before calling FixRecordPreSave(),
'              or the number of BeginTransactions/EndTransactions could become mismatched.
' 12/13/2013 - SBakker
'            - Added a ChangedCount to show how many classes were updated.
' 10/22/2013 - SBakker
'            - Allow any CHAR/VARCHAR fields with MultilineMinLen or greater length to be multiline.
' 10/04/2013 - SBakker
'            - Strip the Drive from all FromPath/ToPath strings upon loading. Makes it easier so you
'              don't have to do it yourself.
'            - Only enable ButtonBuildAll when something is selected.
' 10/01/2013 - SBakker
'            - Added Drive Combo Box and StatusStrip.
' 09/17/2013 - SBakker
'            - Added "Partial Private Shared Sub ValidateRecordLevel()". This allows Record level
'              validation to be added in a partial class, as needed.
'            - Added additional error information.
' 08/23/2013 - SBakker
'            - Made "[INT] NOT NULL" fields actually use "Integer" instead of "Nullable(Of Integer)".
'            - Check for fields that end in "Id" instead of "ID", and any Identity fields.
' 08/19/2013 - SBakker
'            - Added "smalldatetime" support.
'            - Added "IDRIS_Extracts" to database list.
' 07/17/2013 - SBakker
'            - Added dropdown list to handle "IDRIS Advantage" classes, such as "Policy".
'            - Added FixFieldName() to handle any invalid field names, such as "100PctRate".
'            - Added FixClassName() to fix NoteDiary (from below), and also IDRIS class names that
'              would conflict with Arena class names.
' 07/08/2013 - SBakker
'            - Added dropdown list to handle IDR.
'            - Some IDR class names are duplicates, so add "IDR_" at their beginning.
'            - Replace spaces in IDR class names and property names with "_".
' 06/20/2013 - SBakker
'            - Added IdentityField to IgnoreField(). Made it a class-level variable.
'            - Don't exclude IdentityField in CreateProperties or CreateFillFields because
'              there is no BaseClass on Advantage classes.
' 06/12/2013 - SBakker
'            - Added Menu Bar and dropdown list of databases, to handle both "Advantage" and
'              "NoteDiary".
'            - Added a "NoIdentity" script for tables with no identity fields. The records
'              can only be read and not saved or deleted.
'            - Added a check to rename the Advantage "Codes" class to be "Advantage_Codes",
'              so that it doesn't conflict with the IDRIS "CODES" class.
' 05/31/2013 - SBakker
'            - Some NoteDiary class names are reserved words, so add "NoteDiary_" at their
'              beginning.
' 05/28/2013 - SBakker
'            - Added the ability to handle "text" and "varchar(max)" fields.
' 08/17/2012 - SBakker
'            - Changed "Dim cmd..." to "Using cmd..." so it will be properly disposed.
' 05/18/2011 - SBakker
'            - Remove spaces after comments in generated classes.
' 04/26/2011 - SBakker
'            - Added IsFillingFields flag to prevent IsChanged = True while filling fields.
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
'            - Changed internal variable definitions from Private to Protected so
'              they can be overridden in derived classes.
' 05/10/2010 - SBakker
'            - Added in extra "Nothing" checks for any string properties that can
'              be "Null". Otherwise it wasn't swapping the property between "" and
'              Nothing.
' 09/25/2009 - SBakker
'            - Changed SQLTableName, BaseQuery, FirstConj, and DeleteQuery to be
'              properties with private variables. Also they have partial subs so
'              they can be modified in a *.Part1.vb file, without having to chg
'              the original file.
' 05/05/2009 - SBakker
'            - Removed object name from "Not a multiline property" error so all
'              fields can use the same string.
' 03/18/2009 - SBakker
'            - Added proper processing for "real" data types.
'            - Properly set internal flags at end of FillFields.
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
' 12/17/2008 - SBAKKER - Advantage
'            - Removed "Partial" modifier. This will be the one non-partial
'              class, but other partial class definitions can be added as well.
'            - Removed dates in heading comment. This allows easier comparisons.
' 12/04/2008 - SBAKKER - Advantage
'            - Add Regions, new Imports, trim strings, and enhance error msgs.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Reflection

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private IdentityField As String
    Private HasDeletedField As Boolean = False

    Private AdjustedClassName As String = ""

    Private FromPaths() As String
    Private ToPaths() As String

    Private Const MultilineMinLen As Integer = 256

#Region " BlankProperties "

    Private Const BlankIDProp As String = _
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
    Private Const BlankIDPropNotNull As String = _
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankIDPropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankIntPropHeadNotNull As String = _
            "#Region "" Property ### (Int NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Integer = 0" + vbCrLf

    Private Const BlankIntPropHeadNull As String = _
            "#Region "" Property ### (Int Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Integer) = Nothing" + vbCrLf

    Private Const BlankIntPropNull As String = _
            "    Protected _### As Nullable(Of Integer) = _###_Default" + vbCrLf +
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of Integer))" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf

    Private Const BlankIntPropNotNull As String = _
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
            "            CheckValue_###(value)" + vbCrLf +
            "            If _### <> value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf

    Private Const BlankIntPropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankStringPropHeadNotNull As String = _
            "#Region "" Property ### (String ??? NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As String = """"" + vbCrLf

    Private Const BlankStringPropHeadNull As String = _
            "#Region "" Property ### (String ??? Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As String = Nothing" + vbCrLf

    '' ### Implement later ###
    ''If Value.IndexOf(vbCr) >= 0 OrElse Value.IndexOf(vbLf) >= 0 Then

    Private Const BlankStringProp As String = _
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
            "            If ^^^_### <> value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
            "            End If" + vbCrLf +
            "            If Value.IndexOf(vbLf) >= 0 Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Not a multiline property"")" + vbCrLf +
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

    Private Const BlankStringPropMultiline As String = _
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
            "            If ^^^_### <> value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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

    Private Const BlankStringMaxProp As String = _
            "    Protected _### As String = _###_Default" + vbCrLf +
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
            "            If ^^^_### <> value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As String)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
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

    Private Const BlankStringPropNotNull As String = _
            "        If Obj.### Is Nothing Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankStringPropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankBooleanPropHeadNotNull As String = _
            "#Region "" Property ### (Boolean NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Boolean = False" + vbCrLf

    Private Const BlankBooleanPropHeadNull As String = _
            "#Region "" Property ### (Boolean Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Protected _###_Default As Nullable(Of Boolean) = Nothing" + vbCrLf

    Private Const BlankBooleanProp As String = _
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDecimalProp As String = _
            "#Region "" Property ### (Decimal NotNull) """ + vbCrLf +
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
            "               (_###.HasValue AndAlso _###.Value <> value.Value) Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Nullable(Of Decimal))" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDecimalPropNotNull As String = _
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankDecimalPropEnd As String = _
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Nullable(Of Decimal))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Nullable(Of Decimal))" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDoubleProp As String = _
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
    Private Const BlankDoublePropNotNull As String = _
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankDoublePropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankSingleProp As String = _
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
    Private Const BlankSinglePropNotNull As String = _
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankSinglePropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDateProp As String = _
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
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
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
    Private Const BlankDatePropNotNull As String = _
            "        If Not Obj.###.HasValue Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankDatePropEnd As String = _
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

#End Region

#Region " Blank Templates "

    Private Const AppendComma As String = ".Append("","") : "

    Private Const BlankFillFieldsNull As String = _
            "        ' --- ### ---" + vbCrLf +
            "        Static _FieldNum_### As Integer = -99 ' not set yet" + vbCrLf +
            "        If _FieldNum_### = -99 Then" + vbCrLf +
            "            Try" + vbCrLf +
            "                _FieldNum_### = dr.GetOrdinal(""##ORIG##"")" + vbCrLf +
            "            Catch ex As Exception" + vbCrLf +
            "                _FieldNum_### = -1 ' not found" + vbCrLf +
            "            End Try" + vbCrLf +
            "        End If" + vbCrLf +
            "        If _FieldNum_### < 0 OrElse dr.IsDBNull(_FieldNum_###) Then" + vbCrLf +
            "            Obj.### = Nothing" + vbCrLf +
            "        Else" + vbCrLf +
            "            Obj.### = dr.@@@(_FieldNum_###)" + vbCrLf +
            "        End If" + vbCrLf

    Private Const BlankFillFieldsNotNull As String = _
            "        ' --- ### ---" + vbCrLf +
            "        Static _FieldNum_### As Integer = -99 ' not set yet" + vbCrLf +
            "        If _FieldNum_### = -99 Then" + vbCrLf +
            "            Try" + vbCrLf +
            "                _FieldNum_### = dr.GetOrdinal(""##ORIG##"")" + vbCrLf +
            "            Catch ex As Exception" + vbCrLf +
            "                _FieldNum_### = -1 ' not found" + vbCrLf +
            "            End Try" + vbCrLf +
            "        End If" + vbCrLf +
            "        If _FieldNum_### >= 0 Then" + vbCrLf +
            "            Obj.### = dr.@@@(_FieldNum_###)" + vbCrLf +
            "        End If" + vbCrLf

    Private Const BlankValidate As String = "        %%%.Validate_###(Obj)" + vbCrLf

    Private Const BlankFieldList As String = "            $Comma$.Append(""[##ORIG##]"")" + vbCrLf

    ' --- Value constants ---

    Private Const BlankNumericValueNull As String = _
            "            If Me.###.HasValue Then" + vbCrLf +
            "                $Comma$.Append(Me.###.Value.ToString)" + vbCrLf +
            "            Else" + vbCrLf +
            "                $Comma$.Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankNumericValueNotNull As String = _
            "            $Comma$.Append(Me.###.ToString)" + vbCrLf

    Private Const BlankStringValueNull As String = _
            "            If Me.### IsNot Nothing Then" + vbCrLf +
            "                $Comma$.Append(""'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                $Comma$.Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringValueNotNull As String = _
            "            $Comma$.Append(""'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf

    Private Const BlankBooleanValueNull As String = _
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankBooleanValueNotNull As String = _
            "            $Comma$.Append(""'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateValueNull As String = _
            "            If Me.###.HasValue Then" + vbCrLf +
            "                $Comma$.Append(""'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                $Comma$.Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateValueNotNull As String = _
            "            $Comma$.Append(""'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    ' --- Update constants ---

    Private Const BlankNumericUpdateNull As String = _
            "            $Comma$.Append(""[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(Me.###.Value.ToString)" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankNumericUpdateNotNull As String = _
            "            $Comma$.Append(""[##ORIG##] = "") : .Append(Me.###.ToString)" + vbCrLf

    Private Const BlankStringUpdateNull As String = _
            "            $Comma$.Append(""[##ORIG##] = "")" + vbCrLf +
            "            If Me.### IsNot Nothing Then" + vbCrLf +
            "                .Append(""'"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankStringUpdateNotNull As String = _
            "            $Comma$.Append(""[##ORIG##] = '"") : .Append(StringToSQL(Me.###)) : .Append(""'"")" + vbCrLf

    Private Const BlankBooleanUpdateNull As String = _
            "            $Comma$.Append(""[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(""'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankBooleanUpdateNotNull As String = _
            "            $Comma$.Append(""[##ORIG##] = '"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateUpdateNull As String = _
            "            $Comma$.Append(""[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(""'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateUpdateNotNull As String = _
            "            $Comma$.Append(""[##ORIG##] = '"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    ' --- SQL constants ---

    Private Const BlankBaseQuery As String = """SELECT * FROM "" + _SQLTableName"
    Private Const BlankFirstConj As String = """ WHERE"""
    Private Const BlankDeleteQuery As String = """DELETE FROM "" + _SQLTableName"

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

        ToolStripComboBoxApp.Items.AddRange(My.Settings.ApplicationList.Split(";"c))
        FromPaths = My.Settings.LastFromPath.Split(";"c)
        ToPaths = My.Settings.LastToPath.Split(";"c)

        For CurrIndex As Integer = 0 To FromPaths.Count - 1
            If FromPaths(CurrIndex).Length >= 2 AndAlso FromPaths(CurrIndex).Substring(1, 1) = ":" Then
                FromPaths(CurrIndex) = FromPaths(CurrIndex).Substring(2)
            End If
        Next

        For CurrIndex As Integer = 0 To ToPaths.Count - 1
            If ToPaths(CurrIndex).Length >= 2 AndAlso ToPaths(CurrIndex).Substring(1, 1) = ":" Then
                ToPaths(CurrIndex) = ToPaths(CurrIndex).Substring(2)
            End If
        Next

        If FromPaths.GetUpperBound(0) <> ToolStripComboBoxApp.Items.Count - 1 Then
            ReDim Preserve FromPaths(ToolStripComboBoxApp.Items.Count - 1)
            ReDim Preserve ToPaths(ToolStripComboBoxApp.Items.Count - 1)
        End If

        If ToolStripComboBoxApp.SelectedIndex < 0 Then
            TextFromPath.Text = ""
            TextToPath.Text = ""
        Else
            TextFromPath.Text = FromPaths(ToolStripComboBoxApp.SelectedIndex)
            TextToPath.Text = ToPaths(ToolStripComboBoxApp.SelectedIndex)
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
            ' --- Include IdentityField here because there is no BaseClass ---
            If CurrField <> IdentityField Then
                If IgnoreField(CurrField) Then Continue For
            End If
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If CurrField = IdentityField OrElse CurrField.ToUpper.EndsWith("ID") Then
                        OutLine = BlankIDProp
                        If NotNull Then OutLine += BlankIDPropNotNull
                        OutLine += BlankIDPropEnd
                    ElseIf NotNull Then
                        OutLine = BlankIntPropHeadNotNull
                        OutLine += BlankIntPropNotNull
                        OutLine += BlankIntPropEnd.Replace("Nullable(Of Integer)", "Integer")
                    Else
                        OutLine = BlankIntPropHeadNull
                        OutLine += BlankIntPropNull
                        OutLine += BlankIntPropEnd
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "ntext"
                    If CurrType.ToLower = "text" OrElse CurrType.ToLower = "ntext" Then
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
                    ElseIf IsNumeric(CurrLen) AndAlso CInt(CurrLen) >= MultilineMinLen Then
                        OutLine += BlankStringPropMultiline.Replace("???", CurrLen)
                    Else
                        OutLine += BlankStringProp.Replace("???", CurrLen)
                    End If
                    If NotNull Then OutLine += BlankStringPropNotNull
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
                    OutLine = BlankDecimalProp
                    If NotNull Then OutLine += BlankDecimalPropNotNull
                    OutLine += BlankDecimalPropEnd
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
            ' --- Include IdentityField here because there is no BaseClass ---
            If CurrField <> IdentityField Then
                If IgnoreField(CurrField) Then Continue For
            End If
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
                Case "char", "varchar", "nchar", "nvarchar"
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
                Case "text", "ntext"
                    CurrCvt = "GetString"
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
        Dim Comma As String = ""
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
            Result.Append(BlankFieldList.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
            If Comma <> AppendComma Then
                Comma = AppendComma
            End If
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
        Dim Comma As String = ""
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
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
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "ntext"
                    If NotNull Then
                        Result.Append(BlankStringValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankStringValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankBooleanValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "decimal", "money", "smallmoney"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "date", "datetime", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateValueNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankDateValueNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
            If Comma <> AppendComma Then
                Comma = AppendComma
            End If
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
        Dim Comma As String = ""
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim
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
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text", "ntext"
                    If NotNull Then
                        Result.Append(BlankStringUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankStringUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankBooleanUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "decimal", "money", "smallmoney"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case "date", "datetime", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateUpdateNotNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    Else
                        Result.Append(BlankDateUpdateNull.Replace("%%%", AdjustedClassName).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("$Comma$", Comma))
                    End If
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
            If Comma <> AppendComma Then
                Comma = AppendComma
            End If
        Next
        Return Result.ToString
    End Function

    Private Sub ButtonClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DoClearScreen()
    End Sub

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
        IdentityField = "???"
        ' --- split the SQL table definition into Database, TableName, and Fields ---
        If TextClassName.Text = "" Then
            Lines = TextInput.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
            For Each TempLine As String In Lines
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
                ElseIf InFields AndAlso TempLine.StartsWith("[") Then
                    sb.Append(TempLine)
                    sb.Append(vbCrLf)
                    If TempLine.ToUpper.IndexOf("IDENTITY") >= 0 Then
                        IdentityField = TempLine.Substring(1, TempLine.IndexOf("]") - 1)
                    End If
                End If
            Next
            TextFields.Text = sb.ToString
        End If
        ' --- Fix class name if necessary ---
        AdjustedClassName = FixClassName(TextClassName.Text)
        ' --- Build the class from the known information ---
        A = Assembly.GetExecutingAssembly
        Dim BlankClassFileName As String
        If IdentityField = "???" Then
            BlankClassFileName = "BlankAdvantageNoIdentity.txt"
        Else
            BlankClassFileName = "BlankAdvantageDataClass.txt"
        End If
        SR = New StreamReader(A.GetManifestResourceStream("AdvantageClassBuilder." + BlankClassFileName))
        If SR Is Nothing Then
            MessageBox.Show("Error loading file: " + BlankClassFileName)
            Return False
        End If
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
        TempResult = TempResult.Replace("$ClassShort$", ClassShort)
        TempResult = TempResult.Replace("$ClassName$", AdjustedClassName)
        TempResult = TempResult.Replace("$---$", StrDup(22 - 13 + TextDatabaseName.Text.Length + TextClassName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$---1---$", StrDup(21 + TextConnName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$MMDDYYYY$", TextDate.Text)
        TempResult = TempResult.Replace("$Identity$", IdentityField)
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
        With ToolStripComboBoxApp
            FromPaths(.SelectedIndex) = TextFromPath.Text
            ToPaths(.SelectedIndex) = TextToPath.Text
        End With
        Dim TempFromPath As String = ""
        For Each TempItem As String In FromPaths
            If TempFromPath <> "" Then TempFromPath += ";"
            TempFromPath += TempItem
        Next
        Dim TempToPath As String = ""
        For Each TempItem As String In ToPaths
            If TempToPath <> "" Then TempToPath += ";"
            TempToPath += TempItem
        Next
        My.Settings.LastApp = ToolStripComboBoxApp.Text
        My.Settings.LastFromPath = TempFromPath
        My.Settings.LastToPath = TempToPath
        My.Settings.LastDrive = ToolStripComboBoxDrive.Text
        My.Settings.Save()
        FileCount = 0
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
        Dim FromDir As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(ToolStripComboBoxDrive.Text + TextFromPath.Text)
        Dim FromFiles() As FileInfo = FromDir.GetFiles("*.sql")
        For Each CurrFile As FileInfo In FromFiles
            If CurrFile.Name.ToUpper.IndexOf(".TABLE.") < 0 Then Continue For
            FileCount += 1
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            My.Application.DoEvents()
            DoClearScreen()
            TextDatabaseName.Text = TextDatabase.Text
            Dim sr As StreamReader = CurrFile.OpenText
            TextInput.Text = sr.ReadToEnd
            sr.Close()
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
                        Answer = MessageBox.Show("""" + FullFilePath + """ is Read-Only", _
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
        If Not String.IsNullOrWhiteSpace(IdentityField) AndAlso IdentityField <> "???" Then
            If String.Equals(CurrField, IdentityField, StringComparison.OrdinalIgnoreCase) Then Return True
        End If
        If CurrField.ToUpper = "ROWVERSION" Then Return True
        If CurrField.ToUpper = "TIMESTAMP" Then Return True
        Return False
    End Function

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
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
                TextFromPath.Text = FromPaths(.SelectedIndex)
                TextToPath.Text = ToPaths(.SelectedIndex)
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
        If FieldName = "SQLTableName" Then
            FieldName = "SQLTableNameField"
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
        If TextDatabase.Text.ToUpper = "USERREQUEST" Then
            ClassName = "UserRequest_" + ClassName
        End If
        ' --- Some IDRIS class names conflict with Arena class names ---
        If TextDatabase.Text.ToUpper = "IDRIS" Then
            If ClassName.ToUpper.StartsWith("POLICY") Then
                ClassName = "IDRIS_" + ClassName
            ElseIf ClassName.ToUpper.StartsWith("%") OrElse ClassName.ToUpper.StartsWith("_") Then
                ClassName = "IDRIS_" + ClassName.Substring(1)
            End If
        End If
        Return ClassName
    End Function

End Class
