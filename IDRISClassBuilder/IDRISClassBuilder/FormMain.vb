' --------------------------------------
' --- IDRISClassBuilder - 09/28/2017 ---
' --------------------------------------

' ----------------------------------------------------------------------------------------------------
' 09/28/2017 - SBakker
'            - Switched to Arena.Common.Bootstrap.
' 04/26/2016 - SBakker
'            - Split out the string multiline checking so it is not done for [TEXT] fields.
' 02/02/2016 - SBakker
'            - Added soft errors for Path Not Found.
' 10/28/2015 - SBakker - URD 12527
'            - Added code for setting "cmd.CommandTimeout = DataConnection.SQLTimeoutSeconds"
'              everywhere that a SQLCommand object is created. This might help with the infrequent
'              timeout errors.
' 06/08/2015 - SBakker
'            - Added StringToASCII() to all values for all string fields. Someone has been slipping
'              bad characters into IDRIS.
'            - Check for either vbCr or vbLf to determine Multiline.
' 04/15/2015 - SBakker
'            - Modified FixFieldName() to change "CASE" to "CASENUM", for LASTNUM table.
' 03/31/2015 - SBakker
'            - Renamed "My.Settings.Applications" to "My.Settings.ApplicationList". Apparently you
'              can't change a User setting to an Application setting and have it work properly without
'              renaming it.
' 03/11/2015 - SBakker
'            - Changed "My.Settings.Applications" setting to be an application level setting.
' 01/23/2015 - SBakker
' 08/28/2014 - SBakker
'            - Fixed problem with CadolScale being Nothing when there are no decimal digits. It was
'              trying to build History tables and not getting the proper CADOLNUMzz() values.
' 06/18/2014 - SBakker
'            - Ignore missing IDRIS_Cadol_Sql_Xref records when filling KeyPattern.
' 06/16/2014 - SBakker
'            - Use "##ORIG##" for original field name before fixing. This allows filling and saving
'              using the SQL field name, instead of the fixed property name.
' 03/10/2014 - SBakker
'            - Added handling for a newly discovered "intdate" IDRIS data type. It is an integer that
'              holds a zero or a YYYYMMDD date. For data classes, it will convert this into a normal
'              Nullable(Of DateTime), and convert back to SQL Integer when saving. Currently this is
'              only found in ATPTRANS and ATPFILES, but there may be others. No idea why these weren't
'              set to be real DateTime fields back in 2005! But converting them in the data classes is
'              easier than fixing IDRIS tables and triggers.
'            - Fixed BlankDateProp to not include Static FuncName if the Validate routine won't use
'              it. This will only save microscopic amounts of time, but will do it everywhere.
' 02/24/2014 - SBakker
'            - Added Bootstrap loading all programs to another location, and then running from there.
' 01/30/2014 - SBakker
'            - Added a message reminding that the fields must be defined in [%DataFormat] to correctly
'              build triggers.
' 01/17/2014 - SBakker
'            - Updated list of tables which can and can't have history. No triggers for CWVOID* and
'              PYMTLINK, or all POLMAST*, POLMEAP* except for "I" and "L".
'            - Added ChangedCount to Build History.
'            - Made sure %SCF and %SORT get triggers!
' 01/10/2014 - SBakker
'            - Added routines "Public Sub Save(ByVal dc As DataConnection)" so an object can do a
'              "MyObj.Save(dc)". Same for "Delete()". Arena already had this, but IDRIS didn't. Great
'              for transactions!
'            - Must delete record in Save() routine if applicable before calling FixRecordPreSave(),
'              or the number of BeginTransactions/EndTransactions could become mismatched.
' 12/13/2013 - SBakker
'            - Added a ChangedCount to show how many classes were updated.
' 10/04/2013 - SBakker
'            - Strip the Drive from all FromPath/ToPath strings upon loading. Makes it easier so you
'              don't have to do it yourself.
'            - Always use Database and Connection name "IDRIS" for "IDRIS History".
'            - Save the LastFromPath and LastToPath settings as lists separated by ";", just like the
'              other Class Builder programs.
'            - Only enable ButtonBuildHist if "IDRIS History" is selected.
' 10/01/2013 - SBakker
'            - Added Menu Bar, Drive Combo Box, and StatusStrip.
' 09/18/2013 - SBakker
'            - Need to have the "REC_TYPE" field available for the "CODES" data class. It contains
'              alpha or numeric, plus length information.
' 09/17/2013 - SBakker
'            - Added "Partial Private Shared Sub ValidateRecordLevel()". This allows Record level
'              validation to be added in a partial class, as needed.
'            - Added additional error information.
' 08/19/2013 - SBakker
'            - Added "smalldatetime" support.
' 02/13/2013 - SBakker
'            - Added ability to handle [IDR] tables, which sometimes use a [SMALLINT] for
'              the IDENTITY field.
'            - Preserve the case on IDENTITY fields on Non-IDRIS tables.
'            - Change spaces in the output filename, field names, and class names to be
'              underlines.
'            - Prepend the database name if used for other databases, like IDR.
'            - Added FixFieldName() to CreateCloneList. Should always fix the name before
'              checking IgnoreField().
'            - Save the SQLTableName unmodified for use in the classes where appropriate.
' 02/11/2013 - SBakker
'            - Added new Non-IDRIS Data Classes, which still have an IDENTITY field but no
'              KEY, REC, or VOLUME fields.
'            - Added support for [TEXT] fields.
'            - Ignore IDRIS tables "Policy", "PolicyClass", "PolicyBlock" as they conflict
'              with other Arena classes and only contain calculated information.
' 03/20/2012 - SBakker
'            - Include "%SCF" (_SCF) in files that get built into classes. Have need to read
'              mask data. Blah!
' 01/19/2012 - SBakker - Bug 2-69
'            - Removed ZZZ_FixInfo triggers. They used "DISABLE/ENABLE TRIGGER" commands,
'              which are only available to Administrators. Instead, merged the contents into
'              the ZZZ_Ins and ZZZ_Upd triggers.
' 12/28/2011 - SBakker
'            - Fixed double-history problem with new records.
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
' 10/18/2010 - SBakker
'            - Added partial sub FixDefault_###(), so default values can be adjusted if they
'              should be different than the default for the data type. Example: Set default
'              Taxable Percent to 100, not 0.
' 09/08/2010 - SBakker
'            - Added a new Clone routine. This allows a separate object of this or a derived
'              class to become an exact duplicate of this object, for all fields in common.
' 08/27/2010 - SBakker
'            - Added the ability to create Generic Data Classes, for IDRIS tables which only
'              have a PACKED_DATA field and no split-out SQL fields. This is going to be a
'              lot like CADOL...
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
' 05/10/2010 - SBakker
'            - Added in extra "Nothing" checks for any string properties that can
'              be "Null". Otherwise it wasn't swapping the property between "" and
'              Nothing.
' 05/05/2010 - SBakker
'            - Added a button so a single table script can be converted to a
'              class. Useful for excluded files, such as History, when only one
'              is needed.
'            - Added FixFieldName. Some names are reserved words in VB.NET.
' 02/02/2010 - SBakker
'            - Added list of tables with old history info (CWADDR_H, etc). These
'              will fill the old history into the new history table first, then
'              top it off with the current data.
' 02/01/2010 - SBakker
'            - Added many more names to the list of ignored tables. None need
'              history at this point.
' 01/13/2010 - SBakker
'            - Added TRACKING to the list of ignored tables.
'            - Fixed triggers so they match IDRIS_IDE triggers.
' 01/04/2010 - SBakker
'            - Added more files which should be ignored when building History
'              scripts.
' 12/29/2009 - SBakker
'            - Ignore old Advantage Checkwriting History tables (CWXXX_H) when
'              building new history tables.
'            - Fixed BlankHistoryScript.txt to set ChangedBy field properly.
'            - Fixed to ignore certain files which shouldn't get history.
' 09/25/2009 - SBakker
'            - Changed SQLTableName, BaseQuery, FirstConj, and DeleteQuery to be
'              properties with private variables. Also they have partial subs so
'              they can be modified in a *.Part1.vb file, without having to chg
'              the original file.
' 08/12/2009 - SBakker
'            - Added PACKED_DATA updating logic to building history triggers.
'              This will mimic but replace the TRIG_FILENAME for each IDRIS
'              table.
' 06/22/2009 - SBakker
'            - Added ability to create IDRIS History tables and triggers.
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
'            - Fixed the read-only check to only check existing files.
' 12/17/2008 - SBAKKER - Arena
'            - Removed "Partial" modifier. This will be the one non-partial
'              class, but other partial class definitions can be added as well.
'            - Removed dates in heading comment. This allows easier comparisons.
' 12/05/2008 - SBAKKER - Arena
'            - Allow all IDRIS string properties to be "" by default.
'            - Force IDRIS string properties to be "" instead of Nothing.
' 12/04/2008 - SBAKKER - Arena
'            - Exclude any tables without a [KEY] field, and those without any
'              properties. If they don't have properties, all the data is in the
'              PACKED_DATA field, and thus unavailable to VB.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Arena.Common.Bootstrap

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private FromPaths() As String
    Private ToPaths() As String

#Region " Form Routines "

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try
            If Bootstrapper.MustBootstrap Then
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

        TextDatabase.Text = My.Settings.LastDatabase
        TextConnName.Text = My.Settings.LastConnection
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
            If Not Directory.Exists(CStr(ToolStripComboBoxDrive.Items(TempDriveIndex)) + "\Arena_Scripts") AndAlso
                Not Directory.Exists(CStr(ToolStripComboBoxDrive.Items(TempDriveIndex)) + "\Projects\Arena_Scripts") Then
                ToolStripComboBoxDrive.Items.RemoveAt(TempDriveIndex)
            End If
        Next

        ToolStripComboBoxDrive.SelectedItem = My.Settings.LastDrive

    End Sub

    Private Sub DoClearScreen()
        TextInput.Text = ""
        TextDatabaseName.Text = ""
        TextClassName.Text = ""
        TextFields.Text = ""
        TextOutput.Text = ""
        TextInput.Focus()
    End Sub

#End Region

#Region " TextBox Routines "

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

#End Region

#Region " Button Routines "

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
        My.Settings.LastDatabase = TextDatabase.Text
        My.Settings.LastConnection = TextConnName.Text
        My.Settings.LastDrive = ToolStripComboBoxDrive.Text
        My.Settings.Save()
        FileCount = 0
        ChangedCount = 0
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
        Dim FromDir As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(ToolStripComboBoxDrive.Text + TextFromPath.Text)
        Dim FromFiles() As FileInfo = FromDir.GetFiles("*.sql")
        For Each CurrFile As FileInfo In FromFiles
            If CurrFile.Name.IndexOf("%") >= 0 Then Continue For
            If CurrFile.Name.IndexOf("#") >= 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf(".TABLE.") < 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf("TEMP_") >= 0 Then Continue For
            FileCount += 1
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            My.Application.DoEvents()
            DoClearScreen()
            TextDatabaseName.Text = TextDatabase.Text
            Dim sr As StreamReader = CurrFile.OpenText
            TextInput.Text = sr.ReadToEnd
            sr.Close()
            If DoBuildClass() Then
                ' --- Prepend the database name if used for other databases, like IDR ---
                If TextDatabase.Text <> "IDRIS" Then
                    FullFilePath = ToolStripComboBoxDrive.Text + TextToPath.Text + "\" + TextDatabase.Text + "_" + CurrFile.Name.Replace("dbo.", "").Replace(".Table.sql", ".vb").Replace(" ", "_")
                Else
                    FullFilePath = ToolStripComboBoxDrive.Text + TextToPath.Text + "\" + CurrFile.Name.Replace("dbo.", "").Replace(".Table.sql", ".vb").Replace(" ", "_")
                End If
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

    Private Sub ButtonBuildHist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBuildHist.Click
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
        My.Settings.LastDatabase = TextDatabase.Text
        My.Settings.LastConnection = TextConnName.Text
        My.Settings.LastDrive = ToolStripComboBoxDrive.Text
        My.Settings.Save()
        FileCount = 0
        ChangedCount = 0
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
        Dim FromDir As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(ToolStripComboBoxDrive.Text + TextFromPath.Text)
        Dim FromFiles() As FileInfo = FromDir.GetFiles("*.sql")
        For Each CurrFile As FileInfo In FromFiles
            If CurrFile.Name.IndexOf("%") >= 0 Then
                If Not CurrFile.Name.StartsWith("dbo.%SCF.") AndAlso Not CurrFile.Name.StartsWith("dbo.%SORT.") Then
                    Continue For
                End If
            End If
            If CurrFile.Name.IndexOf("#") >= 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf(".TABLE.") < 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf("_HIST.TABLE.") >= 0 Then Continue For
            If CurrFile.Name.ToUpper.IndexOf("TEMP_") >= 0 Then Continue For
            If IgnoreFile(CurrFile.Name) Then Continue For
            FileCount += 1
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString + ChangedCountMsg + ChangedCount.ToString
            My.Application.DoEvents()
            DoClearScreen()
            TextDatabaseName.Text = TextDatabase.Text
            Dim sr As StreamReader = CurrFile.OpenText
            TextInput.Text = sr.ReadToEnd
            sr.Close()
            If DoBuildHistory() Then
                FullFilePath = ""
                If FileCanHaveHistory(CurrFile.Name) Then
                    FullFilePath = ToolStripComboBoxDrive.Text + TextToPath.Text + "\" + CurrFile.Name.Replace(".Table.sql", "_Hist.Table.sql")
                ElseIf FileCanHaveTrigger(CurrFile.Name) Then
                    FullFilePath = ToolStripComboBoxDrive.Text + TextToPath.Text + "\TRIG_" + CurrFile.Name.Replace(".Table.sql", ".Trigger.sql").Replace("dbo.", "")
                End If
                ' --- check if the file exists ---
                If Not String.IsNullOrWhiteSpace(FullFilePath) Then
                    If File.Exists(FullFilePath) Then
                        ' --- check if the file has changed ---
                        Dim OldSR As New StreamReader(FullFilePath)
                        Dim OldFile As String = OldSR.ReadToEnd
                        OldSR.Close()
                        If OldFile = TextOutput.Text Then
                            Continue For
                        End If
                        ' --- check if the file is read-only ---
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
            End If
        Next
        ToolStripStatusLabelMain.Text += " - Done"
    End Sub

#End Region

#Region " BlankProperties "

    Private Const BlankIntProp As String =
            "#Region "" Property ### (Int NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As Integer = 0" + vbCrLf +
            "    Private _### As Integer = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>Int NotNull</summary>" + vbCrLf +
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
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankIntPropNotNull As String =
            ""
    Private Const BlankIntPropEnd As String =
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankStringPropHeadNotNull As String =
            "#Region "" Property ### (String ??? NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As String = """"" + vbCrLf

    Private Const BlankStringPropHeadNull As String =
            "#Region "" Property ### (String ??? Null) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As String = Nothing" + vbCrLf

    Private Const BlankStringProp As String =
            "    Private _### As String = _###_Default" + vbCrLf +
            "    Private _###_Max As Integer = ???" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>String ??? NotNull</summary>" + vbCrLf +
            "    Public Property ###() As String" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As String)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            value = StringToASCII(value)" + vbCrLf +
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
            "        ' --- Allow only valid values ---" + vbCrLf
    Private Const BlankStringPropNotMultiline As String =
            "        If Value IsNot Nothing Then" + vbCrLf +
            "            If Value.Length > _###_Max Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Invalid length: "" + Value.Length.ToString)" + vbCrLf +
            "            End If" + vbCrLf +
            "            If Value.IndexOf(vbCr) >= 0 OrElse Value.IndexOf(vbLf) >= 0 Then" + vbCrLf +
            "                Throw New SystemException(FuncName + vbCrLf + ""Not a multiline property"")" + vbCrLf +
            "            End If" + vbCrLf +
            "        End If" + vbCrLf
    Private Const BlankStringPropPart2 As String =
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankStringPropNotNull As String =
            "        If Obj.### Is Nothing Then" + vbCrLf +
            "            Throw New ArgumentNullException(FuncName)" + vbCrLf +
            "        End If" + vbCrLf
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
            "    ''' <summary>Boolean NotNull</summary>" + vbCrLf +
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

    Private Const BlankDecimalProp As String =
            "#Region "" Property ### (Decimal NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As Decimal = Nothing" + vbCrLf +
            "    Private _### As Decimal = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>Decimal NotNull</summary>" + vbCrLf +
            "    Public Property ###() As Decimal" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Decimal)" + vbCrLf +
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
            "    Private Sub CheckValue_###(ByVal Value As Decimal)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDecimalPropNotNull As String =
            ""
    Private Const BlankDecimalPropEnd As String =
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDoubleProp As String =
            "#Region "" Property ### (Double NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As Double = Nothing" + vbCrLf +
            "    Private _### As Double = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>Double NotNull</summary>" + vbCrLf +
            "    Public Property ###() As Double" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Double)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If _###.Value <> value.Value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Double)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDoublePropNotNull As String =
            ""
    Private Const BlankDoublePropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Double)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Double)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankSingleProp As String =
            "#Region "" Property ### (Single NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As Single = Nothing" + vbCrLf +
            "    Private _### As Single = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>Single NotNull</summary>" + vbCrLf +
            "    Public Property ###() As Single" + vbCrLf +
            "        Get" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            Return _###" + vbCrLf +
            "        End Get" + vbCrLf +
            "        Set(ByVal value As Single)" + vbCrLf +
            "            FixDefault_###()" + vbCrLf +
            "            FixValue_###(value)" + vbCrLf +
            "            CheckValue_###(value)" + vbCrLf +
            "            If _###.Value <> value.Value Then" + vbCrLf +
            "                _### = value" + vbCrLf +
            "                If IsLocal AndAlso Not IsFillingFields Then IsChanged = True" + vbCrLf +
            "            End If" + vbCrLf +
            "        End Set" + vbCrLf +
            "    End Property" + vbCrLf +
            "" + vbCrLf +
            "    Private Sub CheckValue_###(ByVal Value As Single)" + vbCrLf +
            "        ' --- Allow only valid values ---" + vbCrLf +
            "        CheckValueMore_###(Value)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankSinglePropNotNull As String =
            ""
    Private Const BlankSinglePropEnd As String =
            "        ValidateMore_###(Obj)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "    ' --- Partial routines that can be completed in a partial class ---" + vbCrLf +
            "    Partial Private Sub FixDefault_###()" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub FixValue_###(ByRef Value As Single)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Sub CheckValueMore_###(ByVal Value As Single)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

    Private Const BlankDateProp As String =
            "#Region "" Property ### (DateTime NotNull) """ + vbCrLf +
            "" + vbCrLf +
            "    Private _###_Default As Nullable(Of DateTime) = Nothing" + vbCrLf +
            "    Private _### As Nullable(Of DateTime) = _###_Default" + vbCrLf +
            "" + vbCrLf +
            "    ''' <summary>DateTime NotNull</summary>" + vbCrLf +
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
            "    Protected Shared Sub Validate_###(ByVal Obj As %%%)" + vbCrLf
    Private Const BlankDatePropNull As String =
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf
    Private Const BlankDatePropNotNull As String =
            "        Static FuncName As String = ObjName + ""."" + System.Reflection.MethodBase.GetCurrentMethod().Name" + vbCrLf +
            "        ' --- Re-check value before saving ---" + vbCrLf +
            "        Obj.CheckValue_###(Obj.###)" + vbCrLf +
            "        ' --- Allow only valid business logic and database values ---" + vbCrLf +
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
            "    Partial Private Shared Sub ValidateMore_###(ByVal Obj As %%%)" + vbCrLf +
            "    End Sub" + vbCrLf +
            "" + vbCrLf +
            "#End Region" + vbCrLf

#End Region

#Region " Blank Templates "

    Private Const BlankFillFieldsNull As String =
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

    Private Const BlankFillFieldsNotNull As String =
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

    Private Const BlankFillFieldsIntDate As String =
            "        ' --- ### ---" + vbCrLf +
            "        Static _FieldNum_### As Integer = -99 ' not set yet" + vbCrLf +
            "        If _FieldNum_### = -99 Then" + vbCrLf +
            "            Try" + vbCrLf +
            "                _FieldNum_### = dr.GetOrdinal(""##ORIG##"")" + vbCrLf +
            "            Catch ex As Exception" + vbCrLf +
            "                _FieldNum_### = -1 ' not found" + vbCrLf +
            "            End Try" + vbCrLf +
            "        End If" + vbCrLf +
            "        If _FieldNum_### < 0 OrElse dr.IsDBNull(_FieldNum_###) OrElse dr.@@@(_FieldNum_###) = 0 Then" + vbCrLf +
            "            Obj.### = Nothing" + vbCrLf +
            "        Else" + vbCrLf +
            "            Obj.### = CDate(Arena_Utilities.DateUtils.FormatDate(dr.@@@(_FieldNum_###)))" + vbCrLf +
            "        End If" + vbCrLf

    Private Const BlankFillFieldsPackedData As String =
            "        ' --- PACKED_DATA---" + vbCrLf +
            "        Static _FieldNum_PACKED_DATA As Integer = -99 ' not set yet" + vbCrLf +
            "        If _FieldNum_PACKED_DATA = -99 Then" + vbCrLf +
            "            Try" + vbCrLf +
            "                _FieldNum_PACKED_DATA = dr.GetOrdinal(""PACKED_DATA"")" + vbCrLf +
            "            Catch ex As Exception" + vbCrLf +
            "                _FieldNum_PACKED_DATA = -1 ' not found" + vbCrLf +
            "            End Try" + vbCrLf +
            "        End If" + vbCrLf +
            "        If _FieldNum_PACKED_DATA >= 0 Then" + vbCrLf +
            "            Dim FieldLen As Integer = CInt(dr.GetBytes(_FieldNum_PACKED_DATA, 0, Nothing, 0, Obj._PACKED_DATA_Max))" + vbCrLf +
            "            Obj.PACKED_DATA = Nothing" + vbCrLf +
            "            ReDim Obj.PACKED_DATA(FieldLen - 1) ' zero-based" + vbCrLf +
            "            dr.GetBytes(_FieldNum_PACKED_DATA, 0, Obj.PACKED_DATA, 0, FieldLen)" + vbCrLf +
            "        End If" + vbCrLf

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

    Private Const BlankBooleanValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankBooleanValueNotNull As String =
            "            .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankDateValueNull As String =
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append("",'"") : .Append(Me.###.Value.ToString) : .Append(""'"")" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append("",NULL"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankIntDateValue As String =
            "            .Append("","")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(Me.###.Value.ToString(""yyyyMMdd""))" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""0"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateValueNotNull As String =
            "            .Append("",'"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankNumericFillerValue As String =
            "            .Append("",0"") ' ###" + vbCrLf

    Private Const BlankStringFillerValue As String =
            "            .Append("",''"") ' ###" + vbCrLf

    Private Const BlankPackedDataValue As String =
            "            .Append("","") : .Append(""0x"")" + vbCrLf +
            "            For TempIndex As Integer = 0 To Me.PACKED_DATA.Length - 1" + vbCrLf +
            "                .Append(Right(""00"" + Conversion.Hex(Me.PACKED_DATA(TempIndex)), 2))" + vbCrLf +
            "            Next" + vbCrLf

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

    Private Const BlankIntDateUpdate As String =
            "            .Append("",[##ORIG##] = "")" + vbCrLf +
            "            If Me.###.HasValue Then" + vbCrLf +
            "                .Append(Me.###.Value.ToString(""yyyyMMdd""))" + vbCrLf +
            "            Else" + vbCrLf +
            "                .Append(""0"")" + vbCrLf +
            "            End If" + vbCrLf

    Private Const BlankDateUpdateNotNull As String =
            "            .Append("",[##ORIG##] = '"") : .Append(Me.###.ToString) : .Append(""'"")" + vbCrLf

    Private Const BlankNumericFillerUpdate As String =
            "            .Append("",[##ORIG##] = 0"")" + vbCrLf

    Private Const BlankStringFillerUpdate As String =
            "            .Append("",[##ORIG##] = ''"")" + vbCrLf

    Private Const BlankPackedDataUpdate As String =
            "            .Append("",[PACKED_DATA] = "") : .Append(""0x"")" + vbCrLf +
            "            For TempIndex As Integer = 0 To Me.PACKED_DATA.Length - 1" + vbCrLf +
            "                .Append(Right(""00"" + Conversion.Hex(Me.PACKED_DATA(TempIndex)), 2))" + vbCrLf +
            "            Next" + vbCrLf

    ' --- SQL constants ---

    Private Const BlankBaseQuery As String = """SELECT * FROM "" + _SQLTableName"
    Private Const BlankFirstConj As String = """ WHERE"""
    Private Const BlankDeleteQuery As String = """DELETE FROM "" + _SQLTableName"

#End Region

#Region " Create Routines "

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
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            If NotNull AndAlso (CurrLine.ToUpper.IndexOf("IDENTITY") >= 0) Then
                NotNull = False
            End If
            ' --- Check for YYYYMMDD integer date datatype ---
            If CurrType.ToLower = "int" Then
                If TextClassName.Text.ToUpper = "ATPTRANS" OrElse TextClassName.Text.ToUpper = "ATPFILES" Then
                    If CurrField = "EFF_DATE" OrElse CurrField = "DATE_PAID" Then
                        CurrType = "intdate"
                    End If
                End If
            End If
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    OutLine = BlankIntProp
                    If NotNull Then
                        OutLine += BlankIntPropNotNull
                        OutLine += BlankIntPropEnd
                    Else
                        OutLine += BlankIntPropEnd
                        OutLine = OutLine.Replace("As Integer", "As Nullable(Of Integer)")
                    End If
                Case "text"
                    CurrLen = "MAX"
                    If NotNull Then
                        OutLine = BlankStringPropHeadNotNull.Replace("???", CurrLen)
                    Else
                        OutLine = BlankStringPropHeadNull.Replace("???", CurrLen)
                    End If
                    OutLine += BlankStringProp.Replace("???", CurrLen)
                    OutLine += BlankStringPropPart2.Replace("???", CurrLen)
                    If NotNull Then OutLine += BlankStringPropNotNull
                    OutLine += BlankStringPropEnd
                Case "char", "varchar", "nchar", "nvarchar"
                    CurrLen = CurrLine.Substring(CurrLine.IndexOf("(") + 1)
                    CurrLen = CurrLen.Substring(0, CurrLen.IndexOf(")"))
                    If NotNull Then
                        OutLine = BlankStringPropHeadNotNull.Replace("???", CurrLen)
                    Else
                        OutLine = BlankStringPropHeadNull.Replace("???", CurrLen)
                    End If
                    OutLine += BlankStringProp.Replace("???", CurrLen)
                    OutLine += BlankStringPropNotMultiline.Replace("???", CurrLen)
                    OutLine += BlankStringPropPart2.Replace("???", CurrLen)
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
                    If NotNull Then
                        OutLine += BlankDecimalPropNotNull
                        OutLine += BlankDecimalPropEnd
                    Else
                        OutLine += BlankDecimalPropEnd
                        OutLine = OutLine.Replace("As Decimal", "As Nullable(Of Decimal)")
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
                    If NotNull Then
                        OutLine += BlankDatePropNotNull
                    Else
                        OutLine += BlankDatePropNull
                    End If
                    OutLine += BlankDatePropEnd
                Case "intdate"
                    NotNull = False
                    OutLine = BlankDateProp
                    OutLine += BlankDatePropNull
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
            Result.Append(OutLine.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        Return Result.ToString
    End Function

    Private Function CreatePropertiesPackedData() As String
        Dim Result As New StringBuilder
        ' -----------------------------
        With Result

        End With
        Return Result.ToString
    End Function

    Private Function CreateHistoryFields() As String
        Dim Lines() As String
        Dim CurrLine As String
        Dim CurrField As String
        Dim Result As New StringBuilder
        ' -----------------------------
        Lines = TextFields.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
        For Each CurrLine In Lines
            CurrField = CurrLine.Trim.Replace("  ", " ")
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            If CurrField.IndexOf(" NULL") < 0 Then Continue For
            If CurrField.IndexOf("IDENTITY") >= 0 Then
                CurrField = CurrField.Substring(0, CurrField.IndexOf("IDENTITY")).TrimEnd + " NULL"
            End If
            If CurrField.IndexOf("[timestamp]") >= 0 Then
                CurrField = CurrField.Replace("[timestamp]", "[binary](8)")
            End If
            CurrField = Replace(CurrField, "NOT NULL", "NULL")
            CurrField = CurrField.Substring(0, CurrField.IndexOf("NULL") + 4)
            If Result.Length > 0 Then
                Result.Append(",")
                Result.Append(vbCrLf)
            End If
            Result.Append(vbTab)
            Result.Append(CurrField)
        Next
        If Result.Length > 0 Then
            Result.Append(vbCrLf)
        End If
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
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            CurrCvt = "???"
            ' --- Check for YYYYMMDD integer date datatype ---
            If CurrType.ToLower = "int" Then
                If TextClassName.Text.ToUpper = "ATPTRANS" OrElse TextClassName.Text.ToUpper = "ATPFILES" Then
                    If CurrField = "EFF_DATE" OrElse CurrField = "DATE_PAID" Then
                        CurrType = "intdate"
                    End If
                End If
            End If
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
                Case "char", "varchar", "nchar", "nvarchar", "text"
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
                Case "intdate"
                    CurrCvt = "GetInt32"
            End Select
            If CurrCvt = "???" Then
                Result.Append("### Unknown Type ###" + vbCrLf)
            ElseIf CurrType.ToLower = "intdate" Then
                Result.Append(BlankFillFieldsIntDate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("@@@", CurrCvt))
            ElseIf NotNull Then
                Result.Append(BlankFillFieldsNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("@@@", CurrCvt))
            Else
                Result.Append(BlankFillFieldsNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField).Replace("@@@", CurrCvt))
            End If
        Next
        ' --- Check if generic IDRIS table ---
        If Result.Length = 0 Then
            Result.Append(BlankFillFieldsPackedData)
        End If
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
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then Continue For
            Result.Append(BlankValidate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        ' --- Check if generic IDRIS table ---
        If Result.Length = 0 Then
            Result.Append(BlankValidate.Replace("%%%", TextClassName.Text).Replace("###", "PACKED_DATA"))
        End If
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
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            If CurrField.ToUpper.IndexOf("IDENTITY") >= 0 Then Continue For ' ignore identity fields
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            Result.Append(BlankFieldList.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
        Next
        ' --- Check if generic IDRIS table ---
        If Result.Length = 0 Then
            Result.Append(BlankFieldList.Replace("%%%", TextClassName.Text).Replace("###", "PACKED_DATA"))
        End If
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
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            If CurrField.ToUpper.IndexOf("IDENTITY") >= 0 Then Continue For ' ignore identity fields
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            ' --- Check for YYYYMMDD integer date datatype ---
            If CurrType.ToLower = "int" Then
                If TextClassName.Text.ToUpper = "ATPTRANS" OrElse TextClassName.Text.ToUpper = "ATPFILES" Then
                    If CurrField = "EFF_DATE" OrElse CurrField = "DATE_PAID" Then
                        CurrType = "intdate"
                    End If
                End If
            End If
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankNumericFillerValue.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankStringFillerValue.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankStringValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankBooleanValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "decimal", "money", "smallmoney"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankNumericFillerValue.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "date", "datetime", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateValueNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateValueNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "intdate"
                    Result.Append(BlankIntDateValue.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
        Next
        ' --- Check if generic IDRIS table ---
        If Result.Length = 0 Then
            Result.Append(BlankPackedDataValue)
        End If
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
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) < 0 Then Continue For
            If CurrField.IndexOf("]", CurrField.IndexOf("]") + 1) < 0 Then Continue For
            If CurrField.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            If CurrField.ToUpper.IndexOf("IDENTITY") >= 0 Then Continue For ' ignore identity fields
            CurrType = CurrField.Substring(CurrField.IndexOf("[", CurrField.IndexOf("[") + 1) + 1)
            CurrType = CurrType.Substring(0, CurrType.IndexOf("]"))
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            NotNull = (CurrLine.ToUpper.IndexOf("NOT NULL") >= 0)
            ' --- Check for YYYYMMDD integer date datatype ---
            If CurrType.ToLower = "int" Then
                If TextClassName.Text.ToUpper = "ATPTRANS" OrElse TextClassName.Text.ToUpper = "ATPFILES" Then
                    If CurrField = "EFF_DATE" OrElse CurrField = "DATE_PAID" Then
                        CurrType = "intdate"
                    End If
                End If
            End If
            Select Case CurrType.ToLower
                Case "int", "smallint", "tinyint"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankNumericFillerUpdate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "char", "varchar", "nchar", "nvarchar", "text"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankStringFillerUpdate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankStringUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankStringUpdateNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "bit"
                    If NotNull Then
                        Result.Append(BlankBooleanUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankBooleanUpdateNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "decimal", "money", "smallmoney"
                    If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then
                        Result.Append(BlankNumericFillerUpdate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    ElseIf NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "float", "real"
                    If NotNull Then
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankNumericUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "date", "datetime", "smalldatetime"
                    If NotNull Then
                        Result.Append(BlankDateUpdateNotNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    Else
                        Result.Append(BlankDateUpdateNull.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                    End If
                Case "intdate"
                    Result.Append(BlankIntDateUpdate.Replace("%%%", TextClassName.Text).Replace("###", CurrField).Replace("##ORIG##", OrigField))
                Case Else
                    Result.Append("### Unknown Type ###" + vbCrLf)
            End Select
        Next
        ' --- Check if generic IDRIS table ---
        If Result.Length = 0 Then
            Result.Append(BlankPackedDataUpdate)
        End If
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
            If CurrField.IndexOf("[") < 0 Then Continue For
            If CurrField.IndexOf("]") < 0 Then Continue For
            If CurrField.ToUpper.IndexOf(" AS ") >= 0 Then Continue For ' calculated fields
            If CurrField.ToUpper.IndexOf("IDENTITY") >= 0 Then Continue For ' ignore identity fields
            CurrField = CurrField.Substring(0, CurrField.IndexOf("]"))
            CurrField = CurrField.Substring(CurrField.IndexOf("[") + 1)
            OrigField = CurrField
            CurrField = FixFieldName(CurrField)
            If IgnoreField(CurrField) Then Continue For
            If CurrField.ToUpper.IndexOf("FILLER") >= 0 Then Continue For ' filler fields
            ' --- Note that these are copying the internal values directly ---
            Result.Append("        Obj._")
            Result.Append(CurrField)
            Result.Append(" = Me._")
            Result.Append(CurrField)
            Result.Append(vbCrLf)
        Next
        Return Result.ToString
    End Function

#End Region

#Region " DoBuild Routines "

    Private Function DoBuildClass() As Boolean
        Dim MyApp As Assembly
        Dim SR As StreamReader
        Dim TempResult As String
        Dim sb As New StringBuilder
        Dim Lines() As String
        Dim InFields As Boolean = False
        Dim HasKey As Boolean = False
        Dim PropSection As String = ""
        Dim IsHistoryFile As Boolean = False
        Dim IdentityField As String = ""
        Dim SQLTableName As String = ""
        Dim KeyPattern As String = ""
        ' ----------------------------------
        ' --- split the SQL table definition into Database, TableName, and Fields ---
        If TextClassName.Text = "" Then
            Lines = TextInput.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
            For Each TempLine As String In Lines
                TempLine = TempLine.Trim
                If TempLine.StartsWith("USE [") Then
                    TempLine = TempLine.Substring(5, TempLine.IndexOf("]") - 5)
                    TextDatabaseName.Text = TempLine
                ElseIf TempLine.StartsWith("CREATE TABLE [") Then
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1) ' remove schema
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1)
                    TempLine = TempLine.Substring(0, TempLine.IndexOf("]"))
                    SQLTableName = TempLine ' Need to save actual table name
                    TempLine = TempLine.Replace(" ", "_")
                    If TempLine.StartsWith("%") OrElse TempLine.StartsWith("_") Then
                        If TempLine.ToUpper <> "%SCF" Then
                            Return False
                        End If
                        TempLine = TempLine.Replace("%", "_")
                    End If
                    If IgnoreFile(TempLine) Then
                        Return False
                    End If
                    If TextDatabase.Text <> "IDRIS" Then
                        TextClassName.Text = TextDatabase.Text + "_" + TempLine
                    Else
                        TextClassName.Text = TempLine
                    End If
                    If TempLine.ToUpper.EndsWith("_HIST") Then
                        IsHistoryFile = True
                    End If
                    InFields = True
                ElseIf TempLine.StartsWith("CONSTRAINT") OrElse TempLine.StartsWith(")") Then
                    InFields = False
                ElseIf InFields AndAlso TempLine.StartsWith("[") Then
                    If TempLine.ToUpper.StartsWith("[KEY]") Then
                        HasKey = True
                    ElseIf TempLine.ToUpper.IndexOf("IDENTITY") >= 0 AndAlso TempLine.ToUpper.IndexOf("INT]") >= 0 Then
                        IdentityField = TempLine.Substring(1, TempLine.IndexOf("]") - 1)
                    End If
                    sb.Append(TempLine)
                    sb.Append(vbCrLf)
                End If
            Next
            ' --- Check if it had a KEY field ---
            If Not HasKey AndAlso String.IsNullOrWhiteSpace(IdentityField) Then
                Return False
            End If
            TextFields.Text = sb.ToString
        End If
        ' --- Build the class from the known information ---
        MyApp = Assembly.GetExecutingAssembly
        KeyPattern = ""
        If Not HasKey Then
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankNonIDRISDataClass.txt"))
        ElseIf IsHistoryFile Then
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankIDRISHistClass.txt"))
        Else
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankIDRISDataClass.txt"))
            Dim TempCadolSQLXref As IDRIS_Cadol_Sql_Xref
            TempCadolSQLXref = IDRIS_Cadol_Sql_Xref.GetByTableName(SQLTableName)
            If TempCadolSQLXref IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(TempCadolSQLXref.CadolKey) Then
                    KeyPattern = TempCadolSQLXref.CadolKey
                End If
            End If
        End If
        TempResult = SR.ReadToEnd
        SR.Close()
        ' --- Check if it's a generic IDRIS table with no SQL fields ---
        PropSection = CreateProperties()
        If PropSection = "" Then
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankGenericDataClass.txt"))
            TempResult = SR.ReadToEnd
            SR.Close()
            PropSection = CreatePropertiesPackedData()
        End If
        If TextDatabaseName.Text <> "" Then
            TempResult = TempResult.Replace("$Database$", TextDatabaseName.Text)
        End If
        TempResult = TempResult.Replace("$TableName$", SQLTableName)
        TempResult = TempResult.Replace("$ClassName$", TextClassName.Text)
        TempResult = TempResult.Replace("$Class_Spc$", StrDup(TextClassName.Text.Length, " "c))
        TempResult = TempResult.Replace("$---$", StrDup(22 - 13 + TextDatabaseName.Text.Length + TextClassName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$---1---$", StrDup(21 + TextConnName.Text.Length, "-"c))
        TempResult = TempResult.Replace("$MMDDYYYY$", TextDate.Text)
        If TextConnName.Text <> "" Then
            TempResult = TempResult.Replace("$ConnName$", TextConnName.Text)
        End If
        TempResult = TempResult.Replace("$Properties$" + vbCrLf, PropSection)
        TempResult = TempResult.Replace("$BaseQuery$", BlankBaseQuery)
        TempResult = TempResult.Replace("$FirstConj$", BlankFirstConj)
        TempResult = TempResult.Replace("$DeleteQuery$", BlankDeleteQuery)
        TempResult = TempResult.Replace("$KeyPattern$", KeyPattern)
        TempResult = TempResult.Replace("$FillFields$" + vbCrLf, CreateFillFields)
        TempResult = TempResult.Replace("$ValidateList$" + vbCrLf, CreateValidateList)
        TempResult = TempResult.Replace("$FieldList$" + vbCrLf, CreateFieldList)
        TempResult = TempResult.Replace("$ValueList$" + vbCrLf, CreateValueList)
        TempResult = TempResult.Replace("$UpdateList$" + vbCrLf, CreateUpdateList)
        TempResult = TempResult.Replace("$CloneList$" + vbCrLf, CreateCloneList)
        TempResult = TempResult.Replace("[_", "[%") ' have to fix _SQLTableName for "%SCF"
        If Not HasKey Then
            TempResult = TempResult.Replace("$ID$", IdentityField)
        End If
        TempResult = TempResult.Replace("As Integer = MAX", "As Integer = Integer.MaxValue")
        TextOutput.Text = TempResult
        TextOutput.Focus()
        TextOutput.SelectAll()
        Return True
    End Function

    Private Function DoBuildHistory() As Boolean
        Dim MyApp As Assembly
        Dim SR As StreamReader
        Dim TempResult As String
        Dim sb As New StringBuilder
        Dim Lines() As String
        Dim InFields As Boolean = False
        Dim HasKey As Boolean = False
        Dim FieldSection As String = ""
        Dim PackedDataList As String = ""
        Dim SQLTableName As String = ""
        ' -------------------------------
        ' --- split the SQL table definition into Database, TableName, and Fields ---
        If TextClassName.Text = "" Then
            Lines = TextInput.Text.Replace(vbCrLf, vbLf).Split(CChar(vbLf))
            For Each TempLine As String In Lines
                TempLine = TempLine.Trim
                If TempLine.StartsWith("USE [") Then
                    TempLine = TempLine.Substring(5, TempLine.IndexOf("]") - 5)
                    TextDatabaseName.Text = TempLine
                ElseIf TempLine.StartsWith("CREATE TABLE [") Then
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1) ' remove schema
                    TempLine = TempLine.Substring(TempLine.IndexOf("[") + 1)
                    TempLine = TempLine.Substring(0, TempLine.IndexOf("]"))
                    SQLTableName = TempLine ' Need to save actual table name
                    TempLine = TempLine.Replace(" ", "_")
                    If TempLine.StartsWith("%") Then
                        If TempLine <> "%SCF" AndAlso TempLine <> "%SORT" Then
                            Return False
                        End If
                    End If
                    If TempLine.StartsWith("_") Then
                        If TempLine <> "_SCF" AndAlso TempLine <> "_SORT" Then
                            Return False
                        End If
                    End If
                    If TextDatabase.Text <> "IDRIS" Then
                        TextClassName.Text = TextDatabase.Text + "_" + TempLine
                    Else
                        TextClassName.Text = TempLine
                    End If
                    InFields = True
                ElseIf TempLine.StartsWith("CONSTRAINT") OrElse TempLine.StartsWith(")") Then
                    InFields = False
                ElseIf InFields AndAlso TempLine.StartsWith("[") Then
                    If TempLine.ToUpper.StartsWith("[KEY]") Then
                        HasKey = True
                    End If
                    sb.Append(TempLine)
                    sb.Append(vbCrLf)
                End If
            Next
            ' --- Check if it had a KEY field ---
            If Not HasKey Then
                Return False
            End If
            TextFields.Text = sb.ToString
        End If

        ' --- Build the class from the known information ---
        MyApp = Assembly.GetExecutingAssembly
        If FileHasOldHistory(TextClassName.Text) Then
            ' --- Build History file from old history, and build history triggers ---
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankHistoryFromOldScript.txt"))
        ElseIf FileCanHaveHistory(TextClassName.Text) Then
            ' --- Build History file and history triggers ---
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankHistoryScript.txt"))
        Else
            ' --- Only build file triggers ---
            SR = New StreamReader(MyApp.GetManifestResourceStream(My.Application.Info.AssemblyName + ".BlankFileTrigger.txt"))
        End If

        TempResult = SR.ReadToEnd
        SR.Close()
        TempResult = TempResult.Replace("$TableName$", SQLTableName)
        TempResult = TempResult.Replace("$ClassName$", TextClassName.Text)

        FieldSection = CreateHistoryFields()
        If FieldSection = "" Then
            Return False
        End If
        TempResult = TempResult.Replace("$Fields$" + vbCrLf, FieldSection)

        PackedDataList = BuildPackedDataList(TextClassName.Text)
        If PackedDataList <> "" Then
            PackedDataList = "UPDATE [dbo].[" + TextClassName.Text + "]" + vbCrLf +
                             vbTab + "SET PACKED_DATA =" + vbCrLf +
                             PackedDataList +
                             vbTab + "WHERE [REC] IN (SELECT [REC] FROM inserted)" + vbCrLf +
                             vbTab + "AND (NOT UPDATE (PACKED_DATA) OR PACKED_DATA IS NULL);" + vbCrLf
        End If
        If PackedDataList.StartsWith("###") Then
            Dim Answer As DialogResult = MessageBox.Show(PackedDataList, My.Application.Info.AssemblyName,
                                                         MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
            If Answer = DialogResult.Abort Then
                Me.Close()
            End If
            Return False
        End If
        TempResult = TempResult.Replace("$PackedDataList$" + vbCrLf, PackedDataList)

        TextOutput.Text = TempResult
        TextOutput.Focus()
        TextOutput.SelectAll()
        Return True

    End Function

#End Region

#Region " Internal Functions "

    Private Function IgnoreField(ByVal CurrField As String) As Boolean
        If CurrField.ToUpper = "REC" Then Return True
        If CurrField.ToUpper = "KEY" Then Return True
        If CurrField.ToUpper = "DEVICE" Then Return True
        If CurrField.ToUpper = "VOLUME" Then Return True
        If CurrField.ToUpper = "ROWVERSION" Then Return True
        If CurrField.ToUpper = "LASTCHANGED" Then Return True
        If CurrField.ToUpper = "CHANGEDBY" Then Return True
        If CurrField.ToUpper = "NUM_BYTES" Then Return True
        If CurrField.ToUpper = "NUM_ALPHA" Then Return True
        If CurrField.ToUpper = "NUM_ALPHAS" Then Return True
        If CurrField.ToUpper = "REC_TYPE" Then
            If TextClassName.Text.ToUpper = "CODES" Then Return False
            Return True
        End If
        If CurrField.ToUpper = "PACKED_DATA" Then Return True
        Return False
    End Function

    Private Function BuildPackedDataList(ByVal TableName As String) As String
        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name
        Dim TempDataFormatList As List(Of IDRIS_DataFormat)
        Try
            TempDataFormatList = IDRIS_DataFormat.GetAllByTableName(TableName)
            Dim Result As New StringBuilder
            For Each TempDataFormat As IDRIS_DataFormat In TempDataFormatList
                If TempDataFormat.FieldName = "PACKED_DATA" Then
                    Continue For
                End If
                If Result.Length > 0 Then
                    Result.Append(" +")
                    Result.Append(vbCrLf)
                End If
                Result.Append(vbTab)
                With TempDataFormat
                    If .CadolScale Is Nothing Then
                        .CadolScale = 0 ' In database as NULL if no decimal digits exist. Needs value for use below.
                    End If
                    Select Case .CadolType
                        Case "A"
                            Result.Append("dbo.CADOLALPHA([" + .FieldName + "])")
                        Case "N"
                            Result.Append("dbo.CADOLNUM" + .CadolLength.ToString + .CadolScale.ToString +
                                                "([" + .FieldName + "])")
                        Case "D"
                            Result.Append("dbo.CADOLDATE" + .CadolLength.ToString +
                                                "([" + .FieldName + "])")
                        Case "DC"
                            Result.Append("dbo.CADOLDATE3C([" + .FieldName + "])")
                        Case "X"
                            Result.Append("dbo.CADOLFIXED([" + .FieldName + "]," + .CadolLength.ToString + ")")
                        Case "U"
                            Result.Append("dbo.CADOLUNTERM([" + .FieldName + "]," + .CadolLength.ToString + ")")
                        Case "B"
                            Result.Append("[" + .FieldName + "]")
                        Case "C"
                            Result.Append("dbo.CADOLNUM" + .CadolLength.ToString + .CadolScale.ToString +
                                                "(" + .CadolValue.ToString + ")")
                        Case "FN"
                            Result.Append("dbo.CADOLNUM" + .CadolLength.ToString + .CadolScale.ToString +
                                                "(0)")
                        Case "FA"
                            Result.Append("dbo.CADOLALPHA('')")
                        Case Else
                            Return "### " + FuncName + ": Invalid Field Type: """ + .CadolType + """ ###"
                    End Select
                End With
            Next
            If Result.Length > 0 Then
                Result.Append(vbCrLf)
            End If
            Return Result.ToString
        Catch ex As Exception
            Return "### " + FuncName + ": " + ex.Message + " ###"
        End Try
    End Function

    Private Function IgnoreFile(ByVal Filename As String) As Boolean
        Filename = Filename.ToUpper
        If Filename.StartsWith("DBO.") Then
            Filename = Filename.Substring(4)
        End If
        If Filename.EndsWith(".TABLE.SQL") Then
            Filename = Filename.Substring(0, Filename.Length - Len(".TABLE.SQL"))
        End If
        If Filename = "CWTAX" Then Return True
        If Filename = "POLICY" Then Return True
        If Filename = "POLICYCLASS" Then Return True
        If Filename = "POLICYBLOCK" Then Return True
        If Filename.EndsWith("_SAVE") Then Return True
        Return False
    End Function

    Private Function FileCanHaveHistory(ByVal Filename As String) As Boolean
        If Filename.ToLower.StartsWith("dbo.") Then
            Filename = Filename.Substring(4)
        End If
        If Filename.IndexOf("."c) >= 0 Then
            Filename = Filename.Substring(0, Filename.IndexOf("."c))
        End If
        Filename = Filename.ToUpper
        ' --- Exceptions which do need history files ---
        If Filename = "CLNTXREF" Then Return True
        If Filename = "CLXREF" Then Return True
        If Filename = "POLMASTI" Then Return True
        If Filename = "POLMASTL" Then Return True
        If Filename = "POLMEAPI" Then Return True
        If Filename = "POLMEAPL" Then Return True
        '' If Filename = "ATPTRANS" Then Return True
        ' --- These are files which don't need history files ---
        If Filename = "%SCF" Then Return False
        If Filename = "%SORT" Then Return False
        If Filename = "_SCF" Then Return False
        If Filename = "_SORT" Then Return False
        If Filename = "ADJPREM" Then Return False
        If Filename = "ARMAST" Then Return False
        If Filename = "ATPHIST" Then Return False
        If Filename = "ATPSEQNM" Then Return False
        If Filename = "ATPWK1D" Then Return False
        If Filename = "BATCHNO" Then Return False
        If Filename = "CANCON" Then Return False
        If Filename = "CAPOLXR" Then Return False
        If Filename = "CASEBUY" Then Return False
        If Filename = "CESSRPT" Then Return False
        If Filename = "CLADJDU" Then Return False
        If Filename = "CLHIST" Then Return False
        If Filename = "CLIENTUR" Then Return False
        If Filename = "CLMSRCHS" Then Return False
        If Filename = "CLNTMAST" Then Return False
        If Filename = "CLTADVIS" Then Return False
        If Filename = "CLWKRPT" Then Return False
        If Filename = "CODESFLD" Then Return False
        If Filename = "CWADDR_H" Then Return False
        If Filename = "CWCLAIM_H" Then Return False
        If Filename = "CWHIST" Then Return False
        If Filename = "CWINDEX" Then Return False
        If Filename = "CWOFAC" Then Return False
        If Filename = "CWPOL_H" Then Return False
        If Filename = "CWTAX" Then Return False
        If Filename = "DBSSUM" Then Return False
        If Filename = "DETPREM" Then Return False
        If Filename = "FULLPMT" Then Return False
        If Filename = "MFPOLXR" Then Return False
        If Filename = "PICALC" Then Return False
        If Filename = "PIPREM" Then Return False
        If Filename = "POLSRCH" Then Return False
        If Filename = "POLSREAP" Then Return False
        If Filename = "POLXS" Then Return False
        If Filename = "PRSUSP" Then Return False
        If Filename = "PYMTLINK" Then Return False
        If Filename = "SWAPERR" Then Return False
        If Filename = "SWAPLIST" Then Return False
        If Filename = "TRACKING" Then Return False
        If Filename = "TRACKSCH" Then Return False
        If Filename = "TRACKXRF" Then Return False
        If Filename = "UNUMCNV" Then Return False
        ' --- Check for files that come in groups ---
        If Filename.StartsWith("AR") Then Return False
        If Filename.StartsWith("ATPFILE") Then Return False
        If Filename.StartsWith("ATPTRAN") Then Return False
        If Filename.StartsWith("CLDU") Then Return False
        If Filename.StartsWith("CLNTPRM") Then Return False
        If Filename.StartsWith("CLSTD") Then Return False
        If Filename.StartsWith("CWDATA") Then Return False
        If Filename.StartsWith("CWTOPAY") Then Return False
        If Filename.StartsWith("CWVOID") Then Return False
        If Filename.StartsWith("DIRY") Then Return False
        If Filename.StartsWith("DRMK") Then Return False
        If Filename.StartsWith("EARNP") Then Return False
        If Filename.StartsWith("EX3") Then Return False
        If Filename.StartsWith("EXS") Then Return False
        If Filename.StartsWith("FIS") Then Return False
        If Filename.StartsWith("GL") Then Return False
        If Filename.StartsWith("IBNR") Then Return False
        If Filename.StartsWith("LASTNUM") Then Return False
        If Filename.StartsWith("POLMAST") Then Return False
        If Filename.StartsWith("POLMEAP") Then Return False
        If Filename.StartsWith("POOLMA") Then Return False
        If Filename.StartsWith("PRDT") Then Return False
        If Filename.StartsWith("PREM") Then Return False
        If Filename.StartsWith("PRI98") Then Return False
        If Filename.StartsWith("PRM") Then Return False
        If Filename.StartsWith("REHIST") Then Return False
        If Filename.StartsWith("RETRAN") Then Return False
        If Filename.StartsWith("STAT") Then Return False
        If Filename.StartsWith("TEMP") Then Return False
        If Filename.StartsWith("VM") Then Return False
        If Filename.StartsWith("XT") Then Return False
        ' --- Check for files which end with specific strings ---
        If Filename.EndsWith("_DMS") Then Return False
        If Filename.EndsWith("_HIST") Then Return False
        If Filename.EndsWith("_P1") Then Return False
        If Filename.EndsWith("_P2") Then Return False
        If Filename.EndsWith("_P3") Then Return False
        If Filename.EndsWith("_SAVE") Then Return False
        If Filename.EndsWith("WK") Then Return False
        If Filename.EndsWith("WK1") Then Return False
        If Filename.EndsWith("WKT") Then Return False
        If Filename.EndsWith("WORK") Then Return False
        If Filename.EndsWith("WRK1") Then Return False
        If Filename.EndsWith("XREF") Then Return False
        ' --- Find files with specific strings anywhere in their names ---
        If Filename.IndexOf("DUMMY") >= 0 Then Return False
        ' --- Done ---
        Return True
    End Function

    Private Function FileCanHaveTrigger(ByVal Filename As String) As Boolean
        ' --- Done ---
        Return True
    End Function

    Private Function FileHasOldHistory(ByVal Filename As String) As Boolean
        ' --- Some Checkwriting files have old history ---
        If Filename = "CWADDR" Then Return True
        If Filename = "CWCLAIM" Then Return True
        If Filename = "CWPOL" Then Return True
        Return False
    End Function

    Private Function FixFieldName(ByVal Value As String) As String
        Select Case Value.ToUpper
            Case "CASE"
                Value = "CASENUM"
            Case "EVENT"
                Value = "EventFlag"
        End Select
        If Value.IndexOf(" ") >= 0 Then
            Value = Value.Trim.Replace(" ", "_")
        End If
        Return Value
    End Function

#End Region

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
    End Sub

    Private Sub ToolStripComboBoxApp_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBoxApp.SelectedIndexChanged
        ToolStripStatusLabelMain.Text = ""
        With ToolStripComboBoxApp
            If .SelectedIndex < 0 Then Exit Sub
            If ToolStripComboBoxApp.SelectedIndex < 0 Then
                TextFromPath.Text = ""
                TextToPath.Text = ""
                TextDatabase.Text = ""
                TextConnName.Text = ""
                ButtonBuildAll.Enabled = False
                ButtonBuildHist.Enabled = False
            Else
                TextFromPath.Text = FromPaths(.SelectedIndex)
                TextToPath.Text = ToPaths(.SelectedIndex)
                If CStr(.Items(.SelectedIndex)).ToUpper = "IDRIS HISTORY" Then
                    TextDatabase.Text = "IDRIS"
                    TextConnName.Text = "IDRIS"
                    ButtonBuildAll.Enabled = False
                    ButtonBuildHist.Enabled = True
                    ToolStripStatusLabelMain.Text = "Fields must be defined in [%DataFormat] for proper Trigger creation"
                Else
                    TextDatabase.Text = CStr(.Items(.SelectedIndex))
                    TextConnName.Text = CStr(.Items(.SelectedIndex))
                    ButtonBuildAll.Enabled = True
                    ButtonBuildHist.Enabled = False
                End If
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

End Class
