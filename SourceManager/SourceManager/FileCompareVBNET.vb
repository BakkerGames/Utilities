' ----------------------------------------
' --- FileCompareVBNET.vb - 08/05/2016 ---
' ----------------------------------------

' ----------------------------------------------------------------------------------------------------
' 08/05/2016 - SBakker
'            - Also ignore VB6 version numbers in project files.
' 07/20/2016 - SBakker
'            - Removed all special handling of environments, drives, paths, and thumbprints.
'            - Added "VisualStudioVersion" to IgnoreVersionNumbers checking.
' 12/16/2015 - SBakker
'            - Merged all the logic for Local and PC comparisons together, checking both. This fixes
'              an issue with hard-coded paths inside project settings files.
' 10/08/2015 - SBakker
'            - Changed so the C: drive uses the environment "PC" instead of "Local". Avoids confusion
'              when the Y: drive is also "Local" and working in the office.
' 05/23/2014 - SBakker
'            - Ignore "'     Runtime Version:" lines when comparing VB.NET project files.
'            - Cleaned up handling of "$", "%24", and "#PATH#" replacements.
' 05/04/2014 - SBakker
'            - Ignore '<meta name="timestamp" content=' lines.
' 02/19/2014 - SBakker
'            - Ignore "<GenerateManifests>" when IgnoreVersionNumbers is true. For ClickOnce, this
'              needs to be "true" on WinEXE, but on Bootstrap, it needs to be "false.
' 02/03/2014 - SBakker
'            - Make "<PlatformTarget>AnyCPU</PlatformTarget>" an error in project files. Using "x86"
'              makes programs load faster!
' 12/05/2013 - SBakker - URD 12229
'            - Added handling for the new AcceptDN environment.
' 11/11/2013 - SBakker
'            - Ignore <ProductName> with "Microsoft" and "Report Viewer" when checking environments.
' 10/23/2013 - SBakker
'            - Switch to Arena versions of ConfigInfo, DataConn, and Utilities.
' 10/15/2013 - SBakker
'            - Added handling for the new Report environment.
' 09/25/2013 - SBakker
'            - Added special handling for the new Finance environment.
' 02/19/2013 - SBakker
'            - Exclude references to Microsoft and Interop info when ignoring versions.
' 01/08/2013 - SBakker
'            - Exclude Copyright info when ignoring versions.
' 08/24/2012 - SBakker
'            - Exclude lines with "C:\Program Files".
' 02/15/2012 - SBakker
'            - Fixed differences between <PublishUrl> needing "%24" and all others needing
'              "$" in any UNC path like "\\localhost\c$\".
' 02/13/2012 - SBakker
'            - Fixed FileCompareVBNET and FindReplace sections to handle local computer name
'              replacing "LOCALHOST". Now it will work with any computer!
' 08/31/2011 - SBakker
'            - Replace invalid Options with #ERROR-xyz#, so the projects look different
'              and need to be copied.
' 06/21/2011 - SBakker
'            - Removed checking for "_#ENV". Caused issues in unexpected places.
' 06/15/2011 - SBakker
'            - Added checking for "_#ENV" to go along with " - #ENV#".
' 05/23/2011 - SBakker
'            - Added IgnoreVersionNumbers property, for showing only actual changes to
'              project and assembly files.
' 04/13/2011 - SBakker
'            - Replace DebugType = pdb-only with #ERROR-xyz#, so the projects look different
'              and need to be copied.
' 04/07/2011 - SBakker
'            - Made everything that's not Local, Test, Accept, or Prod be a PC.
'            - Allow for both Unc and Drive paths in the settings. Swap as necessary.
' 04/05/2011 - SBakker
'            - Fixed issue with comparing " - #ENV#" to Production. Now it's not a one-way
'              comparison.
'            - Ignore .NET Framework product names.
' 03/29/2011 - SBakker
'            - Don't replace " - Local", etc, with " - Prod", instead replace with blank.
'              This makes conversions one-way, but is nicer for Click-Once applications.
' 02/22/2011 - SBakker
'            - Removed exclusion for XCOPY lines. Should be no issues anymore.
' 01/20/2011 - SBakker
'            - Added checks for PublicKeyTokens. Forgot all about them until VB.NET 4.0.
' ----------------------------------------------------------------------------------------------------

Imports Arena_Utilities.StringUtils

''' <summary>
''' A special version of FileCompareClass which has handling for VB.NET project files and source environments.
''' </summary>
Public Class FileCompareVBNET

    Inherits FileCompareDataClass.FileCompareClass

    Protected ErrorNum As Integer = 0

#Region " Public Properties "

    Public Property IgnoreVersionNumbers As Boolean = False

#End Region

#Region " Public Routines "

    ''' <summary>
    ''' Special exclusions for comparing VB.NET project files.
    ''' </summary>
    Public Function ExcludeLine(ByVal CurrLine As String) As Boolean
        Return False
    End Function

#End Region

#Region " Public Overrides Routines "

    Public Overrides Sub ResetFlags()
        MyBase.ResetFlags()
        IgnoreVersionNumbers = False
    End Sub

#End Region

#Region " Protected Overrides Routines "

    Protected Overrides Function IgnoreLine(ByVal CurrLine As String) As Boolean
        If IgnoreVersionNumbers Then
            ' --- Visual Studio .NET Version Numbers ---
            If CurrLine.Contains("<Assembly: AssemblyVersion(") Then Return True
            If CurrLine.Contains("<Assembly: AssemblyFileVersion(") Then Return True
            If CurrLine.Contains("<ApplicationVersion>") Then Return True
            If CurrLine.Contains("<ApplicationRevision>") Then Return True
            If CurrLine.Contains("<MinimumRequiredVersion>") Then Return True
            If CurrLine.Contains("<Assembly: AssemblyCopyright") Then Return True
            If CurrLine.Contains("<Reference Include=""Microsoft") Then Return True
            If CurrLine.Contains(".Interop.") Then Return True
            If CurrLine.Contains("<GenerateManifests>") Then Return True
            If CurrLine.Contains("VisualStudioVersion =") Then Return True
            ' --- VB6 Version Numbers ---
            If CurrLine.Contains("MajorVer=") Then Return True
            If CurrLine.Contains("MinorVer=") Then Return True
            If CurrLine.Contains("RevisionVer=") Then Return True
        End If
        If CurrLine.StartsWith("<meta name=""timestamp"" content=") Then
            Return True
        End If
        Return MyBase.IgnoreLine(CurrLine)
    End Function

    Protected Overrides Function FixLine(ByVal CurrLine As String) As String
        CurrLine = MyBase.FixLine(CurrLine)
        Return CurrLine
    End Function

#End Region

End Class
