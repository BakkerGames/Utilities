' --------------------------------------
' --- BootstrapClass.vb - 07/29/2016 ---
' --------------------------------------

' ----------------------------------------------------------------------------------------------------
' 07/29/2016 - SBakker - URD 12917
'            - Try to perform bounce, using ".bootstrap" file to check if on second round.
' 07/21/2016 - SBakker - URD 12917
'            - Trying to fix "infinite bounce" issue.
'            - Added AltProgramPath to Arena_ConfigInfo, for use here.
' 04/29/2015 - SBakker
'            - Replaced Directory.Equals() with calls to NormalizePath(). Directory.Equals apparently
'              doesn't convert between mapped drives and UNC paths. NormalizePath does.
' 03/04/2015 - SBakker
'            - Convert pathnames to lowercase before comparing them.
' 07/29/2014 - SBakker
'            - Ignore ".settings" files in NeedDoubleBounce(), as they are never copied to the User's
'              Application directory. A newer version in the program directory was triggering an
'              infinite loop when calling the Main Menu.
' 07/21/2014 - SBakker
'            - Wrapped almost everything below in a Try/Catch, to see if the infinite bounce issue can
'              be caught.
' 06/23/2014 - SBakker
'            - Perform actual File.SetAttributes() before copying files to unset ReadOnly, instead of
'              just changing the internal flag. May have caused errors.
'            - Make sure the LaunchPath and ProgramPath don't match before doing double-bounce.
' 06/17/2014 - SBakker
'            - Added additional error checking to prevent infinite bounce.
'            - Only check for Double Bounce needed if the file length or file date changes. Don't
'              depend on the UTC time matching. This could be a cause of the infinite bounce.
' 04/17/2014 - SBakker
'            - Never copy ".settings" files!
' 03/24/2014 - SBakker
'            - Copy subdirectories also, which might hold resource or configuration files. Only goes
'              one level deep. Not sure how to make recursive Shared routines...
' 03/06/2014 - SBakker
'            - Added CommandLineArguments() to be passed to the new Process in ProcessStartInfo.
'            - Check for blank LaunchPath or ProgramPath.
'            - If any unknown error occurs, just exit quietly.
' 03/05/2014 - SBakker
'            - Added check for NeedDoubleBounce bootstrapping, where a program is started locally but
'              changes exists in original program directory.
' 02/14/2014 - SBakker
'            - Make sure copied files get Normal attributes afterwards.
' 10/24/2013 - SBakker
'            - Building Bootstrap class to move all programs and associated files to the LaunchPath
'              location, if it isn't run from the LaunchPath location.
' ----------------------------------------------------------------------------------------------------

Imports Arena_ConfigInfo.ArenaConfigInfo
Imports Arena_Utilities.AppUtils
Imports Arena_Utilities.SystemUtils
Imports System.IO

Public Class BootstrapClass

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private Const FudgeSeconds As Integer = 5 ' Seconds of difference allowed for file compare

    Public Shared Function CopyProgramsToLaunchPath() As Boolean

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim Normalize_MyAppPath As String
        Dim Normalize_LaunchPath As String
        Dim Normalize_ProgramPath As String
        Dim Normalize_AltProgramPath As String
        ' ------------------------------------

        ' --- Must return false when debugging, or it will close this app and launch the target one! ---
        ' --- Manually skip this section if you really want to debug below ---
#If DEBUG Then
        Return False
#End If
        If Diagnostics.Debugger.IsAttached Then
            Return False
        End If

        ' --- Check if Arena.xml is configured for Bootstrapping ---
        If Bootstrap = False Then
            Return False
        End If

        If String.IsNullOrEmpty(ProgramPath) OrElse String.IsNullOrEmpty(LaunchPath) Then
            Return False
        End If

        Try
#If DEBUG Then
            Normalize_MyAppPath = NormalizePath(ProgramPath)
#Else
            Normalize_MyAppPath = NormalizePath(My.Application.Info.DirectoryPath)
#End If
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Error reading MyAppPath" + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        Try

            Normalize_ProgramPath = NormalizePath(ProgramPath)
            Normalize_AltProgramPath = NormalizePath(AltProgramPath)
            Normalize_LaunchPath = NormalizePath(LaunchPath)

            If Normalize_LaunchPath = Normalize_ProgramPath OrElse
               Normalize_LaunchPath = Normalize_AltProgramPath Then
                Return False
            End If

            If Normalize_MyAppPath <> Normalize_ProgramPath AndAlso
               Normalize_MyAppPath <> Normalize_AltProgramPath AndAlso
               Normalize_MyAppPath <> Normalize_LaunchPath Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                          "MyAppPath doesn't equal either ProgramPath or LaunchPath:" + vbCrLf +
                                          "MyAppPath = '" + Normalize_MyAppPath + "'" + vbCrLf +
                                          "ProgramPath = '" + Normalize_ProgramPath + "'" + vbCrLf +
                                          "AltProgramPath = '" + Normalize_AltProgramPath + "'" + vbCrLf +
                                          "LaunchPath = '" + Normalize_LaunchPath + "'" + vbCrLf)
            End If

        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Error reading LaunchPath or ProgramPath" + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        ' --- Check if current app in LaunchPath is older than ProgramPath ---
        If Normalize_MyAppPath = Normalize_LaunchPath Then
            ' --- Don't do anything if there is no bootstrapping ---
            If Normalize_MyAppPath = Normalize_ProgramPath OrElse Normalize_MyAppPath = Normalize_AltProgramPath Then
                Return False
            End If
            ' --- Save filenames for use below ---
            Dim CurrAppFilename As String = $"{Normalize_MyAppPath}{My.Application.Info.AssemblyName}.exe"
            Dim OrigAppFilename As String = $"{Normalize_ProgramPath}{My.Application.Info.AssemblyName}.exe"
            Dim CurrBootstrapName As String = $"{Normalize_MyAppPath}{My.Application.Info.AssemblyName}.bootstrap"
            ' --- Don't do anything if the files are not found ---
            If Not File.Exists(CurrAppFilename) OrElse Not File.Exists(OrigAppFilename) Then
                Return False
            End If
            ' --- Check if the current file is newer or the same datetime ---
            Dim CurrFileInfo As New FileInfo(CurrAppFilename)
            Dim SourceFileInfo As New FileInfo(OrigAppFilename)
            If DateAdd(DateInterval.Second, FudgeSeconds, CurrFileInfo.LastWriteTimeUtc) >= SourceFileInfo.LastWriteTimeUtc Then
                If File.Exists(CurrBootstrapName) Then
                    File.SetAttributes(CurrBootstrapName, FileAttributes.Normal)
                    File.Delete(CurrBootstrapName)
                End If
                Return False
            End If
            ' --- Current file is older so needs an update ---
            If File.Exists(CurrBootstrapName) Then
                ' --- Have already bounced once, so throw an error ---
                File.SetAttributes(CurrBootstrapName, FileAttributes.Normal)
                File.Delete(CurrBootstrapName)
                Throw New SystemException($"Application needs update: {CurrAppFilename}")
            End If
            ' --- Bounce back to newer version of this app on ProgramPath ---
            File.WriteAllText(CurrBootstrapName, "Bootstrap bounce check file")
            Dim OrigFileProcess As New ProcessStartInfo(OrigAppFilename)
            OrigFileProcess.Arguments = CommandLineArguments()
            If Process.Start(OrigFileProcess) Is Nothing Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                          "Cannot start application: " + OrigAppFilename)
            End If
            Return True
        End If

        Try
            If Not Directory.Exists(Normalize_LaunchPath) Then
                Directory.CreateDirectory(Normalize_LaunchPath)
            End If
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Unable to create path: " + Normalize_LaunchPath + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        Try
            CopyFiles(Normalize_MyAppPath, Normalize_LaunchPath)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Error copying files from """ + Normalize_MyAppPath + """ to """ + Normalize_LaunchPath + """" + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        Try
            CopySubdirectories(Normalize_MyAppPath, Normalize_LaunchPath)
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Error copying subdirectories from """ + Normalize_MyAppPath + """ to """ + Normalize_LaunchPath + """" + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        Dim MyFilename As String = Normalize_LaunchPath + My.Application.Info.AssemblyName + ".exe"

        ' --- Start this application in the new folder ---
        If String.IsNullOrEmpty(MyFilename) Then
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Current application not found: " + My.Application.Info.AssemblyName + ".exe")
        End If
        If Not File.Exists(MyFilename) Then
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Application not found: " + MyFilename)
        End If

        Try
            Dim NewFileProcess As New ProcessStartInfo(MyFilename)
            NewFileProcess.Arguments = CommandLineArguments()
            If Process.Start(NewFileProcess) Is Nothing Then
                Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                          "Cannot start application: " + MyFilename)
            End If
        Catch ex As Exception
            Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                      "Cannot start application: " + MyFilename + vbCrLf + vbCrLf +
                                      ex.Message)
        End Try

        Return True

    End Function

    Private Shared Sub CopySubdirectories(ByVal FromPath As String, ByVal ToPath As String)

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim CurrSubdirectories() As String = Directory.GetDirectories(FromPath)

        Dim NewToPath As String = ""
        For Each NewFromPath As String In CurrSubdirectories

            ' --- Keep all directories ending in "\" ---
            If Not NewFromPath.EndsWith("\") Then
                NewFromPath += "\"
            End If

            NewToPath = ToPath + NewFromPath.Substring(FromPath.Length)

            Try
                If Not Directory.Exists(NewToPath) Then
                    Directory.CreateDirectory(NewToPath)
                End If
            Catch ex As Exception
                Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                          "Unable to create path: " + NewToPath + vbCrLf + vbCrLf +
                                          ex.Message)
            End Try

            CopyFiles(NewFromPath, NewToPath)

        Next

    End Sub

    Private Shared Sub CopyFiles(ByVal FromPath As String, ByVal ToPath As String)

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Dim CurrFiles() As String = Directory.GetFiles(FromPath)
        Dim TargetFilename As String = ""
        Dim CurrFileInfo As FileInfo
        Dim TargetFileInfo As FileInfo

        For Each CurrFilename As String In CurrFiles
            Try
                If (File.GetAttributes(CurrFilename) And FileAttributes.Hidden) = FileAttributes.Hidden Then
                    Continue For
                End If
                ' --- Never copy settings files! ---
                If CurrFilename.ToLower.EndsWith(".settings") Then Continue For
                TargetFilename = ToPath + CurrFilename.Substring(FromPath.Length)
                If Not File.Exists(TargetFilename) Then
                    File.Copy(CurrFilename, TargetFilename, True)
                    File.SetAttributes(TargetFilename, FileAttributes.Normal)
                Else
                    CurrFileInfo = New FileInfo(CurrFilename)
                    TargetFileInfo = New FileInfo(TargetFilename)
                    If TargetFileInfo.Length <> CurrFileInfo.Length OrElse
                       TargetFileInfo.LastWriteTimeUtc < CurrFileInfo.LastWriteTimeUtc Then
                        File.SetAttributes(TargetFilename, FileAttributes.Normal)
                        File.Copy(CurrFilename, TargetFilename, True)
                        File.SetAttributes(TargetFilename, FileAttributes.Normal)
                    End If
                End If
            Catch ex As Exception
                Throw New SystemException(FuncName + vbCrLf + vbCrLf +
                                          "Error copying file """ + CurrFilename + """ to """ + TargetFilename + """" + vbCrLf + vbCrLf +
                                          ex.Message)
            End Try
        Next

    End Sub

End Class
