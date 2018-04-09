// Bootstrapper.cs - 04/04/2018

// ----------------------------------------------------------------------------------------------------------
// 04/03/2018 - SBakker
//            - Added proper checking for null lists.
// 04/02/2018 - SBakker
//            - Use ErrorHandler.FixMessage() to send call stack info upwards.
// 03/29/2018 - SBakker
//            - Added OtherLaunchPaths so Arena can call Arena2 programs.
// 03/19/2018 - SBakker
//            - Simplified error messages to avoid call stacks returning.
// 02/10/2018 - SBakker
//            - Added full path to error message, File not found.
// 01/03/2018 - SBakker
//            - Added LaunchApplication() routine for calls sideways between apps. Used by CallApp.BusinessLogic.
//            - Changed returned values to Process instead of void so calling programs can check processes.
//            - Renamed LaunchProgram() to DoLaunchProgram() to avoid confusion.
//            - Added commandline args to MustBootstrap() for sending to launched app.
// 11/08/2017 - SBakker - URD 15244
//            - Ignore directories which are invalid/missing.
// 10/17/2017 - SBakker - URD 15244
//            - Added more descriptive errors on File.Copy and File.SetAttributes.
//            - Added ErrorHandler.FixMessage() to handle error messages.
// 10/11/2017 - SBakker - URD 15244
//            - Changed launch path from {envName}\\{appName} to {appName}_{envName}. Now it will match the
//              existing Arena directories, so any links or settings will still work.
// 10/06/2017 - SBakker - URD 15244
//            - Adding double-bounce so updates are installed quietly.
//            - Ignore file datetime differences less than one second, just in case.
// 09/27/2017 - SBakker - URD 15244
//            - Added CopyRecursive property to prevent copying undesired files.
//            - Fixed Process.Start to create a new process instead of re-using this one, which loops forever.
// 09/21/2017 - SBakker - URD 15244
//            - Created new bootstrapping routine. Handles multiple master program locations and other
//              application locations, so they will all be copied to USERPROFILE and run there.
// ----------------------------------------------------------------------------------------------------------

using Arena.Common.Errors;
using Arena.Common.JSON;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Arena.Common.Bootstrap
{
    public static class Bootstrapper
    {
        private const string _appConfigFilename = "_AppConfig.json";
        private const string _userAppSubdir = "Applications";
        private static string _baseLaunchPath;

        static Bootstrapper()
        {
            _baseLaunchPath = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\{_userAppSubdir}";
        }

        public static bool MustBootstrap()
        {
            return MustBootstrap(null);
        }

        public static bool MustBootstrap(string[] args)
        {
            string joinedArgs = (args == null) ? "" : string.Join(" ", args);
            string currPath = Environment.CurrentDirectory;
            if (currPath.ToLower().Contains("\\debug\\") || currPath.ToLower().EndsWith("\\debug"))
            {
                return false; // no bootstrapping when debugging
            }
            // not debugging
            bool result = true;
            string mainAppConfigPath = GetSettingsFilePath();
            BootstrapAppConfig appConfig = GetSettingInfo(mainAppConfigPath);
            if (currPath.StartsWith(Environment.GetEnvironmentVariable("USERPROFILE"), StringComparison.OrdinalIgnoreCase))
            {
                result = false;
            }
            if (!result)
            {
                // check if master program is newer/different than this program
                foreach (string path in appConfig.AppPaths)
                {
                    string masterAppFilePath = $"{path}\\{AppDomain.CurrentDomain.FriendlyName}";
                    if (File.Exists(masterAppFilePath))
                    {
                        FileInfo currAppInfo = new FileInfo($"{currPath}\\{AppDomain.CurrentDomain.FriendlyName}");
                        FileInfo masterAppInfo = new FileInfo(masterAppFilePath);
                        if (currAppInfo.Length != masterAppInfo.Length
                            // ignore differences less than one second, in case server filesystems are different
                            || currAppInfo.LastWriteTimeUtc.Ticks + 10000000 < masterAppInfo.LastWriteTimeUtc.Ticks)
                        {
                            // double-bounce
                            DoLaunchProgram(path, AppDomain.CurrentDomain.FriendlyName, joinedArgs);
                            return true;
                        }
                    }
                }
            }
            if (result)
            {
                // does need normal bootstrapping
                CopyProgramsToLaunchPath(appConfig.AppPaths, appConfig.FullLaunchPath, appConfig.CopyRecursive ?? false);
                CopyOtherPrograms(appConfig.OtherAppPaths);
                DoLaunchProgram(appConfig.FullLaunchPath, AppDomain.CurrentDomain.FriendlyName, joinedArgs);
            }
            return result;
        }

        public static Process LaunchApplication(string appName, string arguments)
        {
            string mainAppConfigPath = GetSettingsFilePath();
            BootstrapAppConfig appConfig = GetSettingInfo(mainAppConfigPath);
            string launchPath = appConfig.FullLaunchPath;
            if (!File.Exists($"{appConfig.FullLaunchPath}\\{appName}.exe"))
            {
                launchPath = null;
                foreach (string tempLaunchPath in appConfig.OtherLaunchPaths)
                {
                    if (File.Exists($"{tempLaunchPath}\\{appName}.exe"))
                    {
                        launchPath = tempLaunchPath;
                        break;
                    }
                }
                if (launchPath == null)
                {
                    throw new FileNotFoundException(ErrorHandler.FixMessage($"File not found: {appConfig.FullLaunchPath}\\{appName}.exe"));
                }
            }
            return DoLaunchProgram(launchPath, $"{appName}.exe", arguments);
        }

        #region Private routines

        private static void CopyProgramsToLaunchPath(List<string> appPaths, string fullLaunchPath, bool copyRecursive)
        {
            foreach (string path in appPaths)
            {
                try
                {
                    if (!Directory.Exists(path))
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                CopyAll(path, fullLaunchPath, copyRecursive);
            }
        }

        private static void CopyOtherPrograms(List<string> otherAppPaths)
        {
            foreach (string path in otherAppPaths)
            {
                if (File.Exists($"{path}\\{_appConfigFilename}"))
                {
                    BootstrapAppConfig tempConfig = GetSettingInfo(path);
                    CopyProgramsToLaunchPath(tempConfig.AppPaths, tempConfig.FullLaunchPath, tempConfig.CopyRecursive ?? false);
                    CopyOtherPrograms(tempConfig.OtherAppPaths);
                }
            }
        }

        private static Process DoLaunchProgram(string fullLaunchPath, string fileName, string arguments)
        {
            if (!File.Exists($"{fullLaunchPath}\\{fileName}"))
            {
                throw new FileNotFoundException(ErrorHandler.FixMessage($"File not found: {fullLaunchPath}\\{fileName}"));
            }
            Process newApp = new Process();
            newApp.StartInfo.UseShellExecute = false;
            newApp.StartInfo.CreateNoWindow = true;
            newApp.StartInfo.WorkingDirectory = fullLaunchPath;
            newApp.StartInfo.FileName = $"{fullLaunchPath}\\{fileName}";
            newApp.StartInfo.Arguments = arguments ?? "";
            newApp.Start();
            return newApp;
        }

        private static void CopyAll(string fromPath, string toPath, bool copyRecursive)
        {
            if (!Directory.Exists(toPath))
            {
                Directory.CreateDirectory(toPath);
            }
            foreach (string filename in Directory.EnumerateFiles(fromPath))
            {
                if (filename.StartsWith("."))
                {
                    continue; // skip dot files
                }
                FileInfo currFileInfo = new FileInfo(filename);
                if ((currFileInfo.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                {
                    continue; // skip hidden files
                }
                if (currFileInfo.Extension.ToLower() == "settings")
                {
                    continue; // don't copy these files
                }
                string targetFilename = $"{toPath}\\{currFileInfo.Name}";
                if (File.Exists(targetFilename))
                {
                    FileInfo targetFileInfo = new FileInfo(targetFilename);
                    if (targetFileInfo.Length == currFileInfo.Length
                        && targetFileInfo.LastWriteTimeUtc >= currFileInfo.LastWriteTimeUtc)
                    {
                        continue;
                    }
                    try
                    {
                        File.SetAttributes(targetFilename, FileAttributes.Normal);
                    }
                    catch (Exception ex)
                    {
                        throw new SystemException(ErrorHandler.FixMessage($"Error setting file attributes on {targetFilename}\r\n\r\n{ex.Message}"));
                    }
                }
                try
                {
                    File.Copy(filename, targetFilename, true);
                }
                catch (Exception ex)
                {
                    throw new SystemException(ErrorHandler.FixMessage($"Error copying file {filename} to {targetFilename}\r\n\r\n{ex.Message}"));
                }
            }
            if (copyRecursive)
            {
                foreach (string dirName in Directory.EnumerateDirectories(fromPath))
                {
                    string simpleDirName = dirName.Substring(dirName.LastIndexOf("\\") + 1);
                    if (simpleDirName.StartsWith("."))
                    {
                        continue; // skip dot directories
                    }
                    DirectoryInfo currDirInfo = new DirectoryInfo(dirName);
                    if ((currDirInfo.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                    {
                        continue; // skip hidden directories
                    }
                    CopyAll($"{fromPath}\\{simpleDirName}", $"{toPath}\\{simpleDirName}", copyRecursive);
                }
            }
        }

        private static string GetSettingsFilePath()
        {
            if (File.Exists(_appConfigFilename))
            {
                return "."; // current directory
            }
            string tempPath = Environment.CurrentDirectory;
            while (tempPath.Contains("\\"))
            {
                tempPath = tempPath.Substring(0, tempPath.LastIndexOf("\\"));
                if (File.Exists($"{tempPath}\\{_appConfigFilename}"))
                {
                    return $"{tempPath}";
                }
                if (File.Exists($"{tempPath}\\Bin\\{_appConfigFilename}"))
                {
                    return $"{tempPath}\\Bin";
                }
            }
            throw new SystemException(ErrorHandler.FixMessage($"File not found: {Environment.CurrentDirectory}\\{_appConfigFilename}"));
        }

        private static BootstrapAppConfig GetSettingInfo(string appConfigPath)
        {
            BootstrapAppConfig result = new BootstrapAppConfig();
            JObject appConfigSettings = JObject.Parse(File.ReadAllText($"{appConfigPath}\\{_appConfigFilename}"));
            string envName = (string)appConfigSettings.GetValueOrNull("Environment");
            string appName = (string)appConfigSettings.GetValueOrNull("Application");
            result.FullLaunchPath = $"{_baseLaunchPath}\\{appName}_{envName}";
            result.CopyRecursive = (bool?)appConfigSettings.GetValueOrNull("CopyRecursive");
            result.OtherLaunchPaths = new List<string>();
            result.AppPaths = new List<string>();
            result.OtherAppPaths = new List<string>();
            if (appConfigSettings.GetValueOrNull("OtherApplications") != null)
            {
                foreach (string tempAppName in (JArray)appConfigSettings.GetValueOrNull("OtherApplications"))
                {
                    result.OtherLaunchPaths.Add($"{_baseLaunchPath}\\{tempAppName}_{envName}");
                }
            }
            if (appConfigSettings.GetValueOrNull("AppPaths") != null)
            {
                foreach (string tempPath in (JArray)appConfigSettings.GetValueOrNull("AppPaths"))
                {
                    result.AppPaths.Add(tempPath);
                };
            }
            if (appConfigSettings.GetValueOrNull("OtherAppPaths") != null)
            {
                foreach (string tempPath in (JArray)appConfigSettings.GetValueOrNull("OtherAppPaths"))
                {
                    result.OtherAppPaths.Add(tempPath);
                };
            }
            return result;
        }

        #endregion

    }
}
