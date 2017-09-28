// Bootstrapper.cs - 09/27/2017

// ----------------------------------------------------------------------------------------------------------
// 09/27/2017 - SBakker - URD 15244
//            - Added CopyRecursive property to prevent copying undesired files.
//            - Fixed Process.Start to create a new process instead of re-using this one, which loops forever.
// 09/21/2017 - SBakker - URD 15244
//            - Created new bootstrapping routine. Handles multiple master program locations and other
//              application locations, so they will all be copied to USERPROFILE and run there.
// ----------------------------------------------------------------------------------------------------------

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
                            || currAppInfo.LastWriteTimeUtc < masterAppInfo.LastWriteTimeUtc)
                        {
                            throw new SystemException($"Application needs update, please restart:\r\n{AppDomain.CurrentDomain.FriendlyName}");
                        }
                    }
                }
            }
            if (result)
            {
                // does need normal bootstrapping
                CopyProgramsToLaunchPath(appConfig.AppPaths, appConfig.FullLaunchPath, appConfig.CopyRecursive ?? false);
                CopyOtherPrograms(appConfig.OtherAppPaths);
                LaunchUserCopyOfProgram(appConfig.FullLaunchPath, "");
            }
            return result;
        }

        #region Private routines

        private static void CopyProgramsToLaunchPath(List<string> appPaths, string fullLaunchPath, bool copyRecursive)
        {
            foreach (string path in appPaths)
            {
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

        private static void LaunchUserCopyOfProgram(string fullLaunchPath, string arguments)
        {
            Process newApp = new Process();
            newApp.StartInfo.UseShellExecute = false;
            newApp.StartInfo.CreateNoWindow = true;
            newApp.StartInfo.WorkingDirectory = fullLaunchPath;
            newApp.StartInfo.FileName = $"{fullLaunchPath}\\{AppDomain.CurrentDomain.FriendlyName}";
            newApp.StartInfo.Arguments = arguments;
            newApp.Start();
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
                    File.SetAttributes(targetFilename, FileAttributes.Normal);
                }
                File.Copy(filename, targetFilename, true);
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
            throw new SystemException($"File not found: {_appConfigFilename}");
        }

        private static BootstrapAppConfig GetSettingInfo(string appConfigPath)
        {
            BootstrapAppConfig result = new BootstrapAppConfig();
            JObject appConfigSettings = JObject.Parse(File.ReadAllText($"{appConfigPath}\\{_appConfigFilename}"));
            string envName = (string)appConfigSettings.GetValueOrNull("Environment");
            string appName = (string)appConfigSettings.GetValueOrNull("Application");
            result.FullLaunchPath = $"{_baseLaunchPath}\\{envName}\\{appName}";
            result.AppPaths = new List<string>();
            foreach (string tempPath in (JArray)appConfigSettings.GetValueOrNull("AppPaths"))
            {
                result.AppPaths.Add(tempPath);
            };
            result.CopyRecursive = (bool?)appConfigSettings.GetValueOrNull("CopyRecursive");
            result.OtherAppPaths = new List<string>();
            foreach (string tempPath in (JArray)appConfigSettings.GetValueOrNull("OtherAppPaths"))
            {
                result.OtherAppPaths.Add(tempPath);
            };
            return result;
        }

        #endregion

    }
}
