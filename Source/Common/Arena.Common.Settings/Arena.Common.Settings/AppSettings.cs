// AppSettings.cs - 04/02/2018

using Arena.Common.Errors;
using Arena.Common.JSON;
using System;
using System.IO;

namespace Arena.Common.Settings
{
    public static class AppSettings
    {
        private const string _appConfigFilename = "_AppConfig.json";
        private const string _environmentKeyName = "Environment";

        private static string _settingsFilePath = null;
        private static JObject _appSettings = null;

        public static string GetEnvironmentName()
        {
            FillAppSettings();
            return (string)_appSettings.GetValueOrNull(_environmentKeyName) ?? "UNKNOWN";
        }

        private static void FillAppSettings()
        {
            if (_appSettings == null) // only fill this once 
            {
                FindSettingsFilePath(); // fills _settingsFilePath
                try
                {
                    _appSettings = JObject.Parse(File.ReadAllText($"{_settingsFilePath}{_appConfigFilename}"));
                }
                catch (Exception ex)
                {
                    throw new SystemException(ErrorHandler.FixMessage($"Error parsing settings file {_appConfigFilename}: {ex.Message}"));
                }
            }
        }

        private static void FindSettingsFilePath()
        {
            if (_settingsFilePath != null)
            {
                return;
            }
            if (File.Exists(_appConfigFilename))
            {
                _settingsFilePath = ""; // current directory, no trailing "\\"
                return;
            }
            string tempPath = Environment.CurrentDirectory;
            while (tempPath.Contains("\\"))
            {
                tempPath = tempPath.Substring(0, tempPath.LastIndexOf("\\"));
                if (File.Exists($"{tempPath}\\{_appConfigFilename}"))
                {
                    _settingsFilePath = $"{tempPath}\\";
                    return;
                }
                if (File.Exists($"{tempPath}\\Bin\\{_appConfigFilename}"))
                {
                    _settingsFilePath = $"{tempPath}\\Bin\\";
                    return;
                }
            }
            throw new SystemException(ErrorHandler.FixMessage($"File not found: {_appConfigFilename}"));
        }
    }
}
