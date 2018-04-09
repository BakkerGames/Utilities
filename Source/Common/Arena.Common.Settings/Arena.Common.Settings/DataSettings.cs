// DataSettings.cs - 04/02/2018

using Arena.Common.Errors;
using Arena.Common.JSON;
using System;
using System.IO;

namespace Arena.Common.Settings
{
    public static class DataSettings
    {
        private const string _dataSettingsFilename = "_DataSettings.json";
        private const string _productFamilyNodeName = "productfamily";
        private const string _serverKeyName = "servername";
        private const string _databaseKeyName = "databasename";

        private static string _settingsFilePath = null;
        private static JObject _dataSettings = null;

        public static string GetServerName(string productFamily)
        {
            FillDataSettings();
            return (string)((JObject)((JObject)_dataSettings
                .GetValue(_productFamilyNodeName))
                .GetValue(productFamily))
                .GetValueOrNull(_serverKeyName);
        }

        public static string GetDatabaseName(string productFamily)
        {
            FillDataSettings();
            return (string)((JObject)((JObject)_dataSettings
                .GetValue(_productFamilyNodeName))
                .GetValue(productFamily))
                .GetValueOrNull(_databaseKeyName);
        }

        private static void FillDataSettings()
        {
            if (_dataSettings == null) // only fill this once 
            {
                FindSettingsFilePath(); // fills _settingsFilePath
                try
                {
                    _dataSettings = JObject.Parse(File.ReadAllText($"{_settingsFilePath}{_dataSettingsFilename}"));
                }
                catch (Exception ex)
                {
                    throw new SystemException(ErrorHandler.FixMessage($"Error parsing settings file {_dataSettingsFilename}: {ex.Message}"));
                }
            }
        }

        private static void FindSettingsFilePath()
        {
            if (_settingsFilePath != null)
            {
                return;
            }
            if (File.Exists(_dataSettingsFilename))
            {
                _settingsFilePath = ""; // current directory, no trailing "\\"
                return;
            }
            string tempPath = Environment.CurrentDirectory;
            while (tempPath.Contains("\\"))
            {
                tempPath = tempPath.Substring(0, tempPath.LastIndexOf("\\"));
                if (File.Exists($"{tempPath}\\{_dataSettingsFilename}"))
                {
                    _settingsFilePath = $"{tempPath}\\";
                    return;
                }
                if (File.Exists($"{tempPath}\\Bin\\{_dataSettingsFilename}"))
                {
                    _settingsFilePath = $"{tempPath}\\Bin\\";
                    return;
                }
            }
            throw new SystemException(ErrorHandler.FixMessage($"File not found: {_dataSettingsFilename}"));
        }
    }
}
