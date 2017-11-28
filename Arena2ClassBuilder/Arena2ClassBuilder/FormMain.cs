// FormMain.cs - 11/28/2017

// --------------------------------------------------------------------------------------------------------------------
// 11/28/2017 - SBakker
//            - Better About message.
// 11/27/2017 - SBakker
//            - Moved application names into a Setting.
//            - Handle resulting filenames better.
// 11/22/2017 - SBakker
//            - Moved bootstrapping to Program.cs.
// 10/24/2017 - SBakker
//            - Added D: drive and changed paths to \Projects\...
// 10/17/2017 - SBakker
//            - Added error message when bootstrap fails.
// 09/29/2017 - SBakker
//            - Added Help About with run path.
//            - Bootstrapping must be done in FormMain_Load, not in constructor.
// 09/28/2017 - SBakker
//            - Added bootstrapping.
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Windows.Forms;

namespace Arena2ClassBuilder
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.Apps))
            {
                appToolStripComboBox.Items.Clear();
                string[] AppList = Properties.Settings.Default.Apps.Split(';');
                foreach (string app in AppList)
                {
                    if (!string.IsNullOrEmpty(app))
                    {
                        appToolStripComboBox.Items.Add(app);
                    }
                }
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.LastApp))
            {
                appToolStripComboBox.SelectedItem = Properties.Settings.Default.LastApp;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.LastDrive))
            {
                driveToolStripComboBox.SelectedItem = Properties.Settings.Default.LastDrive;
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void appToolStripComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (appToolStripComboBox.SelectedIndex >= 0)
            {
                Properties.Settings.Default.LastApp = (string)appToolStripComboBox.SelectedItem;
                Properties.Settings.Default.Save();
            }
        }

        private void driveToolStripComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (driveToolStripComboBox.SelectedIndex >= 0)
            {
                Properties.Settings.Default.LastDrive = (string)driveToolStripComboBox.SelectedItem;
                Properties.Settings.Default.Save();
            }
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (driveToolStripComboBox.SelectedIndex < 0 || appToolStripComboBox.SelectedIndex < 0)
            {
                return;
            }
            textBoxInput.Clear();
            textBoxOutput.Clear();
            string fromPath = $"{driveToolStripComboBox.Text}{Properties.Settings.Default.BaseFromPath} {appToolStripComboBox.Text}";
            if (!Directory.Exists(fromPath))
            {
                fromPath = fromPath.Substring(0, fromPath.Length - 1);
            }
            if (!Directory.Exists(fromPath))
            {
                MessageBox.Show($"FromPath not found: {fromPath}");
                return;
            }
            textBoxFromPath.Text = fromPath;
            string toPath = $"{driveToolStripComboBox.Text}{Properties.Settings.Default.BaseToPath} {appToolStripComboBox.Text}";
            if (!Directory.Exists(toPath))
            {
                MessageBox.Show($"ToPath not found: {toPath}");
                return;
            }
            textBoxToPath.Text = toPath;
            Application.DoEvents();
            BuildClasses(fromPath, toPath);
        }

        private void BuildClasses(string fromPath, string toPath)
        {
            int filesFound = 0;
            int filesChanged = 0;
            UpdateStatusBar(filesFound, filesChanged, false);
            DirectoryInfo fromDirInfo = new DirectoryInfo(fromPath);
            FileInfo[] fromFiles = fromDirInfo.GetFiles("*.sql");
            string productFamily = (string)appToolStripComboBox.SelectedItem;
            foreach (FileInfo fi in fromFiles)
            {
                if (!fi.Name.ToUpper().EndsWith(".TABLE.SQL"))
                {
                    continue;
                }
                if (fi.Name.Contains("#"))
                {
                    continue;
                }
                filesFound++;
                UpdateStatusBar(filesFound, filesChanged, false);
                textBoxInput.Clear();
                textBoxInput.AppendText(File.ReadAllText(fi.FullName));
                Application.DoEvents();
                string result = Builder.DoBuildClass(fi, productFamily);
                string outBaseName = fi.Name.Substring(0, fi.Name.Length - 10).Replace(" ", "_");
                // handle legacy databases
                if (!productFamily.Equals("Arena", StringComparison.OrdinalIgnoreCase)
                    && !productFamily.Equals("IDRIS", StringComparison.OrdinalIgnoreCase)
                    && !productFamily.Equals("Security", StringComparison.OrdinalIgnoreCase)
                    && !productFamily.Equals("TempData", StringComparison.OrdinalIgnoreCase))
                {
                    outBaseName = outBaseName.Replace("dbo.", $"dbo.{productFamily}_");
                    if (outBaseName.Contains($"dbo.{productFamily}_{productFamily}"))
                    {
                        outBaseName = outBaseName.Replace($"dbo.{productFamily}_{productFamily}", $"dbo.{productFamily}");
                    }
                }
                string outFileName = $"{toPath}\\{outBaseName}.cs";
                // don't write if file exists and matches
                if (File.Exists(outFileName))
                {
                    string tempOutfile = File.ReadAllText(outFileName);
                    if (result.Equals(tempOutfile))
                    {
                        continue;
                    }
                }
                File.WriteAllText(outFileName, result);
                textBoxOutput.Text = result;
                Application.DoEvents();
                filesChanged++;
                UpdateStatusBar(filesFound, filesChanged, false);
            }
            UpdateStatusBar(filesFound, filesChanged, true);
        }

        private void UpdateStatusBar(int filesFound, int filesChanged, bool done)
        {
            string doneMessage;
            if (done)
            {
                doneMessage = " - Done";
            }
            else
            {
                doneMessage = "";
            }
            toolStripStatusLabelMain.Text = $"Files Found: {filesFound} - Files Changed: {filesChanged}{doneMessage}";
            Application.DoEvents();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileInfo fileInfo = new FileInfo(Application.ExecutablePath);
            string version = fileInfo.LastWriteTime.ToString("yyyy.MM.dd.HHmm");
            MessageBox.Show($"{Environment.CurrentDirectory}\r\n\r\nVersion {version}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK);
        }
    }
}
