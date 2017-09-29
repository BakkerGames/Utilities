// FormMain.cs - 09/29/2017

// --------------------------------------------------------------------------------------------------------------------
// 09/29/2017 - SBakker
//            - Added Help About with run path.
//            - Bootstrapping must be done in FormMain_Load, not in constructor.
// 09/28/2017 - SBakker
//            - Added bootstrapping.
// --------------------------------------------------------------------------------------------------------------------

using Arena.Common.Bootstrap;
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
            if (Bootstrapper.MustBootstrap())
            {
                Close();
                return;
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
                string outFileName = $"{toPath}\\{fi.Name.Substring(0, fi.Name.Length - 10)}.cs";
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
            MessageBox.Show(Environment.CurrentDirectory, AppDomain.CurrentDomain.FriendlyName, MessageBoxButtons.OK);
        }
    }
}
