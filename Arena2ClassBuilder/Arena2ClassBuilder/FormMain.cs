// FormMain.cs - 05/08/2017

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
            textBoxFromPath.Text = fromPath;
            string toPath = $"{driveToolStripComboBox.Text}{Properties.Settings.Default.BaseToPath} {appToolStripComboBox.Text}";
            textBoxToPath.Text = toPath;
            BuildClasses(fromPath, toPath);
        }

        private void BuildClasses(string fromPath, string toPath)
        {
            DirectoryInfo fromDirInfo = new DirectoryInfo(fromPath);
            FileInfo[] fromFiles = fromDirInfo.GetFiles("*.sql");
            foreach (FileInfo fi in fromFiles)
            {
                if (!fi.Name.ToUpper().EndsWith(".TABLE.SQL"))
                {
                    return;
                }
                textBoxInput.Clear();
                textBoxInput.AppendText(File.ReadAllText(fi.FullName));
                bool isIDRIS = ((string)appToolStripComboBox.SelectedItem).ToUpper().Contains("IDRIS");
                string result = Builder.DoBuildClass(fi, isIDRIS);
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
            }
        }

    }
}
