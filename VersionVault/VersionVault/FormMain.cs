// FormMain.cs - 02/23/2017

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Common.JSON;
using System.Diagnostics;
using VVLibrary;

namespace VersionVault
{
    public partial class FormMain : Form
    {
        private const string configFileName = ".vvconfig";
        private JSONObject vvConfig;
        private JSONArray sourceFolders;

        private TreeNode selectedTreeViewNode = null;
        private int selectedListViewIndex = -1;

        private Vault myVault = new Vault();

        public FormMain()
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(Properties.Settings.Default.SourcePathArray))
            {
                sourceFolders = JSONArray.FromString(Properties.Settings.Default.SourcePathArray);
                foreach (string s in sourceFolders)
                {
                    comboBoxFolder.Items.Add(s);
                }
                for (int i = 0; i < comboBoxFolder.Items.Count; i++)
                {
                    if (comboBoxFolder.Items[i].ToString().ToLower() == Properties.Settings.Default.LastSourcePath.ToLower())
                    {
                        comboBoxFolder.SelectedIndex = i;
                        break;
                    }
                }
            }
            else
            {
                sourceFolders = new JSONArray();
            }
        }

        private void buttonGo_Click(object sender, EventArgs e)
        {
            string enteredDir = comboBoxFolder.Text;
            if (string.IsNullOrEmpty(enteredDir))
            {
                return;
            }
            while (enteredDir.EndsWith("\\"))
            {
                enteredDir = enteredDir.Substring(0, enteredDir.Length - 1);
            }
            if (!Directory.Exists(enteredDir))
            {
                MessageBox.Show($"Directory not found:\r\n{enteredDir}", "Error");
                return;
            }
            // add to comboBox if not already there
            bool foundItem = false;
            foreach (string s in comboBoxFolder.Items)
            {
                if (s.ToLower() == enteredDir.ToLower())
                {
                    foundItem = true;
                    break;
                }
            }
            if (!foundItem)
            {
                comboBoxFolder.Items.Add(enteredDir);
                sourceFolders.Add(enteredDir);
                Properties.Settings.Default.SourcePathArray = sourceFolders.ToString();
                Properties.Settings.Default.LastSourcePath = enteredDir;
                Properties.Settings.Default.Save();
            }
            treeViewMainClear();
            listViewMainClear();
            listBoxVVClear();
            Application.DoEvents();
            if (File.Exists($"{enteredDir}\\{configFileName}"))
            {
                vvConfig = JSONObject.FromString(File.ReadAllText($"{enteredDir}\\{configFileName}"));
            }
            else
            {
                MessageBox.Show($"{enteredDir}\\{configFileName} not found", "Error");
                return;
            }
            FillTreeView(enteredDir);
            //backupToolStripMenuItem.Enabled = true;
        }

        private void FillTreeView(string path)
        {
            treeViewMainClear();
            treeViewMain.BeginUpdate();
            foreach (string subdirectory in Directory.GetDirectories(path))
            {
                if (IgnorePath(subdirectory))
                {
                    continue;
                }
                treeViewMain.Nodes.Add(TraverseDirectory(subdirectory));
            }
            treeViewMain.EndUpdate();
        }

        private TreeNode TraverseDirectory(string path)
        {
            if (IgnorePath(path))
            {
                return null;
            }
            TreeNode result = new TreeNode(path);
            result.Text = PathBase(path);
            foreach (string subdirectory in Directory.GetDirectories(path))
            {
                if (IgnorePath(subdirectory))
                {
                    continue;
                }
                result.Nodes.Add(TraverseDirectory(subdirectory));
            }
            return result;
        }

        private string PathBase(string path)
        {
            int posSlash = path.LastIndexOf("\\");
            string pathBase;
            if (posSlash >= 0)
            {
                pathBase = path.Substring(posSlash + 1, path.Length - posSlash - 1);
            }
            else
            {
                pathBase = path;
            }
            return pathBase;
        }

        private bool IgnorePath(string path)
        {
            string pathBase = PathBase(path);
            if (pathBase.StartsWith(".")) { return true; }
            DirectoryInfo di = new DirectoryInfo(path);
            if ((di.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
            {
                return true;
            }
            JSONArray ignoreDir = (JSONArray)vvConfig["IgnoreDir"];
            foreach (string dirName in ignoreDir)
            {
                if (pathBase == dirName)
                {
                    return true;
                }
            }
            return false;
        }

        private bool IgnoreFile(string filename)
        {
            if (filename.StartsWith(".")) { return true; }
            string lowerFilename = filename.ToLower();
            JSONArray ignoreExt = (JSONArray)vvConfig["IgnoreExt"];
            foreach (string extName in ignoreExt)
            {
                if (lowerFilename.EndsWith(extName.ToLower()))
                {
                    return true;
                }
            }
            return false;
        }

        private void treeViewMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (selectedTreeViewNode != null)
            {
                selectedTreeViewNode.BackColor = SystemColors.Window;
                selectedTreeViewNode.ForeColor = SystemColors.WindowText;
                selectedTreeViewNode = null;
            }
            selectedTreeViewNode = treeViewMain.SelectedNode;
            selectedTreeViewNode.BackColor = SystemColors.Highlight;
            selectedTreeViewNode.ForeColor = SystemColors.HighlightText;
            listViewMainClear();
            listBoxVVClear();
            FileInfo fi;
            string sourceDir = $"{comboBoxFolder.Text}\\{treeViewMain.SelectedNode.FullPath}";
            foreach (string filename in Directory.GetFiles(sourceDir))
            {
                if (IgnoreFile(filename))
                {
                    continue;
                }
                fi = new FileInfo(filename);
                ListViewItem item = new ListViewItem(PathBase(filename));
                item.SubItems.Add(fi.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss"));
                item.SubItems.Add(fi.Length.ToString());
                listViewMain.Items.Add(item);
            }
        }

        private void treeViewMainClear()
        {
            selectedTreeViewNode = null;
            treeViewMain.Nodes.Clear();
        }

        private void listViewMainClear()
        {
            selectedListViewIndex = -1;
            listViewMain.Items.Clear();
        }

        private void listBoxVVClear()
        {
            listBoxVV.Items.Clear();
        }

        private void listViewMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selectedListViewIndex >= 0 && selectedListViewIndex < listViewMain.Items.Count)
            {
                listViewMain.Items[selectedListViewIndex].BackColor = SystemColors.Window;
                listViewMain.Items[selectedListViewIndex].ForeColor = SystemColors.WindowText;
            }
            selectedListViewIndex = -1;
            listBoxVVClear();
            ListView.SelectedListViewItemCollection files = listViewMain.SelectedItems;
            if (files.Count == 0)
            {
                return;
            }
            if (vvConfig != null && vvConfig.ContainsKey("VVPath"))
            {
                string vvDir = $"{(string)vvConfig["VVPath"]}\\{treeViewMain.SelectedNode.FullPath}\\{files[0].Text}";
                if (Directory.Exists(vvDir))
                {
                    foreach (string filename in Directory.GetFiles(vvDir))
                    {
                        VVItem item = new VVItem(PathBase(filename));
                        listBoxVV.Items.Add(item);
                    }
                }
            }
        }

        private void listViewMain_Resize(object sender, EventArgs e)
        {
            int minWidth = listViewMain.Columns[1].Width + listViewMain.Columns[2].Width + 30;
            if (listViewMain.Width >= minWidth)
            {
                listViewMain.Columns[0].Width = listViewMain.Width - minWidth;
            }
        }

        private void listViewMain_Leave(object sender, EventArgs e)
        {
            int currIndex = listViewMain.FocusedItem.Index;
            if (currIndex >= 0)
            {
                listViewMain.Items[currIndex].BackColor = SystemColors.Highlight;
                listViewMain.Items[currIndex].ForeColor = SystemColors.HighlightText;
                selectedListViewIndex = currIndex;
            }
        }

        private void listBoxVV_DoubleClick(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection files = listViewMain.SelectedItems;
            string sourceDir = $"{comboBoxFolder.Text}\\{treeViewMain.SelectedNode.FullPath}";
            string sourceFile = $"{sourceDir}\\{files[0].Text}";
            string vvDir = $"{(string)vvConfig["VVPath"]}\\{treeViewMain.SelectedNode.FullPath}\\{files[0].Text}";
            string vvFile = $"{vvDir}\\{((VVItem)listBoxVV.SelectedItem).ItemName}";
            // launch ExternalCompareApp process with the two files
            string compareApp = Properties.Settings.Default.ExternalCompareApp;
            string compareOptions = Properties.Settings.Default.ExternalCompareAppOptions;
            Process p = Process.Start(compareApp, $"{compareOptions} \"{sourceFile}\" \"{vvFile}\"");
            p.WaitForExit();
            int result = p.ExitCode;
        }

        private void externalCompareAppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormExternalCompareApp popup = new FormExternalCompareApp();
            popup.PathToEXE = Properties.Settings.Default.ExternalCompareApp;
            popup.Options = Properties.Settings.Default.ExternalCompareAppOptions;
            popup.ShowDialog();
            if (popup.DialogResult != DialogResult.OK)
            {
                return;
            }
            Properties.Settings.Default.ExternalCompareApp = popup.PathToEXE;
            Properties.Settings.Default.ExternalCompareAppOptions = popup.Options;
            Properties.Settings.Default.Save();
        }

        private void toolStripMenuItemBackup_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not implemented yet!");
        }

        //private void backupToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    myVault.SourcePath = comboBoxFolder.Text;
        //    myVault.VaultPath = (string)vvConfig["VVPath"];
        //    myVault.BackupAll();
        //    MessageBox.Show("Files vaulted", "Backup complete");
        //}

    }
}
