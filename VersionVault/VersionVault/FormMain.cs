// FormMain.cs - 11/28/2017

using Arena.Common.JSON;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using VVLibrary;

namespace VersionVault
{
    public partial class FormMain : Form
    {
        private bool _loading = true;

        private const string configFileName = ".vvconfig";
        private JObject vvConfig;
        private JArray sourceFolders;

        private TreeNode selectedTreeViewNode = null;
        private int selectedListViewIndex = -1;

        private Vault myVault;

        private List<string> VaultedList;

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            myVault = new Vault();
            myVault.FileVaulted += MyVault_FileVaulted;

            if (!string.IsNullOrEmpty(Properties.Settings.Default.SourcePathArray))
            {
                sourceFolders = JArray.Parse(Properties.Settings.Default.SourcePathArray);
                foreach (string s in sourceFolders)
                {
                    comboBoxFolder.Items.Add(s);
                }
                for (int i = 0; i < comboBoxFolder.Items.Count; i++)
                {
                    if (comboBoxFolder.Items[i].ToString().Equals(Properties.Settings.Default.LastSourcePath,
                                                                  StringComparison.OrdinalIgnoreCase))
                    {
                        comboBoxFolder.SelectedIndex = i;
                        break;
                    }
                }
            }
            else
            {
                sourceFolders = new JArray();
            }
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            splitContainerMain.SplitterDistance = Properties.Settings.Default.SplitterPos;
            _loading = false;
        }

        private void MyVault_FileVaulted(string filename)
        {
            if (VaultedList == null)
            {
                VaultedList = new List<string>();
            }
            VaultedList.Add(filename);
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
            else
            {
                Properties.Settings.Default.LastSourcePath = enteredDir;
                Properties.Settings.Default.Save();
            }
            treeViewMainClear();
            listViewMainClear();
            listBoxVVClear();
            Application.DoEvents();
            if (File.Exists($"{enteredDir}\\{configFileName}"))
            {
                vvConfig = JObject.Parse(File.ReadAllText($"{enteredDir}\\{configFileName}"));
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
            JArray ignoreDir = (JArray)vvConfig.GetValue("IgnoreDir");
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
            JArray ignoreExt = (JArray)vvConfig.GetValue("IgnoreExt");
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
            string vvDir = $"{(string)vvConfig.GetValue("VVPath")}\\{treeViewMain.SelectedNode.FullPath}";
            if (Directory.Exists(vvDir))
            {
                foreach (string dirName in Directory.GetDirectories(vvDir))
                {
                    string dirNameBase = PathBase(dirName);
                    if (File.Exists($"{sourceDir}\\{dirNameBase}") || Directory.Exists($"{sourceDir}\\{dirNameBase}"))
                    {
                        continue;
                    }
                    ListViewItem item = new ListViewItem(PathBase(dirNameBase));
                    item.SubItems.Add("(deleted)");
                    item.SubItems.Add("");
                    listViewMain.Items.Add(item);
                }
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
            if (vvConfig != null && vvConfig.Contains("VVPath"))
            {
                string vvDir = $"{(string)vvConfig.GetValue("VVPath")}\\{treeViewMain.SelectedNode.FullPath}\\{files[0].Text}";
                if (Directory.Exists(vvDir))
                {
                    foreach (string filename in Directory.GetFiles(vvDir))
                    {
                        string baseFilename = filename.Substring(filename.LastIndexOf("\\") + 1);
                        if (baseFilename.Length == 32 && !baseFilename.Contains("_"))
                        {
                            continue;
                        }
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
            string vvDir = $"{(string)vvConfig.GetValue("VVPath")}\\{treeViewMain.SelectedNode.FullPath}\\{files[0].Text}";
            string vvFile = $"{vvDir}\\{((VVItem)listBoxVV.SelectedItem).ItemName}";
            if (File.Exists(sourceFile))
            {
                // check if files are exactly the same
                FileInfo sourceInfo = new FileInfo(sourceFile);
                FileInfo vvInfo = new FileInfo(vvFile);
                if (sourceInfo.Length == vvInfo.Length
                    && MD5Utilities.CalcMD5(sourceFile) == MD5Utilities.CalcMD5(vvFile))
                {
                    MessageBox.Show("Files are identical", "File Compare");
                }
                else
                {
                    // launch ExternalCompareApp process with the two files
                    string compareApp = Properties.Settings.Default.ExternalCompareApp;
                    string compareOptions = Properties.Settings.Default.ExternalCompareAppOptions;
                    Process p = Process.Start(compareApp, $"{compareOptions} \"{sourceFile}\" \"{vvFile}\"");
                    p.WaitForExit();
                    int result = p.ExitCode;
                }
            }
            else
            {
                // launch file viewer (default = notepad) to show the deleted file
                string compareApp = Properties.Settings.Default.ExternalFileViewer;
                Process p = Process.Start(compareApp, $"\"{vvFile}\"");
                p.WaitForExit();
                int result = p.ExitCode;
            }
        }

        private void externalCompareAppToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormExternalCompareApp popup = new FormExternalCompareApp();
            popup.PathToEXE = Properties.Settings.Default.ExternalCompareApp;
            popup.Options = Properties.Settings.Default.ExternalCompareAppOptions;
            popup.FileViewer = Properties.Settings.Default.ExternalFileViewer;
            popup.ShowDialog();
            if (popup.DialogResult != DialogResult.OK)
            {
                return;
            }
            Properties.Settings.Default.ExternalCompareApp = popup.PathToEXE;
            Properties.Settings.Default.ExternalCompareAppOptions = popup.Options;
            Properties.Settings.Default.ExternalFileViewer = popup.FileViewer;
            Properties.Settings.Default.Save();
        }

        private void toolStripMenuItemBackup_Click(object sender, EventArgs e)
        {
            if (VaultedList == null)
            {
                VaultedList = new List<string>();
            }
            else
            {
                VaultedList.Clear();
            }
            myVault.SourcePath = comboBoxFolder.Text;
            myVault.VaultPath = (string)vvConfig.GetValue("VVPath");
            myVault.BackupAll();
            string plural = (VaultedList.Count == 1) ? "" : "s";
            MessageBox.Show($"{VaultedList.Count} file{plural} stored", "Backup complete");
        }

        private void splitContainerMain_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (!_loading)
            {
                Properties.Settings.Default.SplitterPos = ((SplitContainer)sender).SplitterDistance;
                Properties.Settings.Default.Save();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileInfo fileInfo = new FileInfo(Application.ExecutablePath);
            string version = fileInfo.LastWriteTime.ToString("yyyy.MM.dd.HHmm");
            MessageBox.Show($"{Environment.CurrentDirectory}\r\n\r\nVersion {version}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK);
        }
    }
}
