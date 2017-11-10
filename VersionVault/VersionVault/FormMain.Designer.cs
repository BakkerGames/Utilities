namespace VersionVault
{
    partial class FormMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.menuStripMain = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemBackup = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemOptions = new System.Windows.Forms.ToolStripMenuItem();
            this.externalCompareAppToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.folderBrowserDialogMain = new System.Windows.Forms.FolderBrowserDialog();
            this.buttonFolder = new System.Windows.Forms.Button();
            this.comboBoxFolder = new System.Windows.Forms.ComboBox();
            this.statusStripMain = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelMain = new System.Windows.Forms.ToolStripStatusLabel();
            this.listBoxVV = new System.Windows.Forms.ListBox();
            this.buttonGo = new System.Windows.Forms.Button();
            this.panelMain = new System.Windows.Forms.Panel();
            this.splitContainerMain = new System.Windows.Forms.SplitContainer();
            this.treeViewMain = new System.Windows.Forms.TreeView();
            this.listViewMain = new System.Windows.Forms.ListView();
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderDateModified = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderSize = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.menuStripMain.SuspendLayout();
            this.statusStripMain.SuspendLayout();
            this.panelMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).BeginInit();
            this.splitContainerMain.Panel1.SuspendLayout();
            this.splitContainerMain.Panel2.SuspendLayout();
            this.splitContainerMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStripMain
            // 
            this.menuStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolStripMenuItemBackup,
            this.toolStripMenuItemOptions,
            this.helpToolStripMenuItem});
            this.menuStripMain.Location = new System.Drawing.Point(0, 0);
            this.menuStripMain.Name = "menuStripMain";
            this.menuStripMain.Size = new System.Drawing.Size(789, 24);
            this.menuStripMain.TabIndex = 0;
            this.menuStripMain.Text = "menuStripMain";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            // 
            // toolStripMenuItemBackup
            // 
            this.toolStripMenuItemBackup.Name = "toolStripMenuItemBackup";
            this.toolStripMenuItemBackup.Size = new System.Drawing.Size(58, 20);
            this.toolStripMenuItemBackup.Text = "&Backup";
            this.toolStripMenuItemBackup.Click += new System.EventHandler(this.toolStripMenuItemBackup_Click);
            // 
            // toolStripMenuItemOptions
            // 
            this.toolStripMenuItemOptions.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.externalCompareAppToolStripMenuItem});
            this.toolStripMenuItemOptions.Name = "toolStripMenuItemOptions";
            this.toolStripMenuItemOptions.Size = new System.Drawing.Size(61, 20);
            this.toolStripMenuItemOptions.Text = "&Options";
            // 
            // externalCompareAppToolStripMenuItem
            // 
            this.externalCompareAppToolStripMenuItem.Name = "externalCompareAppToolStripMenuItem";
            this.externalCompareAppToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.externalCompareAppToolStripMenuItem.Text = "&External Compare App...";
            this.externalCompareAppToolStripMenuItem.Click += new System.EventHandler(this.externalCompareAppToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // buttonFolder
            // 
            this.buttonFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonFolder.Location = new System.Drawing.Point(706, 25);
            this.buttonFolder.Name = "buttonFolder";
            this.buttonFolder.Size = new System.Drawing.Size(71, 23);
            this.buttonFolder.TabIndex = 3;
            this.buttonFolder.Text = "Search";
            this.buttonFolder.UseVisualStyleBackColor = true;
            // 
            // comboBoxFolder
            // 
            this.comboBoxFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxFolder.FormattingEnabled = true;
            this.comboBoxFolder.Location = new System.Drawing.Point(12, 27);
            this.comboBoxFolder.Name = "comboBoxFolder";
            this.comboBoxFolder.Size = new System.Drawing.Size(624, 21);
            this.comboBoxFolder.Sorted = true;
            this.comboBoxFolder.TabIndex = 1;
            // 
            // statusStripMain
            // 
            this.statusStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelMain});
            this.statusStripMain.Location = new System.Drawing.Point(0, 365);
            this.statusStripMain.Name = "statusStripMain";
            this.statusStripMain.Size = new System.Drawing.Size(789, 22);
            this.statusStripMain.TabIndex = 7;
            this.statusStripMain.Text = "statusStrip1";
            // 
            // toolStripStatusLabelMain
            // 
            this.toolStripStatusLabelMain.Name = "toolStripStatusLabelMain";
            this.toolStripStatusLabelMain.Size = new System.Drawing.Size(774, 17);
            this.toolStripStatusLabelMain.Spring = true;
            this.toolStripStatusLabelMain.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // listBoxVV
            // 
            this.listBoxVV.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxVV.FormattingEnabled = true;
            this.listBoxVV.IntegralHeight = false;
            this.listBoxVV.Location = new System.Drawing.Point(642, 54);
            this.listBoxVV.Name = "listBoxVV";
            this.listBoxVV.ScrollAlwaysVisible = true;
            this.listBoxVV.Size = new System.Drawing.Size(135, 308);
            this.listBoxVV.Sorted = true;
            this.listBoxVV.TabIndex = 6;
            this.listBoxVV.DoubleClick += new System.EventHandler(this.listBoxVV_DoubleClick);
            // 
            // buttonGo
            // 
            this.buttonGo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonGo.Location = new System.Drawing.Point(642, 25);
            this.buttonGo.Name = "buttonGo";
            this.buttonGo.Size = new System.Drawing.Size(58, 23);
            this.buttonGo.TabIndex = 2;
            this.buttonGo.Text = "Go";
            this.buttonGo.UseVisualStyleBackColor = true;
            this.buttonGo.Click += new System.EventHandler(this.buttonGo_Click);
            // 
            // panelMain
            // 
            this.panelMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelMain.Controls.Add(this.splitContainerMain);
            this.panelMain.Location = new System.Drawing.Point(12, 54);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(624, 308);
            this.panelMain.TabIndex = 8;
            // 
            // splitContainerMain
            // 
            this.splitContainerMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerMain.Location = new System.Drawing.Point(0, 0);
            this.splitContainerMain.Name = "splitContainerMain";
            // 
            // splitContainerMain.Panel1
            // 
            this.splitContainerMain.Panel1.Controls.Add(this.treeViewMain);
            // 
            // splitContainerMain.Panel2
            // 
            this.splitContainerMain.Panel2.Controls.Add(this.listViewMain);
            this.splitContainerMain.Size = new System.Drawing.Size(624, 308);
            this.splitContainerMain.SplitterDistance = 200;
            this.splitContainerMain.TabIndex = 0;
            this.splitContainerMain.SplitterMoved += new System.Windows.Forms.SplitterEventHandler(this.splitContainerMain_SplitterMoved);
            // 
            // treeViewMain
            // 
            this.treeViewMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewMain.Location = new System.Drawing.Point(0, 0);
            this.treeViewMain.Name = "treeViewMain";
            this.treeViewMain.Size = new System.Drawing.Size(200, 308);
            this.treeViewMain.TabIndex = 5;
            this.treeViewMain.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewMain_AfterSelect);
            // 
            // listViewMain
            // 
            this.listViewMain.BackColor = System.Drawing.SystemColors.Window;
            this.listViewMain.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderName,
            this.columnHeaderDateModified,
            this.columnHeaderSize});
            this.listViewMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewMain.FullRowSelect = true;
            this.listViewMain.GridLines = true;
            this.listViewMain.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewMain.Location = new System.Drawing.Point(0, 0);
            this.listViewMain.MultiSelect = false;
            this.listViewMain.Name = "listViewMain";
            this.listViewMain.ShowGroups = false;
            this.listViewMain.Size = new System.Drawing.Size(420, 308);
            this.listViewMain.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listViewMain.TabIndex = 6;
            this.listViewMain.UseCompatibleStateImageBehavior = false;
            this.listViewMain.View = System.Windows.Forms.View.Details;
            this.listViewMain.SelectedIndexChanged += new System.EventHandler(this.listViewMain_SelectedIndexChanged);
            this.listViewMain.Leave += new System.EventHandler(this.listViewMain_Leave);
            this.listViewMain.Resize += new System.EventHandler(this.listViewMain_Resize);
            // 
            // columnHeaderName
            // 
            this.columnHeaderName.Text = "Name";
            this.columnHeaderName.Width = 100;
            // 
            // columnHeaderDateModified
            // 
            this.columnHeaderDateModified.Text = "Date Modified UTC";
            this.columnHeaderDateModified.Width = 130;
            // 
            // columnHeaderSize
            // 
            this.columnHeaderSize.Text = "Size";
            this.columnHeaderSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.columnHeaderSize.Width = 80;
            // 
            // FormMain
            // 
            this.AcceptButton = this.buttonGo;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(789, 387);
            this.Controls.Add(this.panelMain);
            this.Controls.Add(this.buttonGo);
            this.Controls.Add(this.listBoxVV);
            this.Controls.Add(this.statusStripMain);
            this.Controls.Add(this.comboBoxFolder);
            this.Controls.Add(this.buttonFolder);
            this.Controls.Add(this.menuStripMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStripMain;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Version Vault";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.Shown += new System.EventHandler(this.FormMain_Shown);
            this.menuStripMain.ResumeLayout(false);
            this.menuStripMain.PerformLayout();
            this.statusStripMain.ResumeLayout(false);
            this.statusStripMain.PerformLayout();
            this.panelMain.ResumeLayout(false);
            this.splitContainerMain.Panel1.ResumeLayout(false);
            this.splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerMain)).EndInit();
            this.splitContainerMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStripMain;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogMain;
        private System.Windows.Forms.Button buttonFolder;
        private System.Windows.Forms.ComboBox comboBoxFolder;
        private System.Windows.Forms.StatusStrip statusStripMain;
        private System.Windows.Forms.ListBox listBoxVV;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelMain;
        private System.Windows.Forms.Button buttonGo;
        private System.Windows.Forms.Panel panelMain;
        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.TreeView treeViewMain;
        private System.Windows.Forms.ListView listViewMain;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.ColumnHeader columnHeaderDateModified;
        private System.Windows.Forms.ColumnHeader columnHeaderSize;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemOptions;
        private System.Windows.Forms.ToolStripMenuItem externalCompareAppToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemBackup;
    }
}

