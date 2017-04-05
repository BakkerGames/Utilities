// FormExternalCompareApp.cs - 03/14/2017

using System;
using System.Windows.Forms;

namespace VersionVault
{
    public partial class FormExternalCompareApp : Form
    {

        private string _PathToEXE = "";
        public string PathToEXE
        {
            get
            {
                return _PathToEXE;
            }
            set
            {
                _PathToEXE = value;
                textBoxPath.Text = value;
            }
        }

        private string _Options = "";
        public string Options
        {
            get
            {
                return _Options;
            }
            set
            {
                _Options = value;
                textBoxOptions.Text = value;
            }
        }

        private string _FileViewer = "";
        public string FileViewer
        {
            get
            {
                return _FileViewer;
            }
            set
            {
                _FileViewer = value;
                textBoxFileViewer.Text = value;
            }
        }

        public FormExternalCompareApp()
        {
            InitializeComponent();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            PathToEXE = textBoxPath.Text;
            Options = textBoxOptions.Text;
            FileViewer = textBoxFileViewer.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
