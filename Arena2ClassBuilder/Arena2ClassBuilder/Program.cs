// Program.cs - 11/22/2017

using Arena.Common.Bootstrap;
using System;
using System.Windows.Forms;

namespace Arena2ClassBuilder
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // check if bootstrapping
            try
            {
                if (Bootstrapper.MustBootstrap())
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK);
                return;
            }
            Application.Run(new FormMain());
        }
    }
}
