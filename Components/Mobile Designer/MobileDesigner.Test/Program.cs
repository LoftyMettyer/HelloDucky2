using System;
using System.Windows.Forms;

namespace MobileDesigner.Test
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            MobileDesignerSerivce.Initialise(Properties.Settings.Default.db);
            var f = new DesignerForm();
            //f.ReadOnly = true;
            Application.Run(f);
        }
    }
}
