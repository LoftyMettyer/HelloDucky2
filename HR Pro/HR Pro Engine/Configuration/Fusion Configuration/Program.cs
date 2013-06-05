using System;
using System.Windows.Forms;

namespace Fusion
{
	internal static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		private static void Main()
		{
			//HibernatingRhinos.Profiler.Appender.NHibernate.NHibernateProfiler.Initialize(); 

			Infragistics.Win.AppStyling.StyleManager.Load("Office2007Blue.isl");

			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			using (var f = new LoginForm()) {
				if (f.ShowDialog() == DialogResult.OK)
					using(new WaitCursor()) {
						Application.Run(new MainForm());
					}
			}
		}
	}
}