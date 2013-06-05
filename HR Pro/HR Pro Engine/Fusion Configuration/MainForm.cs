using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win.UltraWinGrid.ExcelExport;
using Infragistics.Win.UltraWinToolbars;
using NHibernate;
using NHibernate.Linq;
using Resources = Fusion.Properties.Resources;
using Fusion.Properties;

namespace Fusion
{
	public partial class MainForm : Form
	{
		public static EventHandler LogsPurged;
		public static EventHandler EndEdits;

		private ISession _session;
		private ISessionFactory _sessionFactory;

		public MainForm()
		{
			InitializeComponent();
			Text = Text.Replace("TITLE", Resources.App_Title) + " - " + Settings.Default.Login_Database;

			//TODO app icon
			//TODO message element filter

			//TODO discuss: locking works the same, although if locked dont display who and dont allow read only access,
			//on save just give message about other users dont give list and allow messaging/auto retry etc...
			//TODO Spec updates for QA
			//TODO Setup / scripting
		}

		private void FrmMainLoad(object sender, EventArgs e)
		{
			_sessionFactory = Data.BuildSessionFactory(App.Database.ConnectionString);
			_session = _sessionFactory.OpenSession();

			viewPanel.Controls.Cast<Control>().ForEach(c => c.Dock = DockStyle.Fill);

			pageConfiguration.Display(_session);
			ShowView(pageConfiguration);
		}

		private void ShowView(Control view)
		{
			viewPanel.Controls.Cast<Control>().ForEach(c => c.TabStop = (c == view));
			view.BringToFront();
		}

		#region Commands

		private void ToolbarsManagerToolClick(object sender, ToolClickEventArgs e)
		{
			if (EndEdits != null)
				EndEdits(this, EventArgs.Empty);

			switch (e.Tool.Key) {
				case "ID_Save":
					Save();
					break;
				case "ID_ShowConfiguration":
					pageConfiguration.Display(_session);
					ShowView(pageConfiguration);
					break;
				case "ID_ShowMessages":
					pageMessages.Display(_session);
					ShowView(pageMessages);
					break;
				case "ID_ShowLogs":
					pageLogs.Display(_sessionFactory);
					ShowView(pageLogs);
					break;
				case "ID_LogExport":
					ExportLogs();
					break;
				case "ID_Purge":
					PurgeLogs();
					break;
				case "ID_StartConnector":
					pageSettings.Start();
					break;
				case "ID_StopConnector":
					pageSettings.Stop();
					break;
			}
		}

		private void ToolbarsManagerBeforeRibbonTabSelected(object sender, BeforeRibbonTabSelectedEventArgs e)
		{
			if (e.Tab.Key == "SystemSettings") {
				pageSettings.Display(_session);
				pageSettings.ServiceStatusChanged += (s, ev) => UpdateServiceStatus();
				ShowView(pageSettings);
				UpdateServiceStatus();
			}
			else if (e.Tab.Key == "Config") {
				pageSettings.SendToBack();
			}
		}

		private void UpdateServiceStatus()
		{
			toolbarsManager.Tools["ID_StartConnector"].SharedProps.Enabled = pageSettings.CanStart();
			toolbarsManager.Tools["ID_StopConnector"].SharedProps.Enabled = pageSettings.CanStop();
		}
		
		#endregion

		public void PurgeLogs()
		{
			if (MessageBox.Show(Resources.MessageBox_PurgeLogs, Resources.MessageBox_Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
				return;

			using(new WaitCursor()) {
				using (var tx = _session.BeginTransaction()) {
					_session.CreateQuery("delete from FusionLog").ExecuteUpdate();
					tx.Commit();
				}
				if (LogsPurged != null) {
					LogsPurged(null, EventArgs.Empty);
				}				
			}
		}

		public void ExportLogs()
		{
			var saveFile = new SaveFileDialog { Filter = Resources.Excel_Filter, FilterIndex = 1, DefaultExt = ".xls", FileName = "FusionLogs" };

			if (saveFile.ShowDialog() != DialogResult.OK)
				return;

			using(new WaitCursor()) {

				var exporter = new UltraGridExcelExporter();
				var grid = new UltraGrid();
				this.Controls.Add(grid);

				grid.DataSource = _session.Query<FusionLog>().Select(x => new { x.LastGeneratedDate, x.MessageType, x.BusRef, x.LastGeneratedXml }).ToList();
				try {
					exporter.Export(grid, saveFile.FileName);
				}
				catch (IOException) {
					MessageBox.Show(Resources.MessageBox_FileLocked, Resources.MessageBox_Title, MessageBoxButtons.OK,MessageBoxIcon.Information);
				}
				Controls.Remove(grid);
				grid.Dispose();
				exporter.Dispose();
			}

		}

		public bool IsDirty()
		{
			return pageSettings.IsDirty || _session.IsDirty();
		}

		public bool Save()
		{
			if (!App.Database.Lock(LockLevel.Saving)) {
				MessageBox.Show(Resources.Database_Lock_CantTakeSaving, Resources.MessageBox_Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}

			if (_session.IsDirty()) {
				using (var tx = _session.BeginTransaction()) {
					_session.Flush();
					tx.Commit();
				}
			}

			if (pageSettings.IsDirty) {
				pageSettings.Save();
			}

			App.Database.Unlock(LockLevel.Saving);
			return true;
		}

		private void MainFormFormClosing(object sender, FormClosingEventArgs e)
		{
			if (e.CloseReason != CloseReason.UserClosing)
				return;

			if (EndEdits != null)
				EndEdits(this, EventArgs.Empty);

			if (IsDirty()) {
				var result = MessageBox.Show(Resources.MessageBox_SaveChanges, Resources.MessageBox_Title, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				switch (result) {
					case DialogResult.Yes:
						if (!Save())
							e.Cancel = true;
						break;
					case DialogResult.No:
						break;
					case DialogResult.Cancel:
						e.Cancel = true;
						break;
				}
			}
		}

		private void MainFormFormClosed(object sender, FormClosedEventArgs e)
		{
			App.Database.Unlock(LockLevel.ReadWrite);
			App.Database.Dispose();
		}
	}
}