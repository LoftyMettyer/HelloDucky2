using System;
using System.ServiceProcess;
using System.Windows.Forms;
using Fusion.Properties;
using NHibernate;
using TimeoutException = System.TimeoutException;

namespace Fusion.Pages
{
	public partial class Settings : UserControl
	{
		public event EventHandler ServiceStatusChanged;
		public bool IsDirty { get; protected set; }

		private ISession _session;
		private ServiceController _service;
		private TimeSpan _timeout;

		public Settings()
		{
			InitializeComponent();
		}

		public void Display(ISession session)
		{
			if(_session == null) {
				_session = session;

				communityDatabase.Text = _session.CreateSQLQuery("select SettingValue from ASRSysSystemSettings where Section='fusion' and SettingKey = 'community database'").UniqueResult<string>();
				communityDatabase.ValueChanged += (s, e) => IsDirty = true;
			}

			if(_service == null) {
				_service = new ServiceController(Properties.Settings.Default.Service_Name.Replace("{db_name}", Properties.Settings.Default.Login_Database));
				_timeout = TimeSpan.FromMilliseconds(Properties.Settings.Default.Service_Timeout);

				try {
					//we can tell if the service is installed by trying to read its status
					var status = _service.Status;
				}
				catch (InvalidOperationException) {
					_service = null;
				}
				serviceStatusLabel.Visible = (_service == null);
			}

			communityDatabase.Select();
		}

		public void Save()
		{
			//plain sql instead of NH cos table doesn't have a PK, yuk!
			var sql = string.Format(
				"update ASRSysSystemSettings set SettingValue = '{2}' where Section = '{0}' and SettingKey = '{1}' \n" +
				"if(@@ROWCOUNT = 0) \n" +
				"insert into ASRSysSystemSettings (Section, SettingKey, SettingValue) values ('{0}', '{1}', '{2}')",
				"fusion", "community database", communityDatabase.Text.Trim());

			_session.CreateSQLQuery(sql).ExecuteUpdate();

			IsDirty = false;
		}

		public void Start()
		{
			PerformServiceAction(s => s.Start(), ServiceControllerStatus.Running);
		}

		public void Stop()
		{
			PerformServiceAction(s => s.Stop(), ServiceControllerStatus.Stopped);
		}

		public bool CanStart()
		{
			return _service != null && (_service.Status == ServiceControllerStatus.Stopped || _service.Status == ServiceControllerStatus.Paused);
		}

		public bool CanStop()
		{
			return _service != null && (_service.Status == ServiceControllerStatus.Running || _service.Status == ServiceControllerStatus.Paused);
		}

		private void PerformServiceAction(Action<ServiceController> action, ServiceControllerStatus waitForStatus)
		{
			try {
				using (new WaitCursor()) {
					action(_service);
					_service.WaitForStatus(waitForStatus, _timeout);
				}
			}
			catch (TimeoutException) {
				MessageBox.Show(Resources.Service_Timeout, Resources.MessageBox_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			if (ServiceStatusChanged != null)
				ServiceStatusChanged(this, EventArgs.Empty);
		}
	}
}
