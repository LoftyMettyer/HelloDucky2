using System;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Fusion.Properties;
using Infragistics.Win.UltraWinEditors;
using NHibernate.Linq;
using Resources = Fusion.Properties.Resources;

namespace Fusion
{
	public partial class LoginForm : Form
	{
		public LoginForm()
		{
			InitializeComponent();
			Text = Text.Replace("TITLE", Resources.App_Title);
			detailsButton.Click += (s, e) => ShowDetails(!databaseLabel.Visible);
		}

		private void LoginFormLoad(object sender, EventArgs e)
		{
			var version = Assembly.GetExecutingAssembly().GetName().Version;
			versionLabel.Text = string.Format("Version {0}.{1}.{2}", version.Major, version.MajorRevision, version.Minor);

			usernameEditor.Text = Settings.Default.Login_User;
			useIntegratedEditor.Checked = Settings.Default.Login_Integrated;
			databaseEditor.Text = Settings.Default.Login_Database;
			serverEditor.Text = Settings.Default.Login_Server;
			ShowDetails(Settings.Default.Login_ShowDetails);

			if(usernameEditor.Text != "")
				passwordEditor.Select();
		}

		private bool IsValid()
		{
			return (databaseEditor.Text != "" && serverEditor.Text != "") &&
			       (useIntegratedEditor.Checked ||
			        (!useIntegratedEditor.Checked && usernameEditor.Text != "" && passwordEditor.Text != ""));
		}

		private void LoginFormFormClosing(object sender, FormClosingEventArgs e)
		{
			if (DialogResult != DialogResult.OK) {
				return;
			}

			Controls.OfType<UltraTextEditor>().ForEach(c => c.Text = c.Text.Trim());

			if (!IsValid()) {
				MessageBox.Show(Resources.LoginForm_ValidationFailed, Resources.MessageBox_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				e.Cancel = true;
				return;
			}

			var builder = new SqlConnectionStringBuilder {
			                                             	ApplicationName = "OpenHR System Manager",
			                                             	IntegratedSecurity = useIntegratedEditor.Checked,
			                                             	DataSource = serverEditor.Text,
			                                             	InitialCatalog = databaseEditor.Text
			                                             };

			if (!useIntegratedEditor.Checked) {
				builder.UserID = usernameEditor.Text;
				builder.Password = passwordEditor.Text;
			}

			var connectionString = builder.ToString();

			string error = null;

			App.Database = new Database(connectionString);

			if (!App.Database.IsValid()) {
				error = Resources.Database_CantConnect;
			}
			else if (!App.Database.IsAdmin()) {
				error = Resources.Database_NotAdmin;
			}
			else if (!App.Database.Lock(LockLevel.ReadWrite)) {
				error = Resources.Database_Lock_CantTakeReadWrite;
			}

			if (error != null) {
				MessageBox.Show(error, Resources.MessageBox_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				e.Cancel = true;
				return;
			}

			Settings.Default.Login_User = useIntegratedEditor.Checked ? "" : usernameEditor.Text;
			Settings.Default.Login_Integrated = useIntegratedEditor.Checked;
			Settings.Default.Login_Database = databaseEditor.Text;
			Settings.Default.Login_Server = serverEditor.Text;
			Settings.Default.Login_ShowDetails = databaseLabel.Visible;
			Settings.Default.Save();
		}

		private void UseIntegratedEditorCheckedChanged(object sender, EventArgs e)
		{
			if (useIntegratedEditor.Checked) {
				usernameEditor.Text = string.Format("{0}\\{1}", Environment.UserDomainName, Environment.UserName);
			}
			usernameEditor.Enabled = !useIntegratedEditor.Checked;
			passwordEditor.Enabled = !useIntegratedEditor.Checked;
		}

		private void ShowDetails(bool show)
		{
			databaseLabel.Visible = show;
			databaseEditor.Visible = show;
			serverLabel.Visible = show;
			serverEditor.Visible = show;
			detailsButton.Text = (show ? "Details <<" : "Details >>");
			Height = show ? 283 : 256;
		}
	}
}