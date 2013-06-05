using System;
using System.Linq;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using NHibernate;
using NHibernate.Linq;

namespace Fusion.Pages
{
	public partial class Logs : UserControl
	{
		private IStatelessSession _session;

		public Logs()
		{
			InitializeComponent();
			LogGrid.InitializeLayout += (s, e) => (s as UltraGrid).ApplyDefaults();
		}

		public void Display(ISessionFactory sessionFactory)
		{
			if (_session == null) {
				using (new WaitCursor()) {
					_session = sessionFactory.OpenStatelessSession();

					messageTypeEditor.Items.Add(new ValueListItem(null, ""));
					messageTypeEditor.Items.AddRange(_session.Query<FusionMessage>().OrderBy(x => x.Name).Select(x => x.Name).Select(x => new ValueListItem(x)).ToArray());

					queryBindingSource.DataSource = new FusionLogQuery {DateLastGeneratedFrom = DateTime.Now.AddMonths(-1)};

					Find();
					FindButton.Click += (s, e) => Find();
					MainForm.LogsPurged += (s, e) => Find();
				}
			}
			messageTypeEditor.Select();
		}

		private void Find()
		{
			using (new WaitCursor()) {
				logBindingSource.DataSource = ((FusionLogQuery) queryBindingSource.DataSource).GetQuery(_session).ToList();
			}
		}
	}

	public class FusionLogQuery
	{
		public string MessageType { get; set; }
		public DateTime? DateLastGeneratedFrom { get; set; }
		public DateTime? DateLastGeneratedTo { get; set; }

		public IQueryable<FusionLog> GetQuery(IStatelessSession session)
		{
			var q = session.Query<FusionLog>();

			if (!string.IsNullOrWhiteSpace(MessageType)) {
				q = q.Where(x => x.MessageType == MessageType);
			}
			if (DateLastGeneratedFrom.HasValue) {
				q = q.Where(x => x.LastGeneratedDate >= DateLastGeneratedFrom.Value);
			}
			if (DateLastGeneratedTo.HasValue) {
				q = q.Where(x => x.LastGeneratedDate < DateLastGeneratedTo.Value.AddDays(1));
			}

			return q;
		}
	}
}