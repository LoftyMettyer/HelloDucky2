using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using NHibernate;
using NHibernate.Linq;
using Nullable = Infragistics.Win.UltraWinGrid.Nullable;
using Resources = Fusion.Properties.Resources;

namespace Fusion.Pages
{
	public partial class Configuration : UserControl
	{
		private ISession _session;

		public Configuration()
		{
			InitializeComponent();
		}

		public void Display(ISession session)
		{
			if (_session == null) {
				_session = session;

				using (new WaitCursor()) {
					fusionCategoryBindingSource.DataSource = _session.Query<FusionCategory>().ToList();

					fusionElementBindingSource.DataSource = fusionCategoryBindingSource;
					fusionElementBindingSource.DataMember = "Elements";
					InitElementGrid();

					drpTables.DropDownStyle = DropDownStyle.DropDownList;
					drpTables.Items.Add(new ValueListItem(null, ""));
					drpTables.Items.AddRange(_session.Query<Table>().OrderBy(x => x.Name).Select(x => new ValueListItem(x, x.Name)).ToArray());
					drpTables.Enabled = (fusionCategoryBindingSource.Count > 0);

					//grids highlight the first row from datasource not the first row in the grid, fix it, for both grid
					if (categoryGrid.Rows.Count > 0) categoryGrid.Rows[0].Activate();

					bool fix = false;
					fusionElementBindingSource.ListChanged += (s, e) => { if (e.ListChangedType == ListChangedType.Reset) fix = true; };
					elementGrid.AfterRowActivate += (s, e) => { if (fix) { elementGrid.Rows[0].Activate(); fix = false; }};

					//if the app is closing make sure changes are pushing into the databinding sources
					MainForm.AppClosing += (s, e) => elementGrid.UpdateData();
				}
			}
			categoryGrid.Select();
		}

		private void categoryGridInitializeLayout(object sender, InitializeLayoutEventArgs e)
		{
			categoryGrid.ApplyDefaults();

			foreach (var col in e.Layout.Bands[0].Columns)
			{
				if (col.Key == "Name") {
					col.Header.Caption = "Categories";
					col.SortIndicator = SortIndicator.Ascending;
				}
				else
					col.Hidden = true;
			}	
		}

		private void InitElementGrid()
		{
			elementGrid.ApplyDefaults();

			if (fusionCategoryBindingSource.Count == 0)
				return;

			var e = elementGrid.DisplayLayout;
			e.Bands[0].Columns["Id"].Hidden = true;
			e.Bands[0].Columns["Category"].Hidden = true;
			e.Bands[0].Columns["DataType"].Hidden = true;
			e.Bands[0].Columns["Lookup"].Hidden = true;
			e.Bands[0].Columns["Column"].Nullable = Nullable.Nothing;
			e.Bands[0].Columns["Name"].Width = 150;
			e.Bands[0].Columns["Name"].SortIndicator = SortIndicator.Ascending;
			e.Bands[0].Columns["Name"].Header.VisiblePosition = 0;
			e.Bands[0].Columns["Description"].Width = 175;
			e.Bands[0].Columns["MinSize"].Header.Caption = "Min Size";
			e.Bands[0].Columns["MaxSize"].Header.Caption = "Max Size";

			var columnList = e.ValueLists.Add("columnList");
			e.Bands[0].Columns["Column"].ValueList = columnList;
			e.Bands[0].Columns["Column"].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;				
		}

		private IEnumerable<Column> GetAvailableColumnsForElement(FusionElement element)
		{
			var columns = _session.Query<Column>().Where(x => x.Table == element.Category.Table && x.DataType == element.DataType);

			if (element.DataType == DataType.Character)
				columns = columns.Where(x => x.Size >= element.MinSize && x.Size <= element.MaxSize);
			if (element.Lookup)
				columns = columns.Where(x => x.LookupTableId > 0);

			return columns.OrderBy(c => c.Name).ToList();
		}

		private void GrdElementsBeforeCellListDropDown(object sender, CancelableCellEventArgs e)
		{
			var availColumns = GetAvailableColumnsForElement((FusionElement)fusionElementBindingSource.Current);

			var columnList = elementGrid.DisplayLayout.ValueLists["columnList"];
			columnList.ValueListItems.Clear();
			columnList.ValueListItems.Add(new ValueListItem(null, ""));
			columnList.ValueListItems.AddRange(availColumns.Select(x => new ValueListItem(x, x.Name)).ToArray());
		}

		//if the table for the category changes then all the columns selected in the elements will be invalid, warn user then clear them
		private void OnTableChanged(object sender, EventArgs e)
		{
			var category = (FusionCategory)fusionCategoryBindingSource.Current;

			var resetColumns = category.Elements.Any(x => x.Column != null && x.Column.Table != category.Table);

			if (resetColumns) {
				var result = MessageBox.Show(Resources.MessageBox_ColumnMappingsLost, Resources.MessageBox_Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

				switch(result) {
					case DialogResult.OK:
						category.Elements.ToArray().ForEach(x => x.Column = null);
						elementGrid.Refresh();
						break;
					default:
						drpTables.Value = _previousTable;
						break;
				}
			}
		}

		//combo doesnt allow you to cancel a selection change without it updating the data source, fix it
		//store the previous value so we can manually set the value back if we are not happy with the change
		private void DrpTablesValueChanged(object sender, EventArgs e)
		{
			_previousTable = ((FusionCategory)fusionCategoryBindingSource.Current).Table;
		}

		private Table _previousTable;
	}
}