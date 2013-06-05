using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Fusion
{
	public static class Grid
	{
		public static void ApplyDefaults(this UltraGrid grid)
		{
			var layout = grid.DisplayLayout;
			layout.TabNavigation = TabNavigation.NextControl;
			layout.GroupByBox.Hidden = true;
			layout.ViewStyle = ViewStyle.SingleBand;
			layout.AutoFitStyle = AutoFitStyle.ExtendLastColumn;
			layout.Override.CellClickAction = CellClickAction.RowSelect;
			layout.Override.BorderStyleCell = UIElementBorderStyle.None;
			layout.Override.BorderStyleRow = UIElementBorderStyle.None;
		}
	}
}
