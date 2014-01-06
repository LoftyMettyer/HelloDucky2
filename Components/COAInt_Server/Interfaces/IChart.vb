Imports HR.Intranet.Server.Enums

Namespace Interfaces
	Public Interface IChart
		Function GetChartData(ByRef plngTableID As Long, ByRef plngColumnID As Long, ByRef plngFilterID As Long,
															 ByRef piAggregateType As Long, ByRef piElementType As ElementType,
															 ByRef plngTableID_2 As Long, ByRef plngColumnID_2 As Long, ByRef plngTableID_3 As Long, ByRef plngColumnID_3 As Long,
															 ByRef plngSortOrderID As Long, ByRef piSortDirection As Long, ByRef plngChart_ColourID As Long) As DataTable
	End Interface
End Namespace