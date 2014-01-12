Imports HR.Intranet.Server.Enums

Namespace Interfaces
	Public Interface IChart
		Function GetChartData(plngTableID As Integer, plngColumnID As Integer, plngFilterID As Integer,
																piAggregateType As Integer, piElementType As ElementType,
																plngTableID_2 As Integer, plngColumnID_2 As Integer, plngTableID_3 As Integer, plngColumnID_3 As Integer,
																plngSortOrderID As Integer, piSortDirection As Integer, plngChart_ColourID As Integer) As DataTable
		Property SessionInfo() As SessionInfo
	End Interface
End Namespace