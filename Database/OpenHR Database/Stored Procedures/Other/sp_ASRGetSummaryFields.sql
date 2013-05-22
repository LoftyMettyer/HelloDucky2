CREATE PROCEDURE [dbo].[sp_ASRGetSummaryFields] (
	@piHistoryTableID	integer,
	@piParentTableID 	integer)
AS
BEGIN
	SELECT DISTINCT ASRSysSummaryFields.sequence, 
	    ASRSysSummaryFields.startOfGroup, 
		ASRSysColumns.columnName, 
		ASRSysColumns.columnID, 
		ASRSysColumns.tableID, 
		ASRSysColumns.dataType, 
		ASRSysColumns.size, 
		ASRSysColumns.decimals, 
		ASRSysColumns.controlType, 
		ASRSysColumns.columnType, 
		ASRSysColumns.multiline,
		ASRSysColumns.alignment,
		ASRSysColumns.BlankIfZero,
		ASRSysColumns.Use1000Separator,		
	    ASRSysSummaryFields.StartOfColumn
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence;
END