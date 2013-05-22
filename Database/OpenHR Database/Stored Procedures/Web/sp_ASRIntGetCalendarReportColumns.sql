CREATE PROCEDURE [dbo].[sp_ASRIntGetCalendarReportColumns] (
	@piBaseTableID 		integer
	)
AS
BEGIN
	
	/* Return a recordset of the columns for the given table IDs.*/
	DECLARE @sUserName sysname;

	SELECT @sUserName = SYSTEM_USER;

	SELECT 	ASRSysColumns.tableID,
			ASRSysColumns.columnID,
			ASRSysTables.tableName,
			ASRSysColumns.columnName
	FROM ASRSysColumns
			INNER JOIN ASRSysTables 
			ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysTables.tableID = @piBaseTableID
			AND ASRSysColumns.columnType NOT IN (3,4) 
			AND ASRSysColumns.dataType NOT IN (-3, -4)
	ORDER BY ASRSysColumns.columnName;

END