CREATE PROCEDURE [dbo].[spASRIntGetCalendarReportColumns]
	(
	@piBaseTableID 		integer,
	@piEventTableID		integer
	)
AS
BEGIN

	/* Return a recordset of the columns for the given table IDs.*/
	SELECT ASRSysColumns.ColumnID, ASRSysColumns.TableID, ASRSysColumns.ColumnName, ASRSysColumns.DataType, 
           ASRSysColumns.ColumnType, ASRSysColumns.Size, ASRSystables.TableName,
           ASRSysColumns.LookupTableID, ASRSysColumns.LookupColumnID
    FROM ASRSysColumns 
		INNER JOIN ASRSystables 
		ON ASRSysColumns.TableID = ASRSystables.TableID
    WHERE (ASRSysColumns.TableID = @piBaseTableID
			OR ASRSysColumns.TableID = @piEventTableID)
		AND ASRSysColumns.columnType <> 3
		AND ASRSysColumns.columnType <> 4
		AND ASRSysColumns.dataType <> -3
		AND ASRSysColumns.dataType <> -4
	ORDER BY ASRSystables.TableName, ASRSysColumns.ColumnName;

END