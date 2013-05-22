CREATE PROCEDURE [dbo].[spASRIntGetSortOrderColumns] (
	@psIncludedColumns		varchar(MAX),
	@psExcludedColumns		varchar(MAX)
)
AS
BEGIN
	DECLARE @sSQL nvarchar(MAX);
	
	/* Clean the input string parameters. */
	IF len(@psIncludedColumns) > 0 SET @psIncludedColumns = replace(@psIncludedColumns, '''', '''''');
	IF len(@psExcludedColumns) > 0 SET @psExcludedColumns = replace(@psExcludedColumns, '''', '''''');

	SET @sSQL = 'SELECT ASRSysColumns.columnID, ' +
		'ASRSysTables.tableName + ''.'' + ASRSysColumns.columnName AS [columnName] ' +
		'FROM ASRSysColumns ' +
		'INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID ' +
		'WHERE ASRSysColumns.columnID IN ('+ @psIncludedColumns + ')';

	IF len(@psExcludedColumns) > 0
	BEGIN
		SET @sSQL = @sSQL + ' AND [columnID] NOT IN (' + @psExcludedColumns + ')';
	END

	SET @sSQL = @sSQL + ' ORDER BY [columnName] ASC';
	
	EXECUTE sp_executeSQL @sSQL;
END