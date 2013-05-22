CREATE PROCEDURE [dbo].[sp_ASRIntGetLookupValues] (
	@piColumnID 	integer
)
AS
BEGIN

	/* Return a recordset of the lookup values for the given lookup column. */
	DECLARE	@sColumnName	sysname,
			@sTableName		sysname,
			@sExecString	nvarchar(MAX);

	SELECT @sTableName = ASRSysTables.tableName,
		@sColumnName = ASRSysColumns.columnName
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE columnID = @piColumnID;

	SET @sExecString = 'SELECT ' + @sColumnName + 
		' FROM ' + @sTableName +
		' ORDER BY ' + @sColumnName;

	/* Return a recordset of the required columns in the required order from the given table/view. */
	EXECUTE sp_executeSQL @sExecString;
	
END