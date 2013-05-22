CREATE PROCEDURE [dbo].[sp_ASRIntGetExprLookupValues]
(	@piColumnID		integer,
	@piDataType		integer		OUTPUT)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of values for the given lookup column. */
	DECLARE @sColumnName	sysname,
			@sTableName		sysname,
			@sExecString	nvarchar(MAX);

	SELECT @sColumnName = ASRSysColumns.columnName, 
		@sTableName = ASRSysTables.tableName,
		@piDataType = ASRSysColumns.dataType
	FROM ASRSysColumns
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysColumns.columnID = @piColumnID

	IF @piDataType = 11
	BEGIN
		SET @sExecString = 'SELECT DISTINCT convert(varchar(10), ' + @sColumnName + ', 101) AS lookUpValue' +
			' FROM ' + @sTableName +
			' ORDER BY lookUpValue;';
	END
	ELSE
	BEGIN
		SET @sExecString = 'SELECT DISTINCT ' + @sColumnName + ' AS lookUpValue' +
			' FROM ' + @sTableName +
			' ORDER BY lookUpValue;';
	END
	
	-- Get the data
	EXECUTE sp_executeSQL @sExecString;
	
END