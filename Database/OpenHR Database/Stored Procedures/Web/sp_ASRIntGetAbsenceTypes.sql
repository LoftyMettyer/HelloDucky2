CREATE PROCEDURE [dbo].[sp_ASRIntGetAbsenceTypes]
AS
BEGIN
	DECLARE	@sExecString	nvarchar(MAX),
		@sTableName			varchar(255),
		@sColumnName		varchar(255),
		@iTableID			integer,
		@iColumnID			integer,
		@sParameterValue	varchar(MAX);

	SET @sTableName = '';
	SET @sColumnName = '';
	SET @sParameterValue = '';

	/* Get the Absence Type table name. */
	SELECT @sParameterValue = parameterValue
	FROM [dbo].[ASRSysModuleSetup]
	WHERE moduleKey = 'MODULE_ABSENCE'
		AND parameterKey = 'Param_TableAbsenceType';

	IF NOT @sParameterValue IS null
	BEGIN
		SET @iTableID = convert(integer, @sParameterValue);

		SELECT @sTableName = tableName 
		FROM [dbo].[ASRSysTables]
		WHERE tableID = @iTableID;
		
 	END

	/* Get the Absence Type Column name. */
	SET @sParameterValue = '';

	SELECT @sParameterValue = parameterValue
	FROM [dbo].[ASRSysModuleSetup]
	WHERE moduleKey = 'MODULE_ABSENCE'
		AND parameterKey = 'Param_FieldTypeType';

	IF NOT @sParameterValue IS null
	BEGIN
		SET @iColumnID = convert(integer, @sParameterValue);

		SELECT @sColumnName = columnName 
		FROM [dbo].[ASRSysColumns]
		WHERE columnID = @iColumnID;
		
 	END

	/* Get the Absence Types if everything is ok. */
	IF len(@sTableName) > 0
		AND len(@sColumnName) > 0
	BEGIN
		SET @sExecString = 'SELECT ' + @sColumnName + 
			' FROM ' + @sTableName +
			' ORDER BY ' + @sColumnName;

		/* Return a recordset of the absence types */
		EXECUTE sp_executeSQL @sExecString;
		
	END
END