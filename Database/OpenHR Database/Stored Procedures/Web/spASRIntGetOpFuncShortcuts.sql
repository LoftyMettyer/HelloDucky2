CREATE PROCEDURE [dbo].[spASRIntGetOpFuncShortcuts]
AS
BEGIN

	/* Return a recordset of the operators and functions that have shortcut keys. */
	DECLARE	@iFunctionID		integer, 
			@sParameter			varchar(MAX),
			@iLastFunctionID	integer,
			@sParameters		varchar(MAX);

	SET @iLastFunctionID = 0;
	SET @sParameters = '';

	DECLARE @tempParams TABLE(
		[functionID]	integer,
		[parameters]	varchar(MAX));

	DECLARE paramsCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysFunctionParameters.functionID, 
			ASRSysFunctionParameters.parameterName
		FROM ASRSysFunctionParameters
		INNER JOIN ASRSysFunctions ON ASRSysFunctionParameters.functionID = ASRSysFunctions.functionID
			AND LEN(ASRSysFunctions.shortcutKeys) > 0
		ORDER BY ASRSysFunctionParameters.functionID, ASRSysFunctionParameters.parameterIndex;

	OPEN paramsCursor;
	FETCH NEXT FROM paramsCursor INTO @iFunctionID, @sParameter;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iFunctionID <> @iLastFunctionID
		BEGIN
			IF LEN(@sParameters) >0 
			BEGIN
				INSERT INTO @tempParams ([functionID], [parameters]) VALUES(@iLastFunctionID, @sParameters);
			END

			SET @iLastFunctionID = @iFunctionID;
			SET @sParameters = @sParameter;
		END
		ELSE
		BEGIN
			SET @sParameters = @sParameters + char(9) + @sParameter;
		END

		FETCH NEXT FROM paramsCursor INTO @iFunctionID, @sParameter;
	END

	IF LEN(@sParameters) >0 
	BEGIN
		INSERT INTO @tempParams ([functionID], [parameters]) VALUES(@iLastFunctionID, @sParameters);
	END

	SET @iLastFunctionID = @iFunctionID;
	SET @sParameters = @sParameter;

	CLOSE paramsCursor;
	DEALLOCATE paramsCursor;
	
	SELECT 5 AS [componentType], 
		ASRSysOperators.operatorID AS [ID], 
		ASRSysOperators.shortcutKeys, 
		'' AS [params],
		name AS [name]
	FROM ASRSysOperators
	WHERE len(shortcutKeys) > 0 
	UNION
	SELECT 2 AS [componentType], 
		ASRSysFunctions.functionID AS [ID], 
		ASRSysFunctions.shortcutKeys, 
		tmp.[parameters] AS [params],
		functionName AS [name]
	FROM ASRSysFunctions
	INNER JOIN @tempParams tmp ON ASRSysFunctions.functionID = tmp.functionID
	WHERE len(shortcutKeys) > 0;

END