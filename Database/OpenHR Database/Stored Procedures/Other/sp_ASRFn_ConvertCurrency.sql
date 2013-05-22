CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency]
(
	  @pfResult		float OUTPUT
	, @pfValue		float
	, @psFromCurr	varchar(MAX)
	, @psToCurr		varchar(MAX)
)
AS
BEGIN

	DECLARE @sCConvTable 			SysName
			, @sCConvExRateCol		SysName
			, @sCConvCurrDescCol	SysName
			, @sCConvDecCol			SysName
			, @sCommandString		nvarchar(MAX)
			, @sParamDefinition		nvarchar(500);
	
	-- Get the name of the Currency Conversion table and Currency Description column.
	SELECT @sCConvCurrDescCol = ASRSysColumns.ColumnName, @sCConvTable = ASRSysTables.TableName 
	FROM ASRSysModuleSetup 
   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
            INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID 
	WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY'  AND  ASRSysModuleSetup.ParameterKey = 'Param_CurrencyNameColumn';

	-- Get the name of the Exchange Rate column.
	SELECT @sCConvExRateCol = ASRSysColumns.ColumnName
	FROM ASRSysModuleSetup 
   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
		WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY'  AND  ASRSysModuleSetup.ParameterKey = 'Param_ConversionValueColumn';

	-- Get the name of the Decimals column.
	SELECT @sCConvDecCol = ASRSysColumns.ColumnName
	FROM ASRSysModuleSetup 
   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
		WHERE ASRSysModuleSetup.ModuleKey = 'MODULE_CURRENCY'  AND  ASRSysModuleSetup.ParameterKey = 'Param_DecimalColumn';

	IF (NOT @sCConvTable IS NULL) AND (NOT @sCConvCurrDescCol IS NULL) AND (NOT @sCConvExRateCol IS NULL) AND (NOT @sCConvDecCol IS NULL) 

	-- Create the SQL string that returns the Coverted value.
	BEGIN
		SET @sCommandString = 'SELECT @pfResult = ROUND(ISNULL((' + LTRIM(RTRIM(STR(@pfValue,20,6))) 
										+ ' / NULLIF((SELECT ' + @sCConvTable + '.' + @sCConvExRateCol
									 			  + ' FROM ' + @sCConvTable
												  + ' WHERE ' + @sCConvTable + '.' + @sCConvCurrDescCol + ' = ''' + @psFromCurr + '''), 0))'
												  + ' * '
												  + '(SELECT ' + @sCConvTable + '.' + @sCConvExRateCol
												  + ' FROM ' + @sCConvTable
												  + ' WHERE ' + @sCConvTable + '.' + @sCConvCurrDescCol + ' = ''' + @psToCurr + '''), 0)'
												  + ' , '
												  + ' ISNULL('
												  + '(SELECT ' + @sCConvTable + '.' + @sCConvDecCol
												  + ' FROM ' + @sCConvTable
												  + ' WHERE ' + @sCConvTable + '.' + @sCConvCurrDescCol + ' = ''' + @psToCurr + '''), 0))';

		SET @sParamDefinition = N'@pfResult float output';

		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfResult output;
	END
	ELSE
	BEGIN
		SET @pfResult = NULL;
	END

END



