CREATE PROCEDURE [dbo].[spASRIntGetUtilityBaseTable] (
	@piUtilityType	integer,
	@plngID			integer,
	@piBaseTableID	integer	OUTPUT
)
AS
BEGIN
	DECLARE 
		@sTableName				sysname,
		@sIDColumnName			sysname,
		@sBaseTableColumnName	sysname,
		@sSQL					nvarchar(MAX),
		@sParamDefinition		nvarchar(500);

	SET @sTableName = '';
	SET @piBaseTableID = 0;
	SET @sBaseTableColumnName = '';
	
	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sIDColumnName = 'ID';
		/* No base table for batch jobs. */
 	END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sIDColumnName = 'ID';
		SET @sBaseTableColumnName = 'BaseTable';
 	END

	IF @piUtilityType = 1 /* Cross Tab */
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sIDColumnName = 'CrossTabID';
		SET @sBaseTableColumnName = 'TableID';
 	END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sIDColumnName = 'ID';
		SET @sBaseTableColumnName = 'BaseTable';
 	END
    
    
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sIDColumnName = 'DataTransferID';
		SET @sBaseTableColumnName = 'FromTableID';
 	END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sIDColumnName = 'ID';
		SET @sBaseTableColumnName = 'BaseTable';
 	END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sIDColumnName = 'functionID';
		SET @sBaseTableColumnName = 'TableID';
 	END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sIDColumnName = 'ID';
		SET @sBaseTableColumnName = 'BaseTable';
 	END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sIDColumnName = 'mailMergeID';
		SET @sBaseTableColumnName = 'TableID';
 	END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sIDColumnName = 'recordProfileID';
		SET @sBaseTableColumnName = 'BaseTable';
 	END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sIDColumnName = 'matchReportID';
		SET @sBaseTableColumnName = 'Table1ID';
 	END

	IF (len(@sTableName) > 0) 
		AND (len(@sBaseTableColumnName) > 0)
	BEGIN
		SET @sSQL = 'SELECT @iTableID = [' + @sTableName + '].[' + @sBaseTableColumnName + ']
				FROM [' + @sTableName + ']
				WHERE [' + @sTableName + '].[' + @sIDColumnName + '] = ' + convert(nvarchar(100), @plngID);

		SET @sParamDefinition = N'@iTableID integer OUTPUT';
		EXEC sp_executesql @sSQL, @sParamDefinition, @piBaseTableID OUTPUT;
	END

	IF @piBaseTableID IS null SET @piBaseTableID = 0;
END