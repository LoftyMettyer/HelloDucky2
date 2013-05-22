CREATE PROCEDURE [dbo].[spASRIntGetUtilityName] (
	@piUtilityType	integer,
	@plngID			integer,
	@psName			varchar(255)	OUTPUT
)
AS
BEGIN
	DECLARE 
		@sTableName			sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500);

	SET @sTableName = '';
	SET @psName = '<unknown>';

	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sIDColumnName = 'ID';
    END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sIDColumnName = 'ID';
    END

	IF @piUtilityType = 1 /* Cross Tab */
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sIDColumnName = 'CrossTabID';
    END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sIDColumnName = 'ID';
    END
        
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sIDColumnName = 'DataTransferID';
    END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sIDColumnName = 'ID';
    END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sIDColumnName = 'functionID';
    END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sIDColumnName = 'ID';
    END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sIDColumnName = 'mailMergeID';
    END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sIDColumnName = 'recordProfileID';
    END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sIDColumnName = 'matchReportID';
    END

	IF (@piUtilityType = 25) /* Workflow */
	BEGIN
		SET @sTableName = 'ASRSysWorkflows';
		SET @sIDColumnName = 'ID';
	END
      	
	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @sName = [' + @sTableName + '].[name]
				FROM [' + @sTableName + ']
				WHERE [' + @sTableName + '].[' + @sIDColumnName + '] = ' + convert(nvarchar(255), @plngID);

		SET @sParamDefinition = N'@sName varchar(255) OUTPUT';
		EXEC sp_executesql @sSQL, @sParamDefinition, @psName OUTPUT;
	END

	IF @psName IS null SET @psName = '<unknown>';
END