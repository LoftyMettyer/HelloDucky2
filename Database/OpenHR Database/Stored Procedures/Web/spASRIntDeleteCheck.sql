CREATE PROCEDURE [dbo].[spASRIntDeleteCheck] (
	@piUtilityType	integer,
	@plngID			integer,
	@pfDeleted		bit				OUTPUT,
	@psAccess		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@sTableName			sysname,
		@sAccessTableName	sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@fNewAccess			bit,
		@iCount				integer,
		@sAccess			varchar(MAX),
		@fSysSecMgr			bit;

	SET @sTableName = '';
	SET @psAccess = 'HD';
	SET @pfDeleted = 0;
	SET @fNewAccess = 0;

	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sAccessTableName = 'ASRSysBatchJobAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sAccessTableName = 'ASRSysCalendarReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END

	IF @piUtilityType = 1 OR @piUtilityType = 35 /* Cross Tab or 9-Box Grid*/
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sAccessTableName = 'ASRSysCrossTabAccess';
		SET @sIDColumnName = 'CrossTabID';
		SET @fNewAccess = 1;
 	END

	IF @piUtilityType = 38 /* Talent Management Report*/
	BEGIN
		SET @sTableName = 'ASRSysTalentReports';
		SET @sAccessTableName = 'ASRSysTalentReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
 	END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sAccessTableName = 'ASRSysCustomReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
 	END
    
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sAccessTableName = 'ASRSysDataTransferAccess';
		SET @sIDColumnName = 'DataTransferID';
		SET @fNewAccess = 1;
  END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sAccessTableName = 'ASRSysExportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sAccessTableName = 'ASRSysGlobalAccess';
		SET @sIDColumnName = 'functionID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sAccessTableName = 'ASRSysImportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sAccessTableName = 'ASRSysMailMergeAccess';
		SET @sIDColumnName = 'mailMergeID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sAccessTableName = 'ASRSysRecordProfileAccess';
		SET @sIDColumnName = 'recordProfileID';
		SET @fNewAccess = 1
  END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sAccessTableName = 'ASRSysMatchReportAccess';
		SET @sIDColumnName = 'matchReportID';
		SET @fNewAccess = 1;
  END

	IF (@piUtilityType = 11) OR (@piUtilityType = 12)  /* Filters/Calcs */
	BEGIN
		SET @sTableName = 'ASRSysExpressions';
		SET @sIDColumnName = 'exprID';
  END

	IF (@piUtilityType = 10)  /* Picklists */
	BEGIN
		SET @sTableName = 'ASRSysPicklistName';
		SET @sIDColumnName = 'picklistID';
  END

  IF @piUtilityType = 39 /* Organisation Report*/
	BEGIN
		SET @sTableName = 'ASRSysOrganisationReport';
		SET @sAccessTableName = 'ASRSysOrganisationReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;		
 	END

	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @iCount = COUNT(*)
				FROM ' + @sTableName + 
				' WHERE ' + @sTableName + '.' + @sIDColumnName + ' = ' + convert(nvarchar(255), @plngID);
		SET @sParamDefinition = N'@iCount integer OUTPUT';
		EXEC sp_executesql @sSQL,  @sParamDefinition, @iCount OUTPUT;

		IF @iCount = 0 
		BEGIN
			SET @pfDeleted = 1;
		END
		ELSE
		BEGIN
			IF @fNewAccess = 1
			BEGIN
				exec [dbo].[spASRIntCurrentUserAccess] @piUtilityType,	@plngID, @psAccess OUTPUT;
			END
			ELSE
			BEGIN
				exec [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
				
				IF @fSysSecMgr = 1 
				BEGIN
					SET @psAccess = 'RW';
				END
				ELSE
				BEGIN
					SET @sSQL = 'SELECT @sAccess = CASE 
								WHEN userName = system_user THEN ''RW''
								ELSE access
							END
							FROM ' + @sTableName + 
							' WHERE ' + @sTableName + '.' + @sIDColumnName + ' = ' + convert(nvarchar(255), @plngID);
					SET @sParamDefinition = N'@sAccess varchar(MAX) OUTPUT';
					EXEC sp_executesql @sSQL,  @sParamDefinition, @sAccess OUTPUT;

					SET @psAccess = @sAccess;
				END
			END
		END
	END
END