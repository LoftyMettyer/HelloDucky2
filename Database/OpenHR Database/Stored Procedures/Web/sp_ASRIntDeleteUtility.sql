CREATE PROCEDURE [dbo].[sp_ASRIntDeleteUtility] (
	@piUtilType	integer,
	@piUtilID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iExprID	integer;

	IF @piUtilType = 0
	BEGIN
		/* Batch Jobs */
		DELETE FROM ASRSysBatchJobName WHERE ID = @piUtilID;
		DELETE FROM ASRSysBatchJobDetails WHERE BatchJobNameID = @piUtilID;
		DELETE FROM ASRSysBatchJobAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 1 OR @piUtilType = 35
	BEGIN
		/* Cross Tabs or 9-Box Grid*/
		DELETE FROM ASRSysCrossTab WHERE CrossTabID = @piUtilID;
		DELETE FROM ASRSysCrossTabAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 38
	BEGIN
		/* Talent Reports*/
		DELETE FROM ASRSysTalentReports WHERE ID = @piUtilID;
		DELETE FROM ASRSysTalentReportAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 2
	BEGIN
		/* Custom Reports. */
		DELETE FROM ASRSysCustomReportsName WHERE id = @piUtilID;
		DELETE FROM ASRSysCustomReportsDetails WHERE customReportID= @piUtilID;
		DELETE FROM ASRSysCustomReportAccess WHERE ID = @piUtilID;
	END
	
	IF @piUtilType = 3
	BEGIN
		/* Data Transfer. */
		DELETE FROM ASRSysDataTransferName WHERE DataTransferID = @piUtilID;
		DELETE FROM ASRSysDataTransferAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 4
	BEGIN
		/* Export. */
		DELETE FROM ASRSysExportName WHERE ID = @piUtilID;
		DELETE FROM ASRSysExportDetails WHERE ExportID = @piUtilID;
		DELETE FROM ASRSysExportAccess WHERE ID = @piUtilID;
	END

	IF (@piUtilType = 5) OR (@piUtilType = 6) OR (@piUtilType = 7)
	BEGIN
		/* Globals. */
		DELETE FROM ASRSysGlobalFunctions  WHERE FunctionID = @piUtilID;
		DELETE FROM ASRSysGlobalAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 8
	BEGIN
		/* Import. */
		DELETE FROM ASRSysImportName  WHERE ID = @piUtilID;
		DELETE FROM ASRSysImportDetails WHERE ImportID = @piUtilID;
		DELETE FROM ASRSysImportAccess WHERE ID = @piUtilID;
	END

	IF (@piUtilType = 9) OR (@piUtilType = 18)
	BEGIN
		/* Mail Merge/ Envelopes & Labels. */
		DELETE FROM ASRSysMailMergeName  WHERE MailMergeID = @piUtilID;
		DELETE FROM ASRSysMailMergeColumns  WHERE MailMergeID = @piUtilID;
		DELETE FROM ASRSysMailMergeAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 10
	BEGIN
		/* Picklists. */
		DELETE FROM ASRSysPickListName WHERE picklistID = @piUtilID;
		DELETE FROM ASRSysPickListItems WHERE picklistID = @piUtilID;
	END
	
	IF @piUtilType = 11 OR @piUtilType = 12
	BEGIN
		/* Filters and Calculations. */
		DECLARE subExpressions_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExpressions.exprID
			FROM ASRSysExpressions
			INNER JOIN ASRSysExprComponents ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID
			AND ASRSysExprComponents.exprID = @piUtilID;
		OPEN subExpressions_cursor;
		FETCH NEXT FROM subExpressions_cursor INTO @iExprID;
		WHILE (@@fetch_status = 0)
		BEGIN
			exec [dbo].[sp_ASRIntDeleteUtility] @piUtilType, @iExprID;
			
			FETCH NEXT FROM subExpressions_cursor INTO @iExprID;
		END
		CLOSE subExpressions_cursor;
		DEALLOCATE subExpressions_cursor;

		DELETE FROM ASRSysExprComponents
		WHERE exprID = @piUtilID;

		DELETE FROM ASRSysExpressions WHERE exprID = @piUtilID;
	END	

	IF (@piUtilType = 14) OR (@piUtilType = 23) OR (@piUtilType = 24)
	BEGIN
		/* Match Reports/Succession Planning/Career Progression. */
		DELETE FROM ASRSysMatchReportName WHERE MatchReportID = @piUtilID;
		DELETE FROM ASRSysMatchReportAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 17 
	BEGIN
		/*Calendar Reports*/
		DELETE FROM ASRSysCalendarReports WHERE ID = @piUtilID;
		DELETE FROM ASRSysCalendarReportEvents WHERE CalendarReportID = @piUtilID;
		DELETE FROM ASRSysCalendarReportOrder WHERE CalendarReportID = @piUtilID;
		DELETE FROM ASRSysCalendarReportAccess WHERE ID = @piUtilID;
	END
	
	IF @piUtilType = 20 
	BEGIN
		/*Record Profile*/
		DELETE FROM ASRSysRecordProfileName WHERE recordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileDetails WHERE RecordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileTables WHERE RecordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileAccess WHERE ID = @piUtilID;
	END
	
END