CREATE PROCEDURE spASRIntGetCalendarReportDefinition 
	(
	@piCalendarReportID 		integer, 
	@psCurrentUser				varchar(255),
	@psAction					varchar(255),
	@psErrorMsg					varchar(MAX)	OUTPUT,
	@psCalendarReportName		varchar(255)	OUTPUT,
	@psCalendarReportOwner		varchar(255)	OUTPUT,
	@psCalendarReportDesc		varchar(MAX)	OUTPUT,
	@piBaseTableID				integer			OUTPUT,
	@pfAllRecords				bit				OUTPUT,
	@piPicklistID				integer			OUTPUT,
	@psPicklistName				varchar(255)	OUTPUT,
	@pfPicklistHidden			bit				OUTPUT,
	@piFilterID					integer			OUTPUT,
	@psFilterName				varchar(255)	OUTPUT,
	@pfFilterHidden				bit				OUTPUT,
	@pfPrintFilterHeader		bit				OUTPUT,
	@piDesc1ID					integer			OUTPUT,
	@piDesc2ID					integer			OUTPUT,
	@piDescExprID				integer			OUTPUT,
	@psDescExprName				varchar(255)	OUTPUT,
	@pfDescCalcHidden			bit				OUTPUT,
	@piRegionID					integer			OUTPUT,
	@pfGroupByDesc				bit				OUTPUT,
	@pfDescSeparator			varchar(255)	OUTPUT,	
	@piStartType				integer			OUTPUT,
	@pdFixedStart				datetime		OUTPUT,
	@piStartFrequency			integer			OUTPUT,
	@piStartPeriod				integer			OUTPUT,
	@piCustomStartID			integer			OUTPUT,
	@psCustomStartName			varchar(MAX)	OUTPUT,
	@pfStartDateCalcHidden		bit				OUTPUT,
	@piEndType					integer			OUTPUT,
	@pdFixedEnd					datetime		OUTPUT,
	@piEndFrequency				integer			OUTPUT,
	@piEndPeriod				integer			OUTPUT,
	@piCustomEndID				integer			OUTPUT,
	@psCustomEndName			varchar(MAX)	OUTPUT,
	@pfEndDateCalcHidden		bit				OUTPUT,
	@pfShadeBHols				bit				OUTPUT,
	@pfShowCaptions				bit				OUTPUT,
	@pfShadeWeekends			bit				OUTPUT,
	@pfStartOnCurrentMonth		bit				OUTPUT,
	@pfIncludeWorkingDaysOnly	bit				OUTPUT,
	@pfIncludeBHols				bit				OUTPUT,
	@pfOutputPreview			bit				OUTPUT,
	@piOutputFormat				integer			OUTPUT,
	@pfOutputScreen				bit				OUTPUT,
	@pfOutputPrinter			bit				OUTPUT,
	@psOutputPrinterName		varchar(MAX)	OUTPUT,
	@pfOutputSave				bit				OUTPUT,
	@piOutputSaveExisting		integer			OUTPUT,
	@pfOutputEmail				bit				OUTPUT,
	@piOutputEmailAddr			integer			OUTPUT,
	@psOutputEmailName			varchar(MAX)	OUTPUT,
	@psOutputEmailSubject		varchar(MAX)	OUTPUT,
	@psOutputEmailAttachAs		varchar(MAX)	OUTPUT,
	@psOutputFilename			varchar(MAX)	OUTPUT,	
 	@piTimestamp				integer			OUTPUT
	)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iCount			integer,
			@sTempHidden	varchar(10),
			@sAccess 		varchar(10);

	SET @psErrorMsg = '';
	SET @psPicklistName = '';
	SET @pfPicklistHidden = 0;
	SET @psFilterName = '';
	SET @pfFilterHidden = 0;
	SET @pfDescCalcHidden = 0;
	SET @pfStartDateCalcHidden = 0;
	SET @pfEndDateCalcHidden = 0;
	
	/* Check the calendar report exists. */
	SELECT @iCount = COUNT(*)
	FROM [dbo].[ASRSysCalendarReports]
	WHERE ID = @piCalendarReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'calendar report has been deleted by another user.';
		RETURN;
	END

	SELECT	@psCalendarReportName = name,
					@psCalendarReportOwner = userName,
					@psCalendarReportDesc = description,
					@piBaseTableID = baseTable,
					@pfAllRecords = allRecords,
					@piPicklistID = picklist,
					@piFilterID = filter,
					@pfPrintFilterHeader = PrintFilterHeader,
					@piDesc1ID = Description1,
					@piDesc2ID = Description2,
					@piDescExprID = DescriptionExpr,
					@piRegionID = Region,
					@pfGroupByDesc = GroupByDesc,
					@pfDescSeparator = DescriptionSeparator,
					@piStartType = StartType,
					@pdFixedStart = FixedStart,
					@piStartFrequency = StartFrequency,
					@piStartPeriod = StartPeriod,
					@piCustomStartID = StartDateExpr,
					@piEndType = EndType,
					@pdFixedEnd = FixedEnd,
					@piEndFrequency = EndFrequency,
					@piEndPeriod = EndPeriod,
					@piCustomEndID = EndDateExpr,
					@pfShadeBHols = ShowBankHolidays,
					@pfShowCaptions = ShowCaptions,
					@pfShadeWeekends = ShowWeekends,
					@pfStartOnCurrentMonth = StartOnCurrentMonth,
					@pfIncludeWorkingDaysOnly	= IncludeWorkingDaysOnly,
					@pfIncludeBHols = IncludeBankHolidays,
					@pfOutputPreview = OutputPreview,
					@piOutputFormat = OutputFormat,
					@pfOutputScreen = OutputScreen,
					@pfOutputPrinter = OutputPrinter,
					@psOutputPrinterName = OutputPrinterName,
					@pfOutputSave = OutputSave,
					@piOutputSaveExisting = OutputSaveExisting,
					@pfOutputEmail = OutputEmail,
					@piOutputEmailAddr = OutputEmailAddr,
					@psOutputEmailSubject = ISNULL(OutputEmailSubject,''),
					@psOutputEmailAttachAs = ISNULL(OutputEmailAttachAs,''),
					@psOutputFilename = ISNULL(OutputFilename,''),
					@piTimestamp = convert(integer, timestamp)
	FROM [dbo].[ASRSysCalendarReports]
	WHERE ID = @piCalendarReportID;

	/* Check the current user can view the calendar report. */
	exec [dbo].[spASRIntCurrentUserAccess]
		17, 
		@piCalendarReportID,
		@sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psCalendarReportOwner <> @psCurrentUser) 
	BEGIN
		SET @psErrorMsg = 'calendar report has been made hidden by another user.';
		RETURN;
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psCalendarReportOwner <> @psCurrentUser) 
	BEGIN
		SET @psErrorMsg = 'calendar report has been made read only by another user.';
		RETURN;
	END

	/* Check the calendar report has details. */
	SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysCalendarReportEvents]
		WHERE ASRSysCalendarReportEvents.calendarReportID = @piCalendarReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'calendar report contains no details.';
		RETURN;
	END

	/* Check the calendar report has sort order details. */
	SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysCalendarReportOrder]
		WHERE ASRSysCalendarReportOrder.calendarReportID = @piCalendarReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'calendar report contains no sort order details.';
		RETURN;
	END

	IF @psAction = 'copy' 
	BEGIN
		SET @psCalendarReportName = left('copy of ' + @psCalendarReportName, 50);
		SET @psCalendarReportOwner = @psCurrentUser;
	END

	IF @piPicklistID > 0 
	BEGIN
		SELECT @psPicklistName = name,
			@sTempHidden = access
		FROM ASRSysPicklistName 
		WHERE picklistID = @piPicklistID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfPicklistHidden = 1;
		END
	END

	IF @piFilterID > 0 
	BEGIN
		SELECT @psFilterName = name,
			@sTempHidden = access
		FROM ASRSysExpressions 
		WHERE exprID = @piFilterID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfFilterHidden = 1;
		END
	END
	
	IF @piDescExprID > 0 
	BEGIN
		SELECT @psDescExprName = name,
			@sTempHidden = access
		FROM ASRSysExpressions 
		WHERE exprID = @piDescExprID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfDescCalcHidden = 1;
		END
	END
	
	IF @piCustomStartID > 0 
	BEGIN
		SELECT @psCustomStartName = name,
			@sTempHidden = access
		FROM ASRSysExpressions 
		WHERE exprID = @piCustomStartID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfStartDateCalcHidden = 1;
		END
	END
	
	IF @piCustomEndID > 0 
	BEGIN
		SELECT @psCustomEndName = name,
			@sTempHidden = access
		FROM ASRSysExpressions 
		WHERE exprID = @piCustomEndID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfEndDateCalcHidden = 1;
		END
	END


	IF @piOutputEmailAddr > 0
	BEGIN
		SELECT @psOutputEmailName = name,
			@sTempHidden = access
		FROM ASRSysEmailGroupName
		WHERE EmailGroupID = @piOutputEmailAddr;
	END
	ELSE
	BEGIN
		SET @piOutputEmailAddr = 0;
		SET @psOutputEmailName = '';
	END

	/* Get the calendar events definition recordset. */
	SELECT 
			ASRSysCalendarReportEvents.Name + char(9) + 
			CONVERT(varchar,ASRSysCalendarReportEvents.TableID) + char(9) +
			(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportEvents.TableID) + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.FilterID) + char(9) +
			CASE 
				WHEN ASRSysCalendarReportEvents.FilterID > 0 THEN
					(SELECT ISNULL(ASRSysExpressions.Name,'') FROM ASRSysExpressions WHERE ASRSysExpressions.ExprID = ASRSysCalendarReportEvents.FilterID) 
				ELSE
					''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventStartDateID) + char(9) +
			(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventStartDateID) + char(9) +
			
			CONVERT(varchar,ASRSysCalendarReportEvents.EventStartSessionID) + char(9) + 
			CASE 
				WHEN ASRSysCalendarReportEvents.EventStartSessionID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventStartSessionID)
				ELSE
					''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventEndDateID) + char(9) +
			CASE 
				WHEN ASRSysCalendarReportEvents.EventEndDateID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventEndDateID)
				ELSE ''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventEndSessionID) + char(9) + 
			CASE 
				WHEN ASRSysCalendarReportEvents.EventEndSessionID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventEndSessionID)
				ELSE
					''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventDurationID) + char(9) + 
			CASE 
				WHEN ASRSysCalendarReportEvents.EventDurationID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDurationID)
				ELSE 
					''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.LegendType) + char(9) +
			CASE 
				WHEN ASRSysCalendarReportEvents.LegendType = 1 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportEvents.LegendLookupTableID) + 
					'.' +
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.LegendLookupCodeID)
				ELSE
					ASRSysCalendarReportEvents.LegendCharacter
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.LegendLookupTableID) + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.LegendLookupColumnID) + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.LegendLookupCodeID) + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.LegendEventColumnID) + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventDesc1ColumnID) + char(9) +
			CASE 
				WHEN ASRSysCalendarReportEvents.EventDesc1ColumnID > 0 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ((SELECT ISNULL(ASRSysColumns.TableID,0) FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc1ColumnID))) + 
					'.' + 
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc1ColumnID)
				ELSE
					''
			END + char(9) +
			CONVERT(varchar,ASRSysCalendarReportEvents.EventDesc2ColumnID) + char(9) +
			CASE
				WHEN ASRSysCalendarReportEvents.EventDesc2ColumnID > 0 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID IN ((SELECT ISNULL(ASRSysColumns.TableID,0) FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc2ColumnID))) + 
					'.' + 
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc2ColumnID)
				ELSE
					''
	 		END + char(9) +
			ASRSysCalendarReportEvents.EventKey	 + char(9) +
			CASE 
				WHEN ASRSysCalendarReportEvents.FilterID > 0 THEN
			  		(SELECT CASE WHEN ASRSysExpressions.Access = 'HD' THEN 'Y' ELSE 'N' END FROM ASRSysExpressions WHERE ASRSysExpressions.ExprID = ASRSysCalendarReportEvents.FilterID) 
				ELSE
					'N'
			END	
			AS [DefinitionString],
		
			CASE 
				WHEN ASRSysCalendarReportEvents.FilterID > 0 THEN
			  		(SELECT CASE WHEN ASRSysExpressions.Access = 'HD' THEN 'Y' ELSE 'N' END FROM ASRSysExpressions WHERE ASRSysExpressions.ExprID = ASRSysCalendarReportEvents.FilterID) 
				ELSE
					'N'
			END AS [FilterHidden]
	FROM ASRSysCalendarReportEvents
	WHERE ASRSysCalendarReportEvents.CalendarReportID = @piCalendarReportID
	ORDER BY ASRSysCalendarReportEvents.ID;
END