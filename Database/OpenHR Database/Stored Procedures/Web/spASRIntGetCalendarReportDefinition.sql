CREATE PROCEDURE [dbo].[spASRIntGetCalendarReportDefinition] (
	@piCalendarReportID 		integer, 
	@psCurrentUser				varchar(255),
	@psAction					varchar(255))
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @psErrorMsg					varchar(MAX) = '',
		@psReportName		varchar(255) = '',
		@psReportOwner		varchar(255) = '',
		@psReportDesc		varchar(MAX) = '',
		@piBaseTableID				integer,
		@pfAllRecords				bit,
		@piPicklistID				integer,
		@psPicklistName				varchar(255) = '',
		@pfPicklistHidden			bit,
		@piFilterID					integer,
		@psFilterName				varchar(255) = '',
		@pfFilterHidden				bit,
		@pfPrintFilterHeader		bit,
		@piDesc1ID					integer,
		@piDesc2ID					integer,
		@piDescExprID				integer,
		@psDescExprName				varchar(255) = '',
		@pfDescCalcHidden			bit,
		@piRegionID					integer,
		@pfGroupByDesc				bit,
		@pfDescSeparator			varchar(255) = '',	
		@piStartType				integer,
		@pdFixedStart				datetime,
		@piStartFrequency			integer,
		@piStartPeriod				integer,
		@piCustomStartID			integer,
		@psCustomStartName			varchar(MAX) = '',
		@pfStartDateCalcHidden		bit,
		@piEndType					integer,
		@pdFixedEnd					datetime,
		@piEndFrequency				integer,
		@piEndPeriod				integer,
		@piCustomEndID				integer,
		@psCustomEndName			varchar(MAX) = '',
		@pfEndDateCalcHidden		bit,
		@pfShadeBHols				bit,
		@pfShowCaptions				bit,
		@pfShadeWeekends			bit,
		@pfStartOnCurrentMonth		bit,
		@pfIncludeWorkingDaysOnly	bit,
		@pfIncludeBHols				bit,
		@pfOutputPreview			bit,
		@piOutputFormat				integer,
		@pfOutputScreen				bit,
		@pfOutputPrinter			bit,
		@psOutputPrinterName		varchar(MAX) = '',
		@pfOutputSave				bit,
		@piOutputSaveExisting		integer		,
		@pfOutputEmail				bit,
		@piOutputEmailAddr			integer,
		@psOutputEmailName			varchar(MAX) = '',
		@psOutputEmailSubject		varchar(MAX) = '',
		@psOutputEmailAttachAs		varchar(MAX) = '',
		@psOutputFilename			varchar(MAX) = '',	
 		@piTimestamp				integer,
		@piCategoryID		integer;

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

	SELECT	@psReportName = name,
					@psReportOwner = userName,
					@psReportDesc = description,
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

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 
	BEGIN
		SET @psErrorMsg = 'calendar report has been made hidden by another user.';
		RETURN;
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 
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
		SET @psReportName = left('copy of ' + @psReportName, 50);
		SET @psReportOwner = @psCurrentUser;
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

	-- Get's the category id associated with the calendar report. Return 0 if not found
	SET @piCategoryID = 0
	SELECT @piCategoryID = ISNULL(categoryid,0)
		FROM [dbo].[tbsys_objectcategories]
		WHERE objectid = @piCalendarReportID AND objecttype = 17

	-- Definition
	SELECT @psReportName AS name, @psReportDesc AS [Description], @piCategoryID As CategoryID , @piBaseTableID AS baseTableID, @psReportOwner AS [Owner],
		CASE WHEN @pfAllRecords = 1 THEN 0 ELSE CASE WHEN ISNULL(@piPicklistID, 0) > 0 THEN 1 ELSE 2 END END AS [SelectionType],
		@piPicklistID AS PicklistID, @piFilterID AS FilterID,
		@psPicklistName AS PicklistName, @psFilterName AS FilterName,@pfPrintFilterHeader AS printFilterHeader,
		@piDesc1ID AS Description1ID, @piDesc2ID AS Description2ID, @piDescExprID AS Description3ID, @psDescExprName AS Description3Name,
		@piRegionID AS RegionID, @pfGroupByDesc AS GroupByDescription, @pfDescSeparator AS Separator,		
		@piStartType AS StartType, @pdFixedStart AS StartFixedDate, @piStartFrequency AS StartOffset, @piStartPeriod AS StartOffsetPeriod, @piCustomStartID AS StartCustomID, @psCustomStartName AS StartCustomName,
		@piEndType AS EndType, @pdFixedEnd AS EndFixedDate,	@piEndFrequency AS EndOffset, @piEndPeriod AS EndOffsetPeriod, @piCustomEndID AS EndCustomID,  @psCustomEndName AS EndCustomName,
		@pfShadeBHols AS ShowBankHolidays, @pfShowCaptions AS ShowCaptions,	@pfShadeWeekends AS ShowWeekends, @pfStartOnCurrentMonth AS StartOnCurrentMonth,
		@pfIncludeWorkingDaysOnly AS WorkingDaysOnly, @pfIncludeBHols AS IncludeBankHolidays,
		@pfOutputPreview AS IsPreview, @piOutputFormat AS [Format], @pfOutputScreen AS ToScreen, @pfOutputPrinter AS ToPrinter,
		@psOutputPrinterName AS PrinterName, @pfOutputSave AS SaveToFile, @piOutputSaveExisting AS SaveExisting,
		@pfOutputEmail AS SendToEmail, @piOutputEmailAddr AS EmailGroupID, @psOutputEmailName AS EmailGroupName,
		@psOutputEmailSubject AS EmailSubject, @psOutputEmailAttachAs AS EmailAttachmentName,
		@psOutputFilename AS [Filename], @piTimestamp AS [timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess],
		CASE WHEN @pfDescCalcHidden = 1 THEN 'HD' ELSE '' END AS [Description3ViewAccess],
		CASE WHEN @pfStartDateCalcHidden = 1 THEN 'HD' ELSE '' END AS [StartCustomViewAccess],
		CASE WHEN @pfEndDateCalcHidden = 1 THEN 'HD' ELSE '' END AS [EndCustomViewAccess];

	-- Calendar events definition recordset
	SELECT 
			ID, Name, TableID,
			(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportEvents.TableID) AS TableName,
			FilterID,
			CASE 
				WHEN ASRSysCalendarReportEvents.FilterID > 0 THEN
					(SELECT ISNULL(ASRSysExpressions.Name,'') FROM ASRSysExpressions WHERE ASRSysExpressions.ExprID = ASRSysCalendarReportEvents.FilterID) 
				ELSE
					''
			END AS FilterName,
			EventStartDateID,
			(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventStartDateID) AS EventStartDateName,			
			EventStartSessionID,
			CASE 
				WHEN EventStartSessionID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventStartSessionID)
				ELSE
					''
			END AS EventStartSessionName,
			CASE WHEN ISNULL(EventDurationID, 0) > 0 THEN 2 ELSE CASE WHEN ISNULL(EventEndDateID, 0) > 0 THEN 1 ELSE 0 END END AS [EventEndType],
			EventEndDateID,
			CASE 
				WHEN EventEndDateID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventEndDateID)
				ELSE ''
			END AS EventEndDateName,
			ASRSysCalendarReportEvents.EventEndSessionID, 
			CASE 
				WHEN ASRSysCalendarReportEvents.EventEndSessionID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventEndSessionID)
				ELSE
					''
			END AS EventEndSessionName,
			EventDurationID,
			CASE 
				WHEN ASRSysCalendarReportEvents.EventDurationID > 0 THEN
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDurationID)
				ELSE 
					''
			END AS EventDurationName,
			LegendType, LegendCharacter,
			CASE 
				WHEN ASRSysCalendarReportEvents.LegendType = 1 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportEvents.LegendLookupTableID) + 
					'.' +
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.LegendLookupCodeID)
				ELSE
					ASRSysCalendarReportEvents.LegendCharacter
			END LegendTypeName,
			LegendLookupTableID, LegendLookupColumnID, LegendLookupCodeID, LegendEventColumnID, EventDesc1ColumnID,
			CASE 
				WHEN ASRSysCalendarReportEvents.EventDesc1ColumnID > 0 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ((SELECT ISNULL(ASRSysColumns.TableID,0) FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc1ColumnID))) + 
					'.' + 
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc1ColumnID)
				ELSE
					''
			END AS EventDesc1ColumnName,
			EventDesc2ColumnID,
			CASE
				WHEN EventDesc2ColumnID > 0 THEN
					(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID IN ((SELECT ISNULL(ASRSysColumns.TableID,0) FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc2ColumnID))) + 
					'.' + 
					(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportEvents.EventDesc2ColumnID)
				ELSE
					''
	 		END AS EventDesc2ColumnName,
			EventKey,
			CASE 
				WHEN ASRSysCalendarReportEvents.FilterID > 0 THEN
			  		(SELECT CASE WHEN ASRSysExpressions.Access = 'HD' THEN 'HD' ELSE 'RW' END FROM ASRSysExpressions WHERE ASRSysExpressions.ExprID = ASRSysCalendarReportEvents.FilterID) 
				ELSE
					'RW'
			END AS FilterViewAccess
	FROM ASRSysCalendarReportEvents
	WHERE CalendarReportID = @piCalendarReportID
	ORDER BY ID;

	-- Orders
	SELECT 
		ColumnID AS Id, TableID, 
		(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportOrder.ColumnID) AS [Name],
		OrderSequence AS [Sequence],
		OrderType AS [Order]
	FROM [dbo].[ASRSysCalendarReportOrder]
	WHERE calendarReportID = @piCalendarReportID
	ORDER BY OrderSequence;

END
