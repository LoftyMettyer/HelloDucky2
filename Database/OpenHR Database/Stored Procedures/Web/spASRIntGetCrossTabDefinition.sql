CREATE PROCEDURE [dbo].[spASRIntGetCrossTabDefinition] (
	@piReportID 			integer, 
	@psCurrentUser			varchar(255),
	@psAction				varchar(255))
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @psErrorMsg				varchar(MAX) = '',
			@psReportName			varchar(255) = '',
			@psReportOwner			varchar(255) = '',
			@psReportDesc			varchar(MAX) = '',
			@piBaseTableID			integer = 0,
			@piSelection			integer = 0,
			@piPicklistID			integer = 0,
			@psPicklistName			varchar(255) = '',
			@pfPicklistHidden		bit,
			@piFilterID				integer = 0,
			@psFilterName			varchar(255) = '',
			@pfFilterHidden			bit,
			@pfPrintFilterHeader	bit,
			@HColID					integer = 0,
			@HStart					varchar(20) = '',
			@HStop					varchar(20) = '',
			@HStep					varchar(20) = '',
			@VColID					integer = 0,
			@VStart					varchar(20) = '',
			@VStop					varchar(20) = '',
			@VStep					varchar(20) = '',
			@PColID					integer = 0,
			@PStart					varchar(20) = '',
			@PStop					varchar(20) = '',
			@PStep					varchar(20) = '',
			@IType					integer = 0,
			@IColID					integer = 0,
			@Percentage				bit,
			@PerPage				bit,
			@Suppress				bit,
			@Thousand				bit,
			@pfOutputPreview		bit,
			@piOutputFormat			integer = 0,
			@pfOutputScreen			bit,
			@pfOutputPrinter		bit,
			@psOutputPrinterName	varchar(MAX) = '',
			@pfOutputSave			bit,
			@piOutputSaveExisting	integer = 0,
			@pfOutputEmail			bit,
			@piOutputEmailAddr		integer = 0,
			@psOutputEmailName		varchar(MAX) = '',
			@psOutputEmailSubject	varchar(MAX) = '',
			@psOutputEmailAttachAs	varchar(MAX) = '',
			@psOutputFilename		varchar(MAX) = '',
 			@piTimestamp			integer	= 0,
			@piCategoryID		integer;	

	DECLARE	@iCount			integer,
			@sTempHidden	varchar(MAX),
			@sAccess 		varchar(MAX);


	/* Check the report exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysCrossTab
	WHERE CrossTabID = @piReportID

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'cross tab has been deleted by another user.'
		RETURN
	END

	SELECT @psReportName = name, @psReportDesc	 = description, @psReportOwner = userName,
		@piBaseTableID = TableID, @piSelection = Selection, @piPicklistID = PicklistID,
		@piFilterID = FilterID,	@pfPrintFilterHeader = PrintFilterHeader, @psReportOwner = userName,
		@HColID = HorizontalColID, @HStart = HorizontalStart, @HStop = HorizontalStop, @HStep = HorizontalStep,
		@VColID = VerticalColID, @VStart = VerticalStart, @VStop = VerticalStop, @VStep = VerticalStep,
		@PColID = PageBreakColID, @PStart = PageBreakStart,	@PStop = PageBreakStop,	@PStep = PageBreakStep,
		@IType = IntersectionType, @IColID = IntersectionColID,	@Percentage = Percentage, @PerPage = PercentageofPage,
		@Suppress = SuppressZeros,@Thousand = ThousandSeparators,
		@pfOutputPreview = OutputPreview, @piOutputFormat = OutputFormat, @pfOutputScreen = OutputScreen,
		@pfOutputPrinter = OutputPrinter, @psOutputPrinterName = OutputPrinterName,
		@pfOutputSave = OutputSave,	@piOutputSaveExisting = OutputSaveExisting,
		@pfOutputEmail = OutputEmail, @piOutputEmailAddr = OutputEmailAddr,
		@psOutputEmailSubject = ISNULL(OutputEmailSubject,''),
		@psOutputEmailAttachAs = ISNULL(OutputEmailAttachAs,''),
		@psOutputFilename = ISNULL(OutputFilename,''),
		@piTimestamp = convert(integer, timestamp)
	FROM ASRSysCrossTab
	WHERE CrossTabID = @piReportID;

	/* Check the current user can view the report. */
	EXEC spASRIntCurrentUserAccess 	1, @piReportID,	@sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 
		SET @psErrorMsg = 'cross tab has been made hidden by another user.';

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 
		SET @psErrorMsg = 'cross tab has been made read only by another user.';

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

	-- Get's the category id associated with the crossTab report. Return 0 if not found
	SET @piCategoryID = 0
	SELECT @piCategoryID = ISNULL(categoryid,0)
		FROM [dbo].[tbsys_objectcategories]
		WHERE objectid = @piReportID AND objecttype = 1

	SELECT @psErrorMsg AS ErrorMsg, @psReportName AS Name, @psReportOwner AS [Owner], @psReportDesc AS [Description]
		, @piBaseTableID AS [BaseTableID],  @piCategoryID As CategoryID, @piSelection AS SelectionType
		, @piPicklistID AS PicklistID, @psPicklistName AS PicklistName, @pfPicklistHidden AS [IsPicklistHidden]
		, @piFilterID AS FilterID, @psFilterName AS [FilterName], @pfFilterHidden AS [IsFilterHidden]
		, @pfPrintFilterHeader AS [PrintFilterHeader]
		, @HColID AS HorizontalID, @HStart AS HorizontalStart, @HStop AS HorizontalStop, @HStep AS HorizontalIncrement
		, @VColID AS VerticalID, @VStart AS VerticalStart, @VStop AS VerticalStop, @VStep AS VerticalIncrement
		, @PColID AS PageBreakID, @PStart AS PageBreakStart, @PStop AS PageBreakStop, @PStep AS PageBreakIncrement
		, @IType AS IntersectionType, @IColID AS IntersectionID
		, @Percentage AS PercentageOfType, @PerPage AS PercentageOfPage
		, @Suppress	AS SuppressZeros, @Thousand AS [UseThousandSeparators]
		, @pfOutputPreview AS IsPreview, @piOutputFormat AS [Format],	@pfOutputScreen AS [ToScreen]
		, @pfOutputPrinter AS [ToPrinter], @psOutputPrinterName	AS [PrinterName]
		, @pfOutputSave AS [SaveToFile], @piOutputSaveExisting AS [SaveExisting]
		, @pfOutputEmail AS [SendToEmail], @piOutputEmailAddr AS [EmailGroupID], @psOutputEmailName AS [EmailGroupName]
		, @psOutputEmailSubject AS [EmailSubject], @psOutputEmailAttachAs AS [EmailAttachmentName]
		, @psOutputFilename AS [FileName], @piTimestamp AS [Timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess];

END
