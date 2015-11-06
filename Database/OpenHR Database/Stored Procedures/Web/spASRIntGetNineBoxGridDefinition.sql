CREATE PROCEDURE [dbo].[spASRIntGetNineBoxGridDefinition] (
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
			@XAxisLabel varchar(255) = '',
			@XAxisSubLabel1 varchar(255) = '',
			@XAxisSubLabel2 varchar(255) = '',
			@XAxisSubLabel3 varchar(255) = '',
			@YAxisLabel varchar(255) = '',
			@YAxisSubLabel1 varchar(255) = '',
			@YAxisSubLabel2 varchar(255) = '',
			@YAxisSubLabel3 varchar(255) = '',
			@Description1 varchar(255) = '',
			@ColorDesc1 varchar(6) = '',
			@Description2 varchar(255) = '',
			@ColorDesc2 varchar(6) = '',
			@Description3 varchar(255) = '',
			@ColorDesc3 varchar(6) = '',
			@Description4 varchar(255) = '',
			@ColorDesc4 varchar(6) = '',
			@Description5 varchar(255) = '',
			@ColorDesc5 varchar(6) = '',
			@Description6 varchar(255) = '',
			@ColorDesc6 varchar(6) = '',
			@Description7 varchar(255) = '',
			@ColorDesc7 varchar(6) = '',
			@Description8 varchar(255) = '',
			@ColorDesc8 varchar(6) = '',
			@Description9 varchar(255) = '',
			@ColorDesc9 varchar(6),
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
		SET @psErrorMsg = '9-Box Grid has been deleted by another user.'
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
		@piTimestamp = convert(integer, timestamp),
		@XAxisLabel = XAxisLabel,
		@XAxisSubLabel1 = XAxisSubLabel1,
		@XAxisSubLabel2 = XAxisSubLabel2,
		@XAxisSubLabel3 = XAxisSubLabel3,
		@YAxisLabel = YAxisLabel,
		@YAxisSubLabel1 = YAxisSubLabel1,
		@YAxisSubLabel2 = YAxisSubLabel2,
		@YAxisSubLabel3 = YAxisSubLabel3,
		@Description1 = Description1,
		@ColorDesc1 = ColorDesc1,
		@Description2 = Description2,
		@ColorDesc2 = ColorDesc2,
		@Description3 = Description3,
		@ColorDesc3 = ColorDesc3,
		@Description4 = Description4,
		@ColorDesc4 = ColorDesc4,
		@Description5 = Description5,
		@ColorDesc5 = ColorDesc5,
		@Description6 = Description6,
		@ColorDesc6 = ColorDesc6,
		@Description7 = Description7,
		@ColorDesc7 = ColorDesc7,
		@Description8 = Description8,
		@ColorDesc8 = ColorDesc8,
		@Description9 = Description9,
		@ColorDesc9 = ColorDesc9
	FROM ASRSysCrossTab
	WHERE CrossTabID = @piReportID;

	/* Check the current user can view the report. */
	EXEC spASRIntCurrentUserAccess 	1, @piReportID,	@sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 
		SET @psErrorMsg = '9-Box Grid has been made hidden by another user.';

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 
		SET @psErrorMsg = '9-Box Grid has been made read only by another user.';

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

	-- Get's the category id associated with the nine box grid report. Return 0 if not found
	SET @piCategoryID = 0
	SELECT @piCategoryID = ISNULL(categoryid,0)
		FROM [dbo].[tbsys_objectcategories]
		WHERE objectid = @piReportID AND objecttype = 35


	SELECT @psErrorMsg AS ErrorMsg, @psReportName AS Name, @psReportOwner AS [Owner], @psReportDesc AS [Description]
		, @piBaseTableID AS [BaseTableID], @piCategoryID As CategoryID, @piSelection AS SelectionType
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
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess],
		@XAxisLabel AS XAxisLabel,
		@XAxisSubLabel1 AS XAxisSubLabel1,
		@XAxisSubLabel2 AS XAxisSubLabel2,
		@XAxisSubLabel3 AS XAxisSubLabel3,
		@YAxisLabel AS YAxisLabel,
		@YAxisSubLabel1 AS YAxisSubLabel1,
		@YAxisSubLabel2 AS YAxisSubLabel2,
		@YAxisSubLabel3 AS YAxisSubLabel3,
		@Description1 AS Description1,
		@ColorDesc1 AS ColorDesc1,
		@Description2 AS Description2,
		@ColorDesc2 AS ColorDesc2,
		@Description3 AS Description3,
		@ColorDesc3 AS ColorDesc3,
		@Description4 AS Description4,
		@ColorDesc4 AS ColorDesc4,
		@Description5 AS Description5,
		@ColorDesc5 AS ColorDesc5,
		@Description6 AS Description6,
		@ColorDesc6 AS ColorDesc6,
		@Description7 AS Description7,
		@ColorDesc7 AS ColorDesc7,
		@Description8 AS Description8,
		@ColorDesc8 AS ColorDesc8,
		@Description9 AS Description9,
		@ColorDesc9 AS ColorDesc9;
END

