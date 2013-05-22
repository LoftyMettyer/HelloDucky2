CREATE PROCEDURE [dbo].[sp_ASRIntGetCrossTabDefinition] (
	@piReportID 			integer, 
	@psCurrentUser			varchar(255),
	@psAction				varchar(255),
	@psErrorMsg				varchar(MAX)	OUTPUT,
	@psReportName			varchar(255)	OUTPUT,
	@psReportOwner			varchar(255)	OUTPUT,
	@psReportDesc			varchar(MAX)	OUTPUT,
	@piBaseTableID			integer			OUTPUT,
	@pfAllRecords			bit				OUTPUT,
	@piPicklistID			integer			OUTPUT,
	@psPicklistName			varchar(255)	OUTPUT,
	@pfPicklistHidden		bit				OUTPUT,
	@piFilterID				integer			OUTPUT,
	@psFilterName			varchar(255)	OUTPUT,
	@pfFilterHidden			bit				OUTPUT,
	@pfPrintFilterHeader	bit				OUTPUT,
	@HColID					integer			OUTPUT,
	@HStart					varchar(20)		OUTPUT,
	@HStop					varchar(20)		OUTPUT,
	@HStep					varchar(20)		OUTPUT,
	@VColID					integer			OUTPUT,
	@VStart					varchar(20)		OUTPUT,
	@VStop					varchar(20)		OUTPUT,
	@VStep					varchar(20)		OUTPUT,
	@PColID					integer			OUTPUT,
	@PStart					varchar(20)		OUTPUT,
	@PStop					varchar(20)		OUTPUT,
	@PStep					varchar(20)		OUTPUT,
	@IType					integer			OUTPUT,
	@IColID					integer			OUTPUT,
	@Percentage				bit				OUTPUT,
	@PerPage				bit				OUTPUT,
	@Suppress				bit				OUTPUT,
	@Thousand				bit				OUTPUT,
	@pfOutputPreview		bit				OUTPUT,
	@piOutputFormat			integer			OUTPUT,
	@pfOutputScreen			bit				OUTPUT,
	@pfOutputPrinter		bit				OUTPUT,
	@psOutputPrinterName	varchar(MAX)	OUTPUT,
	@pfOutputSave			bit				OUTPUT,
	@piOutputSaveExisting	integer			OUTPUT,
	@pfOutputEmail			bit				OUTPUT,
	@piOutputEmailAddr		integer			OUTPUT,
	@psOutputEmailName		varchar(MAX)	OUTPUT,
	@psOutputEmailSubject	varchar(MAX)	OUTPUT,
	@psOutputEmailAttachAs	varchar(MAX)	OUTPUT,
	@psOutputFilename		varchar(MAX)	OUTPUT,
 	@piTimestamp			integer			OUTPUT
)

AS
BEGIN
	DECLARE	@iCount			integer,
			@sTempHidden	varchar(MAX),
			@sAccess 		varchar(MAX);

	SET @psErrorMsg = ''
	--SET @psCurrentUser = ''
	--SET @psAction = ''
	SET @psErrorMsg = ''
	SET @psReportName = ''
	SET @psReportOwner = ''
	SET @psReportDesc = ''
	SET @piBaseTableID = 0
	SET @pfAllRecords = 0
	SET @piPicklistID = 0
	SET @psPicklistName = ''
	SET @pfPicklistHidden = 0
	SET @piFilterID = 0
	SET @psFilterName = ''
	SET @pfFilterHidden = 0
	SET @pfPrintFilterHeader = 0
	SET @HColID = 0
	SET @HStart = ''
	SET @HStop = ''
	SET @HStep = ''
	SET @VColID = 0
	SET @VStart = ''
	SET @VStop = ''
	SET @VStep = ''
	SET @PColID = 0
	SET @PStart = ''
	SET @PStop = ''
	SET @PStep = ''
	SET @IType = 0
	SET @IColID = 0
	SET @Percentage = 0
	SET @PerPage = 0
	SET @Suppress = 0
	SET @Thousand = 0
	SET @pfOutputPreview = 0
	SET @piOutputFormat = 0
	SET @pfOutputScreen = 0
	SET @pfOutputPrinter = 0
	SET @psOutputPrinterName = ''
	SET @pfOutputSave = 0
	SET @piOutputSaveExisting = 0
	SET @pfOutputEmail = 0
	SET @piOutputEmailAddr = 0
	SET @psOutputEmailName = ''
	SET @psOutputEmailSubject = ''
	SET @psOutputEmailAttachAs = ''
	SET @psOutputFilename = ''
 	SET @piTimestamp = 0


	/* Check the report exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysCrossTab
	WHERE CrossTabID = @piReportID

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'cross tab has been deleted by another user.'
		RETURN
	END

	SELECT @psReportName = name,
		@psReportDesc	 = description,
		@psReportOwner = userName,
		@piBaseTableID = TableID,
		@pfAllRecords = Selection,
		@piPicklistID = PicklistID,
		@piFilterID = FilterID,
		@pfPrintFilterHeader = PrintFilterHeader,
		@psReportOwner = userName,
		@HColID = HorizontalColID,
		@HStart = HorizontalStart,
		@HStop = HorizontalStop,
		@HStep = HorizontalStep,
		@VColID = VerticalColID,
		@VStart = VerticalStart,
		@VStop = VerticalStop,
		@VStep = VerticalStep,
		@PColID = PageBreakColID,
		@PStart = PageBreakStart,
		@PStop = PageBreakStop,
		@PStep = PageBreakStep,
		@IType = IntersectionType,
		@IColID = IntersectionColID,
		@Percentage = Percentage,
		@PerPage = PercentageofPage,
		@Suppress = SuppressZeros,
		@Thousand = ThousandSeparators,
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
	FROM ASRSysCrossTab
	WHERE CrossTabID = @piReportID

	/* Check the current user can view the report. */
	exec spASRIntCurrentUserAccess 
		1, 
		@piReportID,
		@sAccess	OUTPUT

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 
	BEGIN
		SET @psErrorMsg = 'cross tab has been made hidden by another user.'
		RETURN
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 
	BEGIN
		SET @psErrorMsg = 'cross tab has been made read only by another user.'
		RETURN
	END

	IF @psAction = 'copy'
	BEGIN
		SET @psReportName = left('copy of ' + @psReportName, 50)
		SET @psReportOwner = @psCurrentUser
	END

	IF @piPicklistID > 0 
	BEGIN
		SELECT @psPicklistName = name,
			@sTempHidden = access
		FROM ASRSysPicklistName 
		WHERE picklistID = @piPicklistID

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfPicklistHidden = 1
		END
	END

	IF @piFilterID > 0 
	BEGIN
		SELECT @psFilterName = name,
			@sTempHidden = access
		FROM ASRSysExpressions 
		WHERE exprID = @piFilterID

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			SET @pfFilterHidden = 1
		END
	END

	IF @piOutputEmailAddr > 0
	BEGIN
		SELECT @psOutputEmailName = name,
			@sTempHidden = access
		FROM ASRSysEmailGroupName
		WHERE EmailGroupID = @piOutputEmailAddr
	END
	ELSE
	BEGIN
		SET @piOutputEmailAddr = 0	
		SET @psOutputEmailName = ''
	END

END
GO

