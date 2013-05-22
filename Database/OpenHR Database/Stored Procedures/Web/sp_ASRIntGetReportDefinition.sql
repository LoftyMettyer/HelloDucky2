CREATE PROCEDURE [dbo].[sp_ASRIntGetReportDefinition] (
	@piReportID 				integer, 
	@psCurrentUser				varchar(255),
	@psAction					varchar(255),
	@psErrorMsg					varchar(MAX)	OUTPUT,
	@psReportName				varchar(255)	OUTPUT,
	@psReportOwner				varchar(255)	OUTPUT,
	@psReportDesc				varchar(MAX)	OUTPUT,
	@piBaseTableID				integer			OUTPUT,
	@pfAllRecords				bit				OUTPUT,
	@piPicklistID				integer			OUTPUT,
	@psPicklistName				varchar(255)	OUTPUT,
	@pfPicklistHidden			bit				OUTPUT,
	@piFilterID					integer			OUTPUT,
	@psFilterName				varchar(255)	OUTPUT,
	@pfFilterHidden				bit				OUTPUT,
	@piParent1TableID			integer			OUTPUT,
	@psParent1Name				varchar(255)	OUTPUT,
	@piParent1FilterID			integer			OUTPUT,
	@psParent1FilterName		varchar(255)	OUTPUT,
	@pfParent1FilterHidden		bit				OUTPUT,
	@piParent2TableID			integer			OUTPUT,
	@psParent2Name				varchar(255)	OUTPUT,
	@piParent2FilterID			integer			OUTPUT,
	@psParent2FilterName		varchar(255)	OUTPUT,
	@pfParent2FilterHidden		bit				OUTPUT,
	@pfSummary					bit				OUTPUT,
	@pfPrintFilterHeader		bit				OUTPUT,
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
 	@piTimestamp				integer			OUTPUT,
	@pfParent1AllRecords		bit				OUTPUT,
	@piParent1PicklistID		integer			OUTPUT,
	@psParent1PicklistName		varchar(255)	OUTPUT,
	@pfParent1PicklistHidden	bit				OUTPUT,
	@pfParent2AllRecords		bit				OUTPUT,
	@piParent2PicklistID		integer			OUTPUT,
	@psParent2PicklistName		varchar(255)	OUTPUT,
	@pfParent2PicklistHidden	bit				OUTPUT,
	@psInfoMsg					varchar(MAX)	OUTPUT,
	@pfIgnoreZeros				bit				OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iCount			integer,
			@sTempHidden	varchar(8000),
			@sAccess		varchar(8000),
			@sTempUsername	varchar(8000),
			@fSysSecMgr		bit;

	SET @psErrorMsg = '';
	SET @psPicklistName = '';
	SET @pfPicklistHidden = 0;
	SET @psFilterName = '';
	SET @pfFilterHidden = 0;
	SET @psParent1Name = '';
	SET @psParent1FilterName = '';
	SET @pfParent1FilterHidden = 0;
	SET @psParent2Name = '';
	SET @psParent2FilterName = '';
	SET @pfParent2FilterHidden = 0;
	SET @psParent1PicklistName = '';
	SET @pfParent1PicklistHidden = 0;
	SET @psParent2PicklistName = '';
	SET @pfParent2PicklistHidden = 0;
	SET @psInfoMsg = '';

	exec [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
	
	/* Check the report exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysCustomReportsName 
	WHERE ID = @piReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'report has been deleted by another user.';
		RETURN;
	END

	SELECT @psReportName = name,
		@psReportDesc	 = description,
		@piBaseTableID = baseTable,
		@pfAllRecords = allRecords,
		@piPicklistID = picklist,
		@piFilterID = filter,
		@piParent1TableID = parent1Table,
		@piParent1FilterID = parent1Filter,
		@piParent2TableID = parent2Table,
		@piParent2FilterID = parent2Filter,
		@pfSummary = summary,
		@pfPrintFilterHeader = printFilterHeader,
		@psReportOwner = userName,
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
		@piTimestamp = convert(integer, timestamp),
		@pfParent1AllRecords = parent1AllRecords,
		@piParent1PicklistID = parent1Picklist,
		@pfParent2AllRecords = parent2AllRecords,
		@piParent2PicklistID = parent2Picklist,
		@pfIgnoreZeros = IgnoreZeros
	FROM [dbo].[ASRSysCustomReportsName]
	WHERE ID = @piReportID;

	/* Check the current user can view the report. */
	exec [dbo].[spASRIntCurrentUserAccess]
		2, 
		@piReportID,
		@sAccess OUTPUT;

	IF @fSysSecMgr = 0 
	BEGIN
		IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 
		BEGIN
			SET @psErrorMsg = 'report has been made hidden by another user.';
			RETURN;
		END

		IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 
		BEGIN
			SET @psErrorMsg = 'report has been made read only by another user.';
			RETURN;
		END
	END
	
	/* Check the report has details. */
	SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysCustomReportsDetails]
		WHERE ASRSysCustomReportsDetails.customReportID = @piReportID;

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'report contains no details.';
		RETURN;
	END

	/* Check the report has sort order details. */
	SELECT @iCount = COUNT(*)
	FROM [dbo].[ASRSysCustomReportsDetails]
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type = 'C'
		AND ASRSysCustomReportsDetails.sortOrderSequence > 0

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'report contains no sort order details.';
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
			@sTempHidden = access,
			@sTempUsername = username
		FROM [dbo].[ASRSysPicklistName]
		WHERE picklistID = @piPicklistID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			IF UPPER(@sTempUsername) = UPPER(system_user)
			BEGIN
				SET @pfPicklistHidden = 1;
			END
			ELSE
			BEGIN
				/* Picklist is hidden by another user. Remove it from the definition. */
				IF @fSysSecMgr = 0
				BEGIN
					SET @piPicklistID = 0;
					SET @psPicklistName = '';
					SET @pfPicklistHidden = 0;

					SET @psInfoMsg = @psInfoMsg +
					CASE
						WHEN LEN(@psInfoMsg) > 0 THEN char(10)
						ELSE ''
					END + 'The base table picklist will be removed from this definition as it has been made hidden by another user.';
				END
			END
		END
	END

	IF @piFilterID > 0 
	BEGIN
		SELECT @psFilterName = name,
			@sTempHidden = access,
			@sTempUsername = username
		FROM [dbo].[ASRSysExpressions]
		WHERE exprID = @piFilterID;

		IF UPPER(@sTempHidden) = 'HD'
		BEGIN
			IF UPPER(@sTempUsername) = UPPER(system_user)
			BEGIN
				SET @pfFilterHidden = 1;
			END
			ELSE
			BEGIN
				/* Filter is hidden by another user. Remove it from the definition. */
				IF @fSysSecMgr = 0
				BEGIN
					SET @piFilterID = 0;
					SET @psFilterName = '';
					SET @pfFilterHidden = 0;

					SET @psInfoMsg = @psInfoMsg +
					CASE
						WHEN LEN(@psInfoMsg) > 0 THEN char(10)
						ELSE ''
					END + 'The base table filter will be removed from this definition as it has been made hidden by another user.';
				END
			END
		END
	END

	IF @piParent1TableID > 0 
	BEGIN
		SELECT @psParent1Name = tableName
		FROM [dbo].[ASRSysTables]
		WHERE tableID = @piParent1TableID;

		IF @piParent1PicklistID > 0 
		BEGIN
			SELECT @psParent1PicklistName = name,
				@sTempHidden = access,
				@sTempUsername = username
			FROM [dbo].[ASRSysPicklistName]
			WHERE picklistID = @piParent1PicklistID;
	
			IF UPPER(@sTempHidden) = 'HD'
			BEGIN
				IF UPPER(@sTempUsername) = UPPER(system_user)
				BEGIN
					SET @pfParent1PicklistHidden = 1;
				END
				ELSE
				BEGIN
					/* Picklist is hidden by another user. Remove it from the definition. */
					IF @fSysSecMgr = 0
					BEGIN
						SET @piParent1PicklistID = 0;
						SET @psParent1PicklistName = '';
						SET @pfParent1PicklistHidden = 0;

						SET @psInfoMsg = @psInfoMsg +
						CASE
							WHEN LEN(@psInfoMsg) > 0 THEN char(10)
							ELSE ''
						END + 'The ''' + @psParent1Name + ''' table picklist will be removed from this definition as it has been made hidden by another user.';
					END
				END
			END
		END

		IF @piParent1FilterID > 0 
		BEGIN
			SELECT @psParent1FilterName = name,
				@sTempHidden = access,
				@sTempUsername = username
			FROM [dbo].[ASRSysExpressions]
			WHERE exprID = @piParent1FilterID;

			IF UPPER(@sTempHidden) = 'HD'
			BEGIN
				IF UPPER(@sTempUsername) = UPPER(system_user)
				BEGIN
					SET @pfParent1FilterHidden = 1;
				END
				ELSE
				BEGIN
					/* Filter is hidden by another user. Remove it from the definition. */
					IF @fSysSecMgr = 0
					BEGIN
						SET @piParent1FilterID = 0;
						SET @psParent1FilterName = '';
						SET @pfParent1FilterHidden = 0;

						SET @psInfoMsg = @psInfoMsg +
						CASE
							WHEN LEN(@psInfoMsg) > 0 THEN char(10)
							ELSE ''
						END + 'The ''' + @psParent1Name + ''' table filter will be removed from this definition as it has been made hidden by another user.';
					END
				END
			END
		END	
	END

	IF @piParent2TableID > 0 
	BEGIN
		SELECT @psParent2Name = tableName 
		FROM [dbo].[ASRSysTables]
		WHERE tableID = @piParent2TableID;

		IF @piParent2PicklistID > 0 
		BEGIN
			SELECT @psParent2PicklistName = name,
				@sTempHidden = access,
				@sTempUsername = username
			FROM [dbo].[ASRSysPicklistName]
			WHERE picklistID = @piParent2PicklistID;
	
			IF UPPER(@sTempHidden) = 'HD'
			BEGIN
				IF UPPER(@sTempUsername) = UPPER(system_user)
				BEGIN
					SET @pfParent2PicklistHidden = 1;
				END
				ELSE
				BEGIN
					/* Picklist is hidden by another user. Remove it from the definition. */
					IF @fSysSecMgr = 0
					BEGIN
						SET @piParent2PicklistID = 0;
						SET @psParent2PicklistName = '';
						SET @pfParent2PicklistHidden = 0;

						SET @psInfoMsg = @psInfoMsg +
						CASE
							WHEN LEN(@psInfoMsg) > 0 THEN char(10)
							ELSE ''
						END + 'The ''' + @psParent2Name + ''' table picklist will be removed from this definition as it has been made hidden by another user.';
					END
				END
			END
		END

		IF @piParent2FilterID > 0 
		BEGIN
			SELECT @psParent2FilterName = name,
				@sTempHidden = access,
				@sTempUsername = username
			FROM [dbo].[ASRSysExpressions]
			WHERE exprID = @piParent2FilterID;

			IF UPPER(@sTempHidden) = 'HD'
			BEGIN
				IF UPPER(@sTempUsername) = UPPER(system_user)
				BEGIN
					SET @pfParent2FilterHidden = 1;
				END
				ELSE
				BEGIN
					/* Filter is hidden by another user. Remove it from the definition. */
					IF @fSysSecMgr = 0
					BEGIN
						SET @piParent2FilterID = 0;
						SET @psParent2FilterName = '';
						SET @pfParent2FilterHidden = 0;

						SET @psInfoMsg = @psInfoMsg +
						CASE
							WHEN LEN(@psInfoMsg) > 0 THEN char(10)
							ELSE ''
						END + 'The ''' + @psParent2Name + ''' table filter will be removed from this definition as it has been made hidden by another user.';
					END
				END
			END
		END	
	END

	IF @piOutputEmailAddr > 0
	BEGIN
		SELECT @psOutputEmailName = name,
			@sTempHidden = access
		FROM [dbo].[ASRSysEmailGroupName]
		WHERE EmailGroupID = @piOutputEmailAddr;
	END
	ELSE
	BEGIN
		SET @piOutputEmailAddr = 0;	
		SET @psOutputEmailName = '';
	END

	/* Get the definition recordset. */
	SELECT 'COLUMN' AS [definitionType],
		'N' AS [hidden],
		convert(varchar(255), ASRSysCustomReportsDetails.type) + char(9) +
		convert(varchar(255), ASRSysColumns.tableID) + char(9) +
		convert(varchar(255), ASRSysCustomReportsDetails.colExprID) + char(9) +
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) + char(9) +
		convert(varchar(100), ASRSysCustomReportsDetails.size) + char(9) +
		convert(varchar(100), ASRSysCustomReportsDetails.dp) + char(9) +
		'N' + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.isNumeric = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		convert(varchar(8000), ASRSysCustomReportsDetails.heading) + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.avge = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE
			WHEN ASRSysCustomReportsDetails.cnt = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.tot = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.Hidden = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.GroupWithNextColumn = 0 THEN '0' 
			ELSE '1' 
		END AS [definitionString],
		ASRSysCustomReportsDetails.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails 
	INNER JOIN ASRSysColumns ON ASRSysCustomReportsDetails.colExprID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type = 'C'
	UNION
	SELECT 'COLUMN' AS [definitionType],
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y'
			ELSE 'N'
		END AS [hidden],
		convert(varchar(255), ASRSysCustomReportsDetails.type) + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysCustomReportsDetails.colExprID) + char(9) +
		convert(varchar(MAX), '<' + ASRSysTables.TableName  + ' Calc> ' + replace(ASRSysExpressions.name, '_', ' ')) + char(9) +
		convert(varchar(100), ASRSysCustomReportsDetails.size) + char(9) +
		convert(varchar(100), ASRSysCustomReportsDetails.dp) + char(9) +
		CASE
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y'
			ELSE 'N'
		END + char(9) +
		CASE
			WHEN ASRSysCustomReportsDetails.isNumeric = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		convert(varchar(8000), ASRSysCustomReportsDetails.heading) + char(9) +

		CASE
			WHEN ASRSysCustomReportsDetails.avge = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			when ASRSysCustomReportsDetails.cnt = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.tot = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.Hidden = 0 THEN '0' 
			ELSE '1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.GroupWithNextColumn = 0 THEN '0' 
			ELSE '1' 
		END AS [definitionString],
		ASRSysCustomReportsDetails.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails
		INNER JOIN ASRSysExpressions ON ASRSysCustomReportsDetails.colExprID = ASRSysExpressions.exprID
		INNER JOIN ASRSysTables ON ASRSysExpressions.tableID = ASRSysTables.tableID
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type <> 'C'
	UNION
	SELECT 'ORDER' AS [definitionType],
		'N' AS [hidden],
		convert(varchar(255), ASRSysCustomReportsDetails.colExprID) + char(9) +
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) + char(9) +
		convert(varchar(255), ASRSysCustomReportsDetails.sortOrder) + char(9) +
		CASE
			WHEN ASRSysCustomReportsDetails.boc = 0 THEN '' 
			ELSE '-1' 
		END + char(9) +
		CASE
			WHEN ASRSysCustomReportsDetails.poc = 0 THEN '' 
			ELSE '-1' 
		END + char(9) +
		CASE
			WHEN ASRSysCustomReportsDetails.voc = 0 THEN '' 
			ELSE '-1' 
		END + char(9) +
		CASE 
			WHEN ASRSysCustomReportsDetails.srv = 0 THEN '' 
			ELSE '-1' 
		END  + char(9) +
		convert(varchar(255), ASRSysTables.tableID) AS [definitionString],
		ASRSysCustomReportsDetails.sortOrderSequence AS [sequence]
	FROM ASRSysCustomReportsDetails
	INNER JOIN ASRSysColumns ON ASRSysCustomReportsDetails.colExprID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type = 'C'
		AND ASRSysCustomReportsDetails.sortOrderSequence > 0
	UNION
	SELECT 'REPETITION' AS [definitionType],
		'N' AS [hidden],
		'C' + convert(varchar(255), ASRSysCustomReportsDetails.colExprID) + char(9) +
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) + char(9) +
		convert(varchar(MAX), ASRSysCustomReportsDetails.repetition)  + char(9) +
		convert(varchar(255), ASRSysTables.tableID) + char(9) +
		convert(varchar(255), ASRSysCustomReportsDetails.Hidden) AS [definitionString],
		ASRSysCustomReportsDetails.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails
		INNER JOIN ASRSysColumns ON ASRSysCustomReportsDetails.colExprID = ASRSysColumns.columnID
		INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type = 'C'
		AND ASRSysCustomReportsDetails.repetition >= 0
	UNION
	SELECT 'REPETITION' AS [definitionType],
		CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 'Y'
			ELSE 'N'
		END AS [hidden],
		'E' + convert(varchar(8000), ASRSysCustomReportsDetails.colExprID) + char(9) +
		'<' + ASRSysTables.TableName + ' Calc> ' + convert(varchar(MAX), ASRSysExpressions.Name) + char(9) +
		convert(varchar(100), ASRSysCustomReportsDetails.repetition)  + char(9) +
		convert(varchar(255), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(255), ASRSysCustomReportsDetails.Hidden) AS [definitionString],
		ASRSysCustomReportsDetails.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails
		INNER JOIN ASRSysExpressions ON ASRSysCustomReportsDetails.colExprID = ASRSysExpressions.ExprID
		INNER JOIN ASRSysTables ON ASRSysExpressions.tableID = ASRSysTables.tableID
	WHERE ASRSysCustomReportsDetails.customReportID = @piReportID
		AND ASRSysCustomReportsDetails.type = 'E'
		AND ASRSysCustomReportsDetails.repetition >= 0
	ORDER BY [sequence] ASC;
	
END