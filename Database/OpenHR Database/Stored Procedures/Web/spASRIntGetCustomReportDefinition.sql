CREATE PROCEDURE [dbo].[spASRIntGetCustomReportDefinition] (
	@piReportID 				integer, 
	@psCurrentUser				varchar(255),
	@psAction					varchar(255))
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iCount				integer,
			@sTempHidden		varchar(MAX),
			@sAccess			varchar(MAX),
			@sTempUsername		varchar(MAX),
			@fSysSecMgr			bit;

	DECLARE @psErrorMsg				varchar(MAX) = '',
		@psReportName				varchar(255) = '',
		@psReportOwner				varchar(255) = '',
		@psReportDesc				varchar(MAX) = '',
		@piBaseTableID				integer = 0,
		@pfAllRecords				bit			,
		@piPicklistID				integer = 0,
		@psPicklistName				varchar(255) = '',
		@pfPicklistHidden			bit			,
		@piFilterID					integer = 0,
		@psFilterName				varchar(255) = '',
		@pfFilterHidden				bit			,
		@piParent1TableID			integer = 0,
		@psParent1Name				varchar(255) = '',
		@piParent1FilterID			integer = 0,
		@psParent1FilterName		varchar(255) = '',
		@pfParent1FilterHidden		bit			,
		@piParent2TableID			integer = 0,
		@psParent2Name				varchar(255) = '',
		@piParent2FilterID			integer = 0,
		@psParent2FilterName		varchar(255) = '',
		@pfParent2FilterHidden		bit,
		@pfSummary					bit,
		@pfPrintFilterHeader		bit,
		@pfOutputPreview			bit,
		@piOutputFormat				integer = 0,
		@pfOutputScreen				bit,
		@pfOutputPrinter			bit,
		@psOutputPrinterName		varchar(MAX) = '',
		@pfOutputSave				bit,
		@piOutputSaveExisting		integer = 0,
		@pfOutputEmail				bit,
		@piOutputEmailAddr			integer = 0,
		@psOutputEmailName			varchar(MAX) = '',
		@psOutputEmailSubject		varchar(MAX) = '',
		@psOutputEmailAttachAs		varchar(MAX) = '',
		@psOutputFilename			varchar(MAX) = '',
		@piTimestamp				integer = 0,
		@pfParent1AllRecords		bit,
		@piParent1PicklistID		integer,
		@psParent1PicklistName		varchar(255) = '',
		@pfParent1PicklistHidden	bit,
		@pfParent2AllRecords		bit,
		@piParent2PicklistID		integer,
		@psParent2PicklistName		varchar(255) = '',
		@pfParent2PicklistHidden	bit,
		@psInfoMsg					varchar(MAX) = '',
		@pfIgnoreZeros				bit,
		@piCategoryID		integer;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
	
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

	-- Get's the category id associated with the custom report. Return 0 if not found
	SET @piCategoryID = 0
	SELECT @piCategoryID = ISNULL(categoryid,0)
		FROM [dbo].[tbsys_objectcategories]
		WHERE objectid = @piReportID AND objecttype = 2;


	-- Definition
	SELECT @psReportName AS name, @psReportDesc AS [Description], @piCategoryID As CategoryID, @piBaseTableID AS baseTableID, @psReportOwner AS [Owner],
		CASE WHEN @pfAllRecords = 1 THEN 0 ELSE CASE WHEN ISNULL(@piPicklistID, 0) > 0 THEN 1 ELSE 2 END END AS [SelectionType],
		@piPicklistID AS PicklistID, @piFilterID AS FilterID,
		@psPicklistName AS PicklistName, @psFilterName AS FilterName,
		CASE WHEN @piParent1FilterID > 0 THEN 2 ELSE CASE WHEN @piParent1PicklistID > 0 THEN 1 ELSE 0 END END AS [Parent1SelectionType],
		@piParent1TableID AS parent1ID, @psParent1Name AS Parent1Name, @piParent1FilterID AS parent1FilterID, @piParent1PicklistID AS Parent1PicklistID,
		@psParent1FilterName AS Parent1FilterName, @psParent1PicklistName AS Parent1PicklistName, @piParent2PicklistID AS Parent2PicklistID,
		CASE WHEN @piParent2FilterID > 0 THEN 2 ELSE CASE WHEN @piParent2PicklistID > 0 THEN 1 ELSE 0 END END AS [Parent2SelectionType],
		@piParent2TableID AS parent2ID, @psParent2Name AS Parent2Name, @piParent2FilterID AS parent2FilterID, 
		@psParent2FilterName AS Parent2FilterName, @psParent2PicklistName AS Parent2PicklistName,
		@pfSummary AS IsSummary,@pfPrintFilterHeader AS printFilterHeader,
		@pfOutputPreview AS IsPreview, @piOutputFormat AS [Format], @pfOutputScreen AS ToScreen, @pfOutputPrinter AS ToPrinter,
		@psOutputPrinterName AS PrinterName, @pfOutputSave AS SaveToFile, @piOutputSaveExisting AS SaveExisting,
		@pfOutputEmail AS SendToEmail, @piOutputEmailAddr AS EmailGroupID, @psOutputEmailName AS EmailGroupName,
		@psOutputEmailSubject AS EmailSubject, @psOutputEmailAttachAs AS EmailAttachmentName,
		@psOutputFilename AS [Filename], @piTimestamp AS [timestamp],
		@pfParent1AllRecords AS parent1AllRecords, @piParent1PicklistID AS parent1Picklist,
		@pfParent2AllRecords AS parent2AllRecords,@piParent2PicklistID AS parent2Picklist,
		@pfIgnoreZeros AS IgnoreZerosForAggregates,
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess],
		CASE WHEN @pfParent1PicklistHidden = 1 OR @pfParent1FilterHidden = 1 THEN 'HD' ELSE '' END AS [Parent1ViewAccess],
		CASE WHEN @pfParent2PicklistHidden = 1 OR @pfParent2FilterHidden = 1 THEN 'HD' ELSE '' END AS [Parent2ViewAccess];

	-- Get the definition columns
	SELECT 0 AS [AccessHidden],
		0 AS [IsExpression],
		ASRSysColumns.tableID,
		cd.colExprID AS [id],
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) AS [Name],
		cd.size AS [size],
		cd.dp AS [decimals],
		cd.heading AS Heading,
		ASRSysColumns.DataType,
		ISNULL(cd.avge, 0) AS IsAverage, ISNULL(cd.cnt, 0) AS IsCount, ISNULL(cd.tot, 0) AS IsTotal,
		ISNULL(cd.Hidden, 0) AS IsHidden,	ISNULL(cd.GroupWithNextColumn, 0) AS IsGroupWithNext,
		CASE cd.Repetition WHEN 1 THEN 1 ELSE 0 END AS IsRepeated,
		cd.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails cd
		INNER JOIN ASRSysColumns ON cd.colExprID = ASRSysColumns.columnId
		INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE cd.customReportID = @piReportID
		AND cd.type = 'C'
	UNION
	SELECT CASE 
			WHEN ASRSysExpressions.access = 'HD' THEN 1
			ELSE 0
		END,
		1 AS [IsExpression],
		ASRSysExpressions.tableID,
		cd.colExprID,
		ASRSysTables.TableName  + ' Calc: ' + replace(ASRSysExpressions.name, '_', ' ') AS [Heading],
		cd.size,
		cd.dp,
		cd.heading,
		0 AS [DataType],
		ISNULL(cd.avge, 0) AS IsAverage, ISNULL(cd.cnt, 0) AS IsCount, ISNULL(cd.tot, 0) AS IsTotal,
		ISNULL(cd.Hidden, 0) AS IsHidden,	ISNULL(cd.GroupWithNextColumn, 0) AS IsGroupWithNext,
		CASE cd.Repetition WHEN 1 THEN 1 ELSE 0 END AS IsRepeated,
		cd.sequence AS [sequence]
	FROM ASRSysCustomReportsDetails cd
		INNER JOIN ASRSysExpressions ON cd.colExprID = ASRSysExpressions.exprID
		INNER JOIN ASRSysTables ON ASRSysExpressions.tableID = ASRSysTables.tableID
	WHERE cd.customReportID = @piReportID
		AND cd.type <> 'C';

	-- Orders
	SELECT cd.colExprID AS [ID],
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) as [Name],
		cd.SortOrderSequence AS [Sequence],
		ISNULL(cd.boc, 0) AS [BreakOnChange],
		ISNULL(cd.poc, 0) AS [PageOnChange],
		ISNULL(cd.voc, 0) AS [ValueOnChange],
		ISNULL(cd.srv, 0) AS [SuppressRepeated],
		cd.sortOrder AS [Order],
		ASRSysTables.tableID
	FROM ASRSysCustomReportsDetails cd
	INNER JOIN ASRSysColumns ON cd.colExprID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
	WHERE cd.customReportID = @piReportID
		AND cd.type = 'C'
		AND cd.sortOrderSequence > 0
	ORDER BY cd.SortOrderSequence;

	-- Return the child table information
	SELECT  C.ChildTable AS [TableID],
		T.TableName AS [TableName],
		CASE WHEN (X.Access <> 'HD') OR (X.userName = system_user) THEN isnull(X.ExprID, 0) ELSE 0 END AS [FilterID],
		CASE WHEN (X.Access <> 'HD') OR (X.userName = system_user) THEN isnull(X.Name, '') ELSE '' END AS [FilterName],
		isnull(O.OrderID, 0) AS [OrderID],
	  ISNULL(O.Name, '') AS [OrderName],
	  C.ChildMaxRecords AS [Records],
		X.Access AS FilterViewAccess,
		CASE WHEN (X.Access = 'HD') AND (X.userName = system_user) THEN 'Y' ELSE 'N' END AS [FilterHidden],
		CASE WHEN isnull(O.OrderID, 0) <> isnull(C.ChildOrder,0) THEN 'Y' ELSE 'N' END AS [OrderDeleted],
		CASE WHEN isnull(X.ExprID, 0) <> isnull(C.ChildFilter,0) THEN 'Y' ELSE 'N' END AS [FilterDeleted],
		CASE WHEN (X.Access = 'HD') AND (X.userName <> system_user) THEN 'Y' ELSE 'N' END AS [FilterHiddenByOther]
	FROM [dbo].[ASRSysCustomReportsChildDetails] C 
	INNER JOIN [dbo].[ASRSysTables] T ON C.ChildTable = T.TableID 
		LEFT OUTER JOIN [dbo].[ASRSysExpressions] X ON C.ChildFilter = X.ExprID 
		LEFT OUTER JOIN [dbo].[ASRSysOrders] O ON C.ChildOrder = O.OrderID
	WHERE C.CustomReportID = @piReportID
	ORDER BY T.TableName;
	
END
