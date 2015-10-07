CREATE PROCEDURE [dbo].[spASRIntGetMailMergeDefinition] (	
	@piReportID 			integer, 	
	@psCurrentUser			varchar(255),		
	@psAction				varchar(255)
)		
AS		
BEGIN		
	SET NOCOUNT ON;

	DECLARE	@iCount		integer,		
			@sTempHidden	varchar(MAX),		
			@sAccess 		varchar(MAX),		
			@fSysSecMgr		bit;		

	DECLARE @psErrorMsg			varchar(MAX) = '',	
			@psPicklistName		varchar(255) = '',
			@pfPicklistHidden	bit = 0,
			@psFilterName		varchar(255) = '',
			@pfFilterHidden		bit = 0,
			@psWarningMsg		varchar(255) = '',
			@psReportOwner		varchar(255),
			@psReportName		varchar(255),
			@piPicklistID		integer = 0,
			@piFilterID			integer = 0;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;

	/* Check the mail merge exists. */		
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysMailMergeName]		
	WHERE MailMergeID = @piReportID;		

	IF @iCount = 0		
		SET @psErrorMsg = 'mail merge has been deleted by another user.';		

	SELECT @psReportOwner = [username], @psReportName = [name]
			, @piPicklistID = picklistID, @piFilterID = FilterID
		FROM [dbo].[ASRSysMailMergeName]		
		WHERE MailMergeID = @piReportID;
	
	-- Check the current user can view the report.
	EXEC [dbo].[spASRIntCurrentUserAccess] 9, @piReportID, @sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'mail merge has been made hidden by another user.';		

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'mail merge has been made read only by another user.';		

	-- Check the report has details.
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysMailMergeColumns]		
	WHERE MailMergeID = @piReportID;		
	IF @iCount = 0		
		SET @psErrorMsg = 'mail merge contains no details.';		

	-- Check the report has sort order details.
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysMailMergeColumns]		
	WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
		AND ASRSysMailMergeColumns.sortOrderSequence > 0;		
	IF @iCount = 0		
		SET @psErrorMsg = 'mail merge contains no sort order details.';		

	IF @psAction = 'copy' 		
	BEGIN		
		SET @psReportName = left('copy of ' + @psReportName, 50);		
		SET @psReportOwner = @psCurrentUser;		
	END		

	IF @piPicklistID > 0 		
	BEGIN		
		SELECT @psPicklistName = name, @sTempHidden = access		
		FROM [dbo].[ASRSysPicklistName]		
		WHERE picklistID = @piPicklistID;		
		IF UPPER(@sTempHidden) = 'HD'		
			SET @pfPicklistHidden = 1;		

	END		
	IF @piFilterID > 0 		
	BEGIN		
		SELECT @psFilterName = name, @sTempHidden = access		
		FROM [dbo].[ASRSysExpressions]		
		WHERE exprID = @piFilterID;		
		IF UPPER(@sTempHidden) = 'HD'		
			SET @pfFilterHidden = 1;		

	END

	-- Definition
	SELECT @psReportName AS [Name], m.[description], @psReportOwner AS [owner],		
		m.tableID AS BaseTableID,		
		m.selection AS SelectionType,
		m.picklistID,	
		@psPicklistName AS PicklistName,
		m.FilterID,
		@psFilterName AS FilterName,
		m.outputformat AS [Format],		
		m.outputsave AS [SaveToFile],		
		m.outputfilename AS [Filename],		
		m.emailAddrID AS [EmailGroupID],		
		m.emailSubject,		
		ISNULL(m.UploadTemplateName, '') AS [UploadTemplateName],
		m.UploadTemplate,
		m.outputscreen AS [DisplayOutputOnScreen],		
		m.emailasattachment AS [EmailAsAttachment],		
		ISNULL(m.emailattachmentname,'') AS [EmailAttachmentName],		
		m.suppressblanks AS SuppressBlankLines,		
		m.PauseBeforeMerge,		
		m.outputprinter AS [SendToPrinter],		
		m.outputprintername AS [PrinterName],		
		m.documentmapid,		
		m.manualdocmanheader,
		m.PromptStart AS PauseBeforeMerge,
		CONVERT(integer, m.[timestamp]) AS [Timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess]
	FROM [dbo].[ASRSysMailMergeName] m
	WHERE m.MailMergeID = @piReportID;		

	-- Columns
	SELECT ASRSysMailMergeColumns.ColumnID AS [ID],
		0 AS [IsExpression],
		0 AS [accesshidden],
		ASRSysColumns.tableID,
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName AS [name],
		ASRSysColumns.columnName AS [heading], 
		ASRSysColumns.DataType,
		ASRSysMailMergeColumns.size,
		ASRSysMailMergeColumns.decimals,
		'' AS Heading,
		0 AS IsAverage,
		0 AS IsCount,
		0 AS IsTotal,
		0 AS IsHidden,
		0 AS IsGroupWithNext,
		0 AS IsRepeated,
		ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
	FROM ASRSysMailMergeColumns		
	INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnID = ASRSysColumns.columnId		
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID		
	WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
		AND ASRSysMailMergeColumns.type = 'C'
	UNION
	SELECT ASRSysMailMergeColumns.columnID AS [ID],
		1 AS [IsExpression],
		CASE WHEN ASRSysExpressions.access = 'HD' THEN 1 ELSE 0 END AS [accesshidden],		
		ASRSysExpressions.tableID,
		ASRSysExpressions.name AS [name],
		convert(varchar(MAX), replace(ASRSysExpressions.name, '_', ' ')) AS [heading],
		0 AS DataType,
		ASRSysMailMergeColumns.size,
		ASRSysMailMergeColumns.decimals,
		'' AS Heading,
		0 AS IsAverage,
		0 AS IsCount,
		0 AS IsTotal,
		0 AS IsHidden,
		0 AS IsGroupWithNext,
		0 AS IsRepeated,
		ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
	FROM ASRSysMailMergeColumns		
	INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnID = ASRSysExpressions.exprID		
	WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
		AND ASRSysMailMergeColumns.type <> 'C'		
		AND ((ASRSysExpressions.username = @psReportOwner)	OR (ASRSysExpressions.access <> 'HD'))		

	-- Orders
	SELECT ASRSysMailMergeColumns.columnID AS [id],
		convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) AS [name],
		ASRSysMailMergeColumns.sortOrder AS [order],
		ASRSysTables.tableID,
		ASRSysMailMergeColumns.sortOrderSequence AS [sequence]
	FROM ASRSysMailMergeColumns		
	INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnid = ASRSysColumns.columnId		
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID		
	WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
		AND ASRSysMailMergeColumns.sortOrderSequence > 0		
	ORDER BY ASRSysMailMergeColumns.type, [sequence] ASC;

	IF @fSysSecMgr = 0 		
	BEGIN		
		SELECT @iCount = COUNT(ASRSysMailMergeColumns.ID)		
		FROM [dbo].[ASRSysMailMergeColumns]		
		INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnID = ASRSysExpressions.exprID		
		WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
			AND ASRSysMailMergeColumns.type <> 'C'		
			and ((ASRSysExpressions.username <> @psReportOwner) and (ASRSysExpressions.access = 'HD'));		
							
		IF @iCount > 0 		
		BEGIN		
			IF @iCount = 1		
			BEGIN		
				SET @psWarningMsg = 'A calculation used in this definition has been made hidden by another user. It will be removed from the definition';		
			END		
			ELSE		
			BEGIN		
				SET @psWarningMsg = 'Some calculations used in this definition have been made hidden by another user. They will be removed from the definition';		
			END		
		END		
	END		
END
