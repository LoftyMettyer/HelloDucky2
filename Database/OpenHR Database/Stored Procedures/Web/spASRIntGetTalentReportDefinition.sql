CREATE PROCEDURE [dbo].[spASRIntGetTalentReportDefinition] (	
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
			@piFilterID			integer = 0,
			@piCategoryID		integer;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;

	/* Check the mail merge exists. */		
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysTalentReports]		
	WHERE ID = @piReportID;		

	IF @iCount = 0		
		SET @psErrorMsg = 'talent report has been deleted by another user.';		

	SELECT @psReportOwner = [username], @psReportName = [name]
			, @piPicklistID = BasePicklistID, @piFilterID = BaseFilterID
		FROM [dbo].[ASRSysTalentReports]		
		WHERE ID = @piReportID;

	-- Check the current user can view the report.
	EXEC [dbo].[spASRIntCurrentUserAccess] 9, @piReportID, @sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'talent report has been made hidden by another user.';		

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'talent report has been made read only by another user.';		

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

	-- Get's the category id associated with the mail merge utility. Return 0 if not found
	SET @piCategoryID = 0;
	SELECT @piCategoryID = ISNULL(categoryid,0)
		FROM [dbo].[tbsys_objectcategories]
		WHERE objectid = @piReportID AND objecttype = 38;

	-- Definition
	SELECT @psReportName AS [Name], @piCategoryID As CategoryID, m.[description], @psReportOwner AS [owner],		
		m.BaseTableID AS BaseTableID,		
		m.BaseSelection AS SelectionType,
		m.BasePicklistID AS PicklistID,	
		@psPicklistName AS PicklistName,
		m.BaseFilterID AS FilterID,
		@psFilterName AS FilterName,
	  ISNULL(m.BaseChildTableID, 0) AS BaseChildTableID,
	  ISNULL(m.BaseChildColumnID, 0) AS BaseChildColumnID,
		ISNULL(m.BasePreferredRatingColumnID, 0) AS BasePreferredRatingColumnID,
		ISNULL(m.BaseMinimumRatingColumnID, 0) AS BaseMinimumRatingColumnID,
    ISNULL(m.MatchTableID, 0) AS MatchTableID,
	  ISNULL(m.MatchSelection, 0) AS MatchSelection,
	  ISNULL(m.MatchPicklistID, 0) AS MatchPicklistID,
	  ISNULL(m.MatchFilterID, 0) AS MatchFilterID,
	  ISNULL(m.MatchChildTableID, 0) AS MatchChildTableID,
	  ISNULL(m.MatchChildColumnID, 0) AS MatchChildColumnID,
	  ISNULL(m.MatchChildRatingColumnID, 0) AS MatchChildRatingColumnID,
	  ISNULL(m.MatchAgainstType, 0) AS MatchAgainstType,
		m.outputformat AS [Format],		
		m.outputsave AS [SaveToFile],		
		m.outputfilename AS [Filename],		
		m.emailAddrID AS [EmailGroupID],		
		m.emailSubject,		
		m.outputscreen AS [DisplayOutputOnScreen],		
		CONVERT(integer, m.[timestamp]) AS [Timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess]
	FROM [dbo].ASRSysTalentReports m
	WHERE m.ID = @piReportID;		

	
	-- Columns
	SELECT r.ColumnID AS [ID],
		0 AS [IsExpression],
		0 AS [accesshidden],
		c.tableID,
		t.tableName + '.' + c.columnName AS [name],
		c.columnName AS [heading], 
		c.DataType,
		r.size,
		r.decimals,
		'' AS Heading,
		0 AS IsAverage,
		0 AS IsCount,
		0 AS IsTotal,
		0 AS IsHidden,
		0 AS IsGroupWithNext,
		0 AS IsRepeated,
		r.SortOrderSequence AS [sequence]
	FROM ASRSysTalentReportColumns r	
	INNER JOIN ASRSysColumns c ON r.columnID = c.columnId		
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	WHERE r.TalentReportID = @piReportID;

	-- Orders
	SELECT r.columnID AS [id],
		convert(varchar(MAX), t.tableName + '.' + c.columnName) AS [name],
		r.sortOrder AS [order],
		t.tableID,
		r.sortOrderSequence AS [sequence]
	FROM ASRSysTalentReportColumns r		
	INNER JOIN ASRSysColumns c ON r.columnid = c.columnId		
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	WHERE r.TalentReportID = @piReportID		
		AND r.sortOrderSequence > 0		
	ORDER BY [sequence] ASC;

END
