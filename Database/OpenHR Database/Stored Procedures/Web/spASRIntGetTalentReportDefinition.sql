﻿CREATE PROCEDURE [dbo].[spASRIntGetTalentReportDefinition] (	
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

	DECLARE @psErrorMsg				varchar(MAX) = '',	
			@psPicklistName			varchar(255) = '',
			@psMatchPicklistName	varchar(255) = '',
			@pfPicklistHidden		bit = 0,
			@pfMatchPicklistHidden	bit = 0,
			@psFilterName			varchar(255) = '',
			@psMatchFilterName		varchar(255) = '',
			@pfFilterHidden			bit = 0,
			@pfMatchFilterHidden	bit = 0,
			@psWarningMsg			varchar(255) = '',
			@psReportOwner			varchar(255),
			@psReportName			varchar(255),
			@piPicklistID			integer = 0,
			@piFilterID				integer = 0,
			@piMatchPicklistID		integer = 0,
			@piMatchFilterID		integer = 0,
			@piCategoryID			integer;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;

	/* Check the mail merge exists. */		
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysTalentReports]		
	WHERE ID = @piReportID;		

	IF @iCount = 0		
		SET @psErrorMsg = 'talent report has been deleted by another user.';		

	SELECT @psReportOwner = [username], @psReportName = [name]
			, @piPicklistID = BasePicklistID, @piFilterID = BaseFilterID
			, @piMatchPicklistID = MatchPicklistID, @piMatchFilterID = MatchFilterID
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

	IF @piMatchPicklistID > 0 		
	BEGIN		
		SELECT @psMatchPicklistName = name, @sTempHidden = access		
		FROM [dbo].[ASRSysPicklistName]		
		WHERE picklistID = @piMatchPicklistID;	
		IF UPPER(@sTempHidden) = 'HD'
			SET @pfMatchPicklistHidden = 1;	
	END			

	IF @piMatchFilterID > 0 		
	BEGIN		
		SELECT @psMatchFilterName = name, @sTempHidden = access		
		FROM [dbo].[ASRSysExpressions]		
		WHERE exprID = @piMatchFilterID;		
		IF UPPER(@sTempHidden) = 'HD'
			SET @pfMatchFilterHidden = 1;
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
		@psMatchPicklistName AS PersonPicklistName,
		m.BaseFilterID AS FilterID,
		@psFilterName AS FilterName,
		@psMatchFilterName AS PersonFilterName,
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
		ISNULL(m.IncludeUnmatched, 0) AS IncludeUnmatched,
		ISNULL(m.MinimumScore, 0) AS MinimumScore,
	  m.OutputEmail AS [SendToEmail],
		m.outputformat AS [Format],		
		m.outputsave AS [SaveToFile],		
		m.outputfilename AS [Filename],		
		m.emailAddrID AS [EmailGroupID],	
		(SELECT Name FROM [dbo].[ASRSysEmailGroupName] WHERE EmailGroupID = m.emailAddrID) AS EmailGroupName,	
		m.emailSubject,		
		m.EmailAttachmentName,	
		m.outputscreen AS [DisplayOutputOnScreen],		
		CONVERT(integer, m.[timestamp]) AS [Timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess],
		CASE WHEN @pfMatchPicklistHidden = 1 OR @pfMatchFilterHidden = 1 THEN 'HD' ELSE '' END AS [MatchViewAccess]
	FROM [dbo].ASRSysTalentReports m
	WHERE m.ID = @piReportID;		

	
	-- Columns
	SELECT r.ColExprID AS [ID],
		0 AS [IsExpression],
		0 AS [accesshidden],
		c.tableID,
		t.tableName + '.' + c.columnName AS [name],		
		c.DataType,
		r.ColSize AS [Size],
		r.ColDecs AS [Decimals],
		r.ColHeading AS Heading,
		0 AS IsAverage,
		0 AS IsCount,
		0 AS IsTotal,
		0 AS IsHidden,
		0 AS IsGroupWithNext,
		0 AS IsRepeated,
		r.ColSequence AS [sequence]
	FROM ASRSysTalentReportDetails r	
	INNER JOIN ASRSysColumns c ON r.ColExprID = c.columnId		
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	WHERE r.TalentReportID = @piReportID;

	-- Orders
	SELECT r.ColExprID AS [id],
		convert(varchar(MAX), t.tableName + '.' + c.columnName) AS [name],
		r.SortOrderDirection AS [order],
		t.tableID,
		r.sortOrderSeq AS [sequence]
	FROM ASRSysTalentReportDetails r		
	INNER JOIN ASRSysColumns c ON r.ColExprID = c.columnId		
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	WHERE r.TalentReportID = @piReportID		
		AND r.sortOrderSeq > 0		
	ORDER BY [sequence] ASC;

END