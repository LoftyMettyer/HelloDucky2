CREATE PROCEDURE [dbo].[spASRIntGetOrganisationReportDefinition] (	
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
			@psBaseViewName			varchar(255) = '',			
			@pfBaseViewHidden		   bit = 0,			
			@psWarningMsg			   varchar(255) = '',
			@psReportOwner			   varchar(255),
			@psReportName			   varchar(255),
			@piBaseViewID			   integer = 0,			
			@piCategoryID			   integer;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;

	/* Check the organisation report is exists. */		
	SELECT @iCount = COUNT(*)		
	FROM [dbo].[ASRSysOrganisationReport]		
	WHERE ID = @piReportID;		

	IF @iCount = 0		
		SET @psErrorMsg = 'organisation report has been deleted by another user.';		

	SELECT  @psReportOwner = [username], 
			@psReportName = [name]
		  , @piBaseViewID = BaseViewID
	FROM [dbo].[ASRSysOrganisationReport]		
	WHERE ID = @piReportID;

	-- Check the current user can view the report.
	EXEC [dbo].[spASRIntCurrentUserAccess] 39, @piReportID, @sAccess OUTPUT;

	IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'organisation report has been made hidden by another user.';		

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 		
		SET @psErrorMsg = 'organisation report has been made read only by another user.';		

	IF @psAction = 'copy' 		
	BEGIN		
		SET @psReportName = left('copy of ' + @psReportName, 50);		
		SET @psReportOwner = @psCurrentUser;		
	END		
	

	-- Get's the category id associated with the mail merge utility. Return 0 if not found
	SET @piCategoryID = 0;
	SELECT @piCategoryID = ISNULL(categoryid,0)
	FROM [dbo].[tbsys_objectcategories]
	WHERE objectid = @piReportID AND objecttype = 39;

	-- Definition
	SELECT  @psReportName AS [Name],
			  @piBaseViewID AS [BaseViewID],
			  @piCategoryID AS [CategoryID], 
			  m.[description], 
			  @psReportOwner AS [owner],					
			 CONVERT(integer, m.[timestamp]) AS [Timestamp]
	FROM [dbo].ASRSysOrganisationReport m
	WHERE m.ID = @piReportID;

	---- Filter
	SELECT r.OrganisationID    AS [OrganisationID],		
		    r.FieldID           AS [FieldID],
		    r.Operator          AS [Operator],
		    r.Value             AS [Value],
		    c.ColumnName        AS [FieldName],
		    c.datatype          AS [FieldDataType]		
	FROM ASRSysOrganisationReportFilters r
	INNER JOIN ASRSysColumns c ON r.FieldID = c.columnId
	WHERE r.OrganisationID	= @piReportID;	

	-- Columns
	SELECT   r.ColumnID	AS [ID],					
			   r.ViewID	   AS [ViewID],		
			   c.tableID,
            c.columnName AS [Heading],
			   CASE
					   WHEN r.ViewID > 0 THEN	v.ViewName + '.' + c.columnName
					   WHEN r.ViewID = 0 THEN	t.tableName + '.' + c.columnName			
			   END			AS [name],					
			   c.DataType,
			   r.Prefix	   AS [Prefix],
			   r.Suffix	   AS [Suffix],
			   r.FontSize	AS [FontSize],
			   r.Decimals	AS [Decimals],
			   r.Height	   AS [Height],
			   r.ConcatenateWithNext AS [ConcatenateWithNext]				
	FROM ASRSysOrganisationColumns r	
	INNER JOIN ASRSysColumns c ON r.ColumnID = c.columnId		
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID		
	LEFT JOIN ASRSysViews v ON r.ViewID = v.ViewID
	WHERE r.OrganisationID = @piReportID;

END