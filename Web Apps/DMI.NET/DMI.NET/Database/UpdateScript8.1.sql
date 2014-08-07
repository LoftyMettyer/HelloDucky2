
---- Drop redundant functions (or renamed)
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetMailMergeDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetMailMergeDefinition];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetCrossTabDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetCrossTabDefinition];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetReportDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetReportDefinition];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCalendarReportOrder]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCalendarReportOrder];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetReportChilds]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetReportChilds];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCalendarReportColumns]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCalendarReportColumns];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetReportColumns]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetReportColumns];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntDefProperties]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntDefProperties];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetEmailGroups]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetEmailGroups];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetEmailAddresses]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetEmailAddresses];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntSaveMailMerge]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntSaveMailMerge];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntSaveCrossTab]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntSaveCrossTab];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntSaveCustomReport]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntSaveCustomReport];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntValidateCrossTab]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntValidateCrossTab];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntValidateReport]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntValidateReport];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntValidateMailMerge]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntValidateMailMerge];
GO









-- modified (chr(9) to be , AS [xxxx] so that columns come back in non string delimated format, also return types are noiw rw/ro/hd instead of readable text
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetUtilityAccessRecords]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetUtilityAccessRecords];
GO
CREATE PROCEDURE [dbo].[spASRIntGetUtilityAccessRecords] (
	@piUtilityType		integer,
	@piID				integer,
	@piFromCopy			integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE
		@sDefaultAccess	varchar(2),
		@sAccessTable	sysname,
		@sKey			varchar(255),
		@sSQL			nvarchar(MAX);

	SET @sAccessTable = '';

	IF @piUtilityType = 17
	BEGIN
		/* Calendar Reports */
		SET @sAccessTable = 'ASRSysCalendarReportAccess';
		SET @sKey = 'dfltaccess CalendarReports';
	END

	IF @piUtilityType = 1
	BEGIN
		/* Cross Tabs */
		SET @sAccessTable = 'ASRSysCrossTabAccess';
		SET @sKey = 'dfltaccess CrossTabs';
	END

	IF @piUtilityType = 2
	BEGIN
		/* Custom Reports */
		SET @sAccessTable = 'ASRSysCustomReportAccess';
		SET @sKey = 'dfltaccess CustomReports';
	END

	IF @piUtilityType = 9
	BEGIN
		/* Mail Merge */
		SET @sAccessTable = 'ASRSysMailMergeAccess';
		SET @sKey = 'dfltaccess MailMerge';
	END

	IF LEN(@sAccessTable) > 0
	BEGIN
		IF (@piID = 0) OR (@piFromCopy = 1)
		BEGIN
			SELECT @sDefaultAccess = SettingValue 
			FROM ASRSysUserSettings
			WHERE UserName = system_user
				AND Section = 'utils&reports'
				AND SettingKey = @sKey;
	
			IF (@sDefaultAccess IS null)
			BEGIN
				SET @sDefaultAccess = 'RW';
			END
		END
		ELSE
		BEGIN
			SET @sDefaultAccess = 'HD';
		END
		
		SET @sSQL = 'SELECT sysusers.name ,
				CASE WHEN	
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
								OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
								AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
							WHERE sysusers.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
						ELSE ';
  
		IF (@piID = 0) OR (@piFromCopy = 1)
		BEGIN
			SET @sSQL = @sSQL + ' ''' + @sDefaultAccess + '''';
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 
				' CASE
					WHEN ' + @sAccessTable + '.access IS null THEN ''' + @sDefaultAccess + '''
					ELSE ' + @sAccessTable + '.access
				END';
		END

		SET @sSQL = @sSQL + 
			' END = ''RW'' THEN ''RW''
			 WHEN	CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
						OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
			ELSE '
  
		IF (@piID = 0) OR (@piFromCopy = 1)
		BEGIN
			SET @sSQL = @sSQL + ' ''' + @sDefaultAccess + '''';
		END
		ELSE
		BEGIN
			SET @sSQL = @sSQL + 
				' CASE
					WHEN ' + @sAccessTable + '.access IS null THEN ''' + @sDefaultAccess + '''
					ELSE ' + @sAccessTable + '.access
				END';
		END

		SET @sSQL = @sSQL + 
			' END = ''RO'' THEN ''RO''
			ELSE ''HD'' 
			END AS [access] ,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
 						OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''1''
				ELSE
					''0''
			END AS [isOwner]
			FROM sysusers
			LEFT OUTER JOIN ' + @sAccessTable + ' ON (sysusers.name = ' + @sAccessTable + '.groupName
				AND ' + @sAccessTable + '.id = ' + convert(nvarchar(100), @piID) + ')
			WHERE sysusers.uid = sysusers.gid
				AND sysusers.uid <> 0 AND NOT (sysusers.name LIKE ''ASRSys%'') AND NOT (sysusers.name LIKE ''db_%'')
			ORDER BY sysusers.name';

			EXEC sp_executesql @sSQL;

	END

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetMailMergeDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetMailMergeDefinition];
GO
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
	SELECT Name, [description], userName AS [owner],		
		tableID AS BaseTableID,		
		selection AS SelectionType,
		picklistID,	
		@psPicklistName AS PicklistName,
		FilterID,
		@psFilterName AS FilterName,
		outputformat AS [Format],		
		outputsave AS [SaveToFile],		
		outputfilename AS [Filename],		
		emailAddrID AS [EmailGroupID],		
		emailSubject,		
		templateFileName,		
		outputscreen AS [DisplayOutputOnScreen],		
		emailasattachment AS [EmailAsAttachment],		
		ISNULL(emailattachmentname,'') AS [EmailAttachmentName],		
		suppressblanks AS SuppressBlankLines,		
		PauseBeforeMerge,		
		outputprinter AS [SendToPrinter],		
		outputprintername AS [PrinterName],		
		documentmapid,		
		manualdocmanheader,
		PromptStart AS PauseBeforeMerge,
		CONVERT(integer, timestamp) AS [Timestamp],
		CASE WHEN @pfPicklistHidden = 1 OR @pfFilterHidden = 1 THEN 'HD' ELSE '' END AS [BaseViewAccess]
	FROM [dbo].[ASRSysMailMergeName]		
	WHERE MailMergeID = @piReportID;		

	-- Columns
	SELECT ASRSysMailMergeColumns.ColumnID AS [ID],
		0 AS [IsExpression],
		0 AS [accesshidden],
		ASRSysColumns.tableID,
		ASRSysColumns.columnName AS [name], 
		ASRSysTables.tableName + '.' + ASRSysColumns.columnName AS [heading],
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
		convert(varchar(MAX), '<Calc> ' + replace(ASRSysExpressions.name, '_', ' ')) AS [heading],
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
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCrossTabDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCrossTabDefinition];
GO
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
 			@piTimestamp			integer	= 0;	

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

	SELECT @psErrorMsg AS ErrorMsg, @psReportName AS Name, @psReportOwner AS [Owner], @psReportDesc AS [Description]
		, @piBaseTableID AS [BaseTableID], @piSelection AS SelectionType
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
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCustomReportDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCustomReportDefinition];
GO

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
		@pfIgnoreZeros				bit;

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


	-- Definition
	SELECT @psReportName AS name, @psReportDesc AS [Description], @piBaseTableID AS baseTableID, @psReportOwner AS [Owner],
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
		ASRSysTables.TableName  + ' Calc> ' + replace(ASRSysExpressions.name, '_', ' ') AS [Heading],
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
		AND cd.sortOrderSequence > 0;

	-- Return the child table information
	SELECT  C.ChildTable AS [TableID],
		T.TableName AS [TableName],
		CASE WHEN (X.Access <> 'HD') OR (X.userName = system_user) THEN isnull(X.ExprID, 0) ELSE 0 END AS [FilterID],
		CASE WHEN (X.Access <> 'HD') OR (X.userName = system_user) THEN isnull(X.Name, '') ELSE '' END AS [FilterName],
		isnull(O.OrderID, 0) AS [OrderID],
	  ISNULL(O.Name, '') AS [OrderName],
	  C.ChildMaxRecords AS [Records], 
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
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetCalendarReportDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetCalendarReportDefinition];
GO
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
 		@piTimestamp				integer;

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


	-- Definition
	SELECT @psReportName AS name, @psReportDesc AS [Description], @piBaseTableID AS baseTableID, @psReportOwner AS [Owner],
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
		(SELECT ISNULL(ASRSysTables.TableName,'') FROM ASRSysTables WHERE ASRSysTables.TableID = ASRSysCalendarReportOrder.TableID) + '.' +
		(SELECT ISNULL(ASRSysColumns.ColumnName,'') FROM ASRSysColumns WHERE ASRSysColumns.ColumnID = ASRSysCalendarReportOrder.ColumnID) AS [Name],
		OrderSequence AS [Sequence],
		OrderType AS [Order]
	FROM [dbo].[ASRSysCalendarReportOrder]
	WHERE calendarReportID = @piCalendarReportID
	ORDER BY OrderSequence;

END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCalculationsForTable]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRGetCalculationsForTable];
GO

CREATE PROCEDURE dbo.[spASRGetCalculationsForTable](@piTableID as integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT ExprID AS ID,
			Name,
			0 AS DataType,
			0 AS Size,
			0 AS Decimals
	 FROM ASRSysExpressions
		WHERE type = 10 AND (returnType = 0 OR type = 10) AND parentComponentID = 0	AND TableID  = @piTableID
		ORDER BY Name;

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetRecordSelection]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetRecordSelection];
GO

CREATE PROCEDURE [dbo].[spASRIntGetRecordSelection]
(
	@psType		varchar(255),
	@piTableID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @fSysSecMgr	bit;

	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT

	IF UPPER(@psType) = 'EMAIL'
	BEGIN
		SELECT emailGroupID AS [ID], name, userName, access , [Description]
		FROM ASRSysEmailGroupName 
		ORDER BY [name];
	END

	IF UPPER(@psType) = 'PICKLIST'
	BEGIN
		SELECT picklistid AS ID, name, username, access, [Description]
		FROM [dbo].[ASRSysPicklistName]
		WHERE (tableid = @piTableID)
			AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
		ORDER BY [name];
	END

	IF UPPER(@psType) = 'ORDER'
	BEGIN
			SELECT orderid AS [ID], name, '' AS username, '' AS access , '' AS [Description]
		FROM ASRSysOrders 
		WHERE tableid = @piTableID AND type = 1 
			ORDER BY [name];
	END

	IF UPPER(@psType) = 'FILTER'
	BEGIN
		SELECT exprid AS ID, name, username, access, [Description]
		FROM [dbo].[ASRSysExpressions]
		WHERE tableid = @piTableID 
			AND type = 11 
			AND (returnType = 3 OR type = 10) 
			AND parentComponentID = 0 
			AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
		ORDER BY [name];
	END
	
	IF UPPER(@psType) = 'CALC'
	BEGIN
		IF @piTableID > 0
		BEGIN
			SELECT exprid AS ID, name, username, access, [Description]
			FROM [dbo].[ASRSysExpressions]
			WHERE (tableid = @piTableID)
				AND  type = 10 
				AND (returnType = 0 OR type = 10) 
				AND parentComponentID = 0 
				AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
			ORDER BY [name];
		END
		ELSE
		BEGIN
			SELECT exprid AS ID, name, username, access, [Description]
			FROM [dbo].[ASRSysExpressions] 
			WHERE  type = 18 
				AND (returnType = 4 OR type = 10) 
				AND parentComponentID = 0 
				AND (@fSysSecMgr = 1 OR username = SYSTEM_USER OR Access <> 'HD')
			ORDER BY [name];
		END
	END
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntDefProperties]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntDefProperties];
GO

CREATE PROCEDURE [dbo].[spASRIntDefProperties] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @Name	nvarchar(255);

	-- Definition details
	EXEC [spASRIntGetUtilityName] @intType, @intID, @Name OUTPUT

	SELECT @name AS Name;

	-- Access details of object
	SELECT convert(varchar, CreatedDate,103) + ' ' + convert(varchar, CreatedDate,108) as [CreatedDate], 
		convert(varchar, SavedDate,103) + ' ' + convert(varchar, SavedDate,108) as [SavedDate], 
		convert(varchar, RunDate,103) + ' ' + convert(varchar, RunDate,108) as [RunDate], 
		Createdby, 
		Savedby, 
		Runby 
	FROM [dbo].[ASRSysUtilAccessLog]
	WHERE UtilID = @intID AND [Type] = @intType;

	-- Get usage of this object
	EXEC sp_ASRIntDefUsage @intType, @intID;

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetUtilityName]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetUtilityName];
GO

CREATE PROCEDURE [dbo].[spASRIntGetUtilityName] (
	@piUtilityType	integer,
	@plngID			integer,
	@psName			varchar(255)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@sTableName			sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500);

	SET @sTableName = '';
	SET @psName = '<unknown>';

	IF @piUtilityType IN (11, 12)  -- Calculations and filters
	BEGIN
		SET @sTableName = 'ASRSysExpressions';
		SET @sIDColumnName = 'ExprID';
  END

	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sIDColumnName = 'ID';
    END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sIDColumnName = 'ID';
    END

	IF @piUtilityType = 1 /* Cross Tab */
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sIDColumnName = 'CrossTabID';
    END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sIDColumnName = 'ID';
    END
        
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sIDColumnName = 'DataTransferID';
    END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sIDColumnName = 'ID';
    END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sIDColumnName = 'functionID';
    END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sIDColumnName = 'ID';
    END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sIDColumnName = 'mailMergeID';
    END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sIDColumnName = 'recordProfileID';
    END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sIDColumnName = 'matchReportID';
    END

	IF (@piUtilityType = 25) /* Workflow */
	BEGIN
		SET @sTableName = 'ASRSysWorkflows';
		SET @sIDColumnName = 'ID';
	END
      	
	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @sName = [' + @sTableName + '].[name]
				FROM [' + @sTableName + ']
				WHERE [' + @sTableName + '].[' + @sIDColumnName + '] = ' + convert(nvarchar(255), @plngID);

		SET @sParamDefinition = N'@sName varchar(255) OUTPUT';
		EXEC sp_executesql @sSQL, @sParamDefinition, @psName OUTPUT;
	END

	IF @psName IS null SET @psName = '<unknown>';
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetEmailAddresses]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetEmailAddresses];
GO

CREATE PROCEDURE [dbo].[spASRIntGetEmailAddresses]
(@baseTableID int)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT convert(char(10),e.emailid) AS [ID], e.name AS [Name]
		FROM ASRSysEmailAddress e
		WHERE e.tableid = @baseTableID OR e.tableid = 0
		ORDER BY e.name;

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetMetadata]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRGetMetadata];
GO

CREATE PROCEDURE [dbo].[spASRGetMetadata] (@Username varchar(255))
WITH ENCRYPTION
AS
BEGIN

	DECLARE @licenseKey			varchar(MAX);

	EXEC [dbo].[sp_ASRIntGetSystemSetting] 'Licence', 'Key', 'moduleCode', @licenseKey OUTPUT, 0, 0;


	SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM dbo.ASRSysTables;

	SELECT ColumnID, TableID, ColumnName, DataType, ColumnType, Use1000Separator, Size, Decimals FROM dbo.ASRSysColumns;

	SELECT ParentID, ChildID FROM dbo.ASRSysRelations;

	SELECT ModuleKey, ParameterKey, ISNULL(ParameterValue,'') AS ParameterValue, ParameterType FROM dbo.ASRSysModuleSetup;

	SELECT * FROM dbo.ASRSysUserSettings WHERE Username = @Username;

	SELECT functionID, functionName, returnType FROM dbo.ASRSysFunctions;

	SELECT * FROM dbo.ASRSysFunctionParameters ORDER BY functionID, parameterIndex;

	SELECT * FROM dbo.ASRSysOperators;

	SELECT * FROM dbo.ASRSysOperatorParameters ORDER BY OperatorID, parameterIndex;
	
	-- Which modules are enabled?
	SELECT 'WORKFLOW' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,1024) AS [enabled]
	UNION
	SELECT 'PERSONNEL' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,1) AS [enabled]
	UNION
	SELECT 'ABSENCE' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,4) AS [enabled]
	UNION
	SELECT 'TRAINING' AS [name],  dbo.udfASRNetIsModuleLicensed(@licenseKey,8) AS [enabled]
	UNION
	SELECT  'VERSIONONE' AS [name], dbo.udfASRNetIsModuleLicensed(@licenseKey,2048) AS [enabled];


	-- Selected system settings
	SELECT * FROM ASRSysSystemSettings;


END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntSaveMailMerge]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntSaveMailMerge];
GO
CREATE PROCEDURE [dbo].[spASRIntSaveMailMerge] (
	@psName				varchar(255),
	@psDescription		varchar(MAX),
	@piTableID			integer,
	@piSelection		integer,
	@piPicklistID		integer,
	@piFilterID			integer,
	@piOutputFormat			integer,
	@pfOutputSave			bit,
	@psOutputFilename		varchar(MAX),
	@piEmailAddrID		integer,
	@psEmailSubject		varchar(MAX),
	@psTemplateFileName	varchar(MAX),
	@pfOutputScreen			bit,
	@psUserName			varchar(255),
	@pfEmailAsAttachment	bit,
	@psEmailAttachmentName	varchar(MAX),
	@pfSuppressBlanks		bit,
	@pfPauseBeforeMerge		bit,
	@pfOutputPrinter			bit,
	@psOutputPrinterName	varchar(255),
	@piDocumentMapID			integer,
	@pfManualDocManHeader		bit,	
	@psAccess			varchar(MAX),
	@psJobsToHide		varchar(MAX),
	@psJobsToHideGroups	varchar(MAX),
	@psColumns			varchar(MAX),
	@psColumns2			varchar(MAX),
	@piID				integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sTemp			varchar(MAX),
			@sColumnDefn	varchar(MAX),
			@sColumnParam	varchar(MAX),
			@iSequence		integer,
			@sType			varchar(MAX),
			@iColExprID		integer,
			@sHeading		varchar(MAX),
			@iSize			integer,
			@iDP			integer,
			@fIsNumeric		bit,
			@fAvge			bit,
			@fCnt			bit,
			@fTot			bit,
			@iSortOrderSequence	integer,
			@sSortOrder		varchar(MAX),
			@fBOC			bit,
			@fPOC			bit,
			@fVOC			bit,
			@fSRV			bit,
			@iCount			integer,
			@fIsNew			bit,
			@sGroup			varchar(255),
			@sAccess		varchar(MAX),
			@sSQL			nvarchar(MAX);
	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''')
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''')
	SET @fIsNew = 0
	/* Insert/update the report header. */
	IF @piID = 0
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysMailMergeName (
			Name, 
			Description, 
			TableID, 
			Selection, 
			PicklistID, 
			FilterID, 
			OutputFormat, 
			OutputSave, 
			OutputFilename, 
			EmailAddrID, 
			EmailSubject, 
			TemplateFileName, 
			OutputScreen, 
			UserName, 
			EMailAsAttachment,
			EmailAttachmentName, 
			SuppressBlanks, 
			PauseBeforeMerge, 
			OutputPrinter,
			OutputPrinterName,
			DocumentMapID,
			ManualDocManHeader,			
			IsLabel, 
			LabelTypeID, 
			PromptStart) 
		VALUES (
			@psName,
			@psDescription,
			@piTableID,
			@piSelection,
			@piPicklistID,
			@piFilterID,
			@piOutputFormat,
			@pfOutputSave,
			@psOutputFilename,
			@piEmailAddrID,
			@psEmailSubject,
			@psTemplateFileName,
			@pfOutputScreen,
			@psUserName,
			@pfEmailAsAttachment,
			@psEmailAttachmentName,
			@pfSuppressBlanks,
			@pfPauseBeforeMerge,
			@pfOutputPrinter,
			@psOutputPrinterName,
			@piDocumentMapID,
			@pfManualDocManHeader,
			0, 
			0, 
			0)
		SET @fIsNew = 1
		/* Get the ID of the inserted record.*/
		SELECT @piID = MAX(MailMergeID) FROM ASRSysMailMergeName
	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysMailMergeName SET 
			Name = @psName,
			Description = @psDescription,
			TableID = @piTableID,
			Selection = @piSelection,
			PicklistID = @piPicklistID,
			FilterID = @piFilterID,
			OutputFormat = @piOutputFormat,
			OutputSave = @pfOutputSave,
			OutputFilename = @psOutputFilename,
			EmailAddrID = @piEmailAddrID,
			EmailSubject = @psEmailSubject,
			TemplateFileName = @psTemplateFileName,
			OutputScreen = @pfOutputScreen,
			EMailAsAttachment = @pfEmailAsAttachment,
			EmailAttachmentName = @psEmailAttachmentName,
			SuppressBlanks = @pfSuppressBlanks,
			PauseBeforeMerge = @pfPauseBeforeMerge,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName,
			DocumentMapID = @piDocumentMapID,
			ManualDocManHeader = @pfManualDocManHeader,
			IsLabel = 0,
			LabelTypeID = 0,
			PromptStart = 0
		WHERE MailMergeID = @piID
		/* Delete existing report details. */
		DELETE FROM ASRSysMailMergeColumns
		WHERE MailMergeID = @piID
	END
	/* Create the details records. */
	SET @sTemp = @psColumns
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sColumnDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)
			IF len(@sTemp) <= 7000
			BEGIN
				SET @sTemp = @sTemp + LEFT(@psColumns2, 1000)
				IF len(@psColumns2) > 1000
				BEGIN
					SET @psColumns2 = SUBSTRING(@psColumns2, 1001, len(@psColumns2) - 1000)
				END
				ELSE
				BEGIN
					SET @psColumns2 = ''
				END
			END
		END
		ELSE
		BEGIN
			SET @sColumnDefn = @sTemp
			SET @sTemp = ''
		END
		/* Rip out the column definition parameters. */
		SET @iSequence = 0
		SET @sType = ''
		SET @iColExprID = 0
		SET @iSize = 0
		SET @iDP = 0
		SET @fIsNumeric = 0
		SET @iSortOrderSequence = 0
		SET @sSortOrder = ''
		SET @fBOC = 0
		SET @fPOC = 0
		SET @fVOC = 0
		SET @fSRV = 0
		SET @iCount = 0
		WHILE LEN(@sColumnDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sColumnDefn) > 0
			BEGIN
				SET @sColumnParam = LEFT(@sColumnDefn, CHARINDEX('||', @sColumnDefn) - 1)
				SET @sColumnDefn = RIGHT(@sColumnDefn, LEN(@sColumnDefn) - CHARINDEX('||', @sColumnDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sColumnParam = @sColumnDefn
				SET @sColumnDefn = ''
			END
			IF @iCount = 0 SET @iSequence = convert(integer, @sColumnParam)
			IF @iCount = 1 SET @sType = @sColumnParam
			IF @iCount = 2 SET @iColExprID = convert(integer, @sColumnParam)
			IF @iCount = 3 SET @iSize = convert(integer, @sColumnParam)
			IF @iCount = 4 SET @iDP = convert(integer, @sColumnParam)
			IF @iCount = 5 SET @fIsNumeric = convert(bit, @sColumnParam)
			IF @iCount = 6 SET @iSortOrderSequence = convert(integer, @sColumnParam)
			IF @iCount = 7 SET @sSortOrder = @sColumnParam
			SET @iCount = @iCount + 1
		END
		INSERT ASRSysMailMergeColumns (MailMergeID,Type, ColumnID, SortOrderSequence, SortOrder, Size, Decimals)
		VALUES (@piID, @sType, @iColExprID, @iSortOrderSequence, @sSortOrder, @iSize, @iDP)
	END
	DELETE FROM ASRSysMailMergeAccess WHERE ID = @piID
	INSERT INTO ASRSysMailMergeAccess (ID, groupName, access)
		(SELECT @piID, sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.name <> 'ASRSysGroup'
			AND sysusers.uid <> 0)
	SET @sTemp = @psAccess
	
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			IF EXISTS (SELECT * FROM ASRSysMailMergeAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysMailMergeAccess
					SET access = @sAccess
					WHERE ID = @piID
						AND groupName = @sGroup
		END
	END
	IF (@fIsNew = 1)
	BEGIN
		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (9, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 9
		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (9, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 9
		END
	END
	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + ''''
		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' +@psJobsToHide + ')
				AND groupName IN (' + @psJobsToHideGroups + ')'
		EXEC sp_executesql @sSQL
		SET @sSQL = 'INSERT INTO ASRSysBatchJobAccess
			(ID, groupName, access)
			(SELECT ASRSysBatchJobName.ID, 
				sysusers.name,
				CASE
					WHEN (SELECT count(*)
						FROM ASRSysGroupPermissions
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
							OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')

						WHERE sysusers.Name = ASRSysGroupPermissions.groupname
							AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
					ELSE ''HD''
				END
			FROM sysusers,
				ASRSysBatchJobName
			WHERE sysusers.uid = sysusers.gid
				AND sysusers.uid <> 0
				AND sysusers.name IN (' + @psJobsToHideGroups + ')
				AND ASRSysBatchJobName.ID IN (' + @psJobsToHide + '))'
		EXEC sp_executesql @sSQL
	END
END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntSaveCrossTab]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntSaveCrossTab];
GO
CREATE PROCEDURE [dbo].[spASRIntSaveCrossTab] (
	@psName				varchar(255),
	@psDescription		varchar(MAX),
	@piTableID			integer,
	@piSelection		integer,
	@piPicklistID		integer,
	@piFilterID			integer,
	@pfPrintFilter		bit,
	@psUserName			varchar(255),
	@piHColID			integer,
	@psHStart			varchar(100),
	@psHStop			varchar(100),
	@psHStep			varchar(100),
	@piVColID			integer,
	@psVStart			varchar(100),
	@psVStop			varchar(100),
	@psVStep			varchar(100),
	@piPColID			integer,
	@psPStart			varchar(100),
	@psPStop			varchar(100),
	@psPStep			varchar(100),
	@piIType			integer,
	@piIColID			integer,
	@pfPercentage		bit,
	@pfPerPage			bit,
	@pfSuppress			bit,
	@pfUse1000Separator	bit,
	@pfOutputPreview	bit,
	@piOutputFormat		integer,
	@pfOutputScreen		bit,
	@pfOutputPrinter	bit,
	@psOutputPrinterName	varchar(MAX),
	@pfOutputSave		bit,
	@piOutputSaveExisting	integer,
	@pfOutputEmail		bit,
	@piOutputEmailAddr	integer,
	@psOutputEmailSubject	varchar(MAX),
	@psOutputEmailAttachAs	varchar(MAX),
	@psOutputFilename	varchar(MAX),
	@psAccess			varchar(MAX),
	@psJobsToHide		varchar(MAX),
	@psJobsToHideGroups	varchar(MAX),
	@piID				integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
			@fIsNew		bit,
			@sTemp		varchar(MAX),
			@iCount		integer,
			@sGroup		varchar(MAX),
			@sAccess	varchar(MAX),
			@sSQL		nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''')
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''')

	SET @fIsNew = 0

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysCrossTab (
			Name, 
			Description, 
			TableID, 
			Selection, 
			PicklistID, 
			FilterID, 
 			PrintFilterHeader, 
 			UserName, 
 			HorizontalColID, 
 			HorizontalStart, 
 			HorizontalStop, 
 			HorizontalStep, 
			VerticalColID, 
			VerticalStart, 
			VerticalStop, 
			VerticalStep, 
			PageBreakColID, 
			PageBreakStart, 
			PageBreakStop, 
			PageBreakStep, 
			IntersectionType, 
			IntersectionColID, 
			Percentage, 
			PercentageofPage, 
			SuppressZeros, 
			ThousandSeparators, 
			OutputPreview, 
			OutputFormat, 
			OutputScreen, 
			OutputPrinter, 
			OutputPrinterName, 
			OutputSave, 
			OutputSaveExisting, 
			OutputEmail, 
			OutputEmailAddr, 
			OutputEmailSubject, 
			OutputEmailAttachAs, 
			OutputFileName)
		VALUES (
			@psName,
			@psDescription,
			@piTableID,
			@piSelection,
			@piPicklistID,
			@piFilterID,
			@pfPrintFilter,
			@psUserName,
			@piHColID,
			@psHStart,
			@psHStop,
			@psHStep,
			@piVColID,
			@psVStart,
			@psVStop,
			@psVStep,
			@piPColID,
			@psPStart,
			@psPStop,
			@psPStep,
			@piIType,
			@piIColID,
			@pfPercentage,
			@pfPerPage,
			@pfSuppress,
			@pfUse1000Separator,
			@pfOutputPreview,
			@piOutputFormat,
			@pfOutputScreen,
			@pfOutputPrinter,
			@psOutputPrinterName,
			@pfOutputSave,
			@piOutputSaveExisting,
			@pfOutputEmail,
			@piOutputEmailAddr,
			@psOutputEmailSubject,
			@psOutputEmailAttachAs,
			@psOutputFilename
		)

		SET @fIsNew = 1
		/* Get the ID of the inserted record.*/
		SELECT @piID = MAX(CrossTabID) FROM ASRSysCrossTab
	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSysCrossTab SET 
			Name = @psName,
			Description = @psDescription,
			TableID = @piTableID,
			Selection = @piSelection,
			PicklistID = @piPicklistID,
			FilterID = @piFilterID,
			PrintFilterHeader = @pfPrintFilter,
			HorizontalColID = @piHColID,
			HorizontalStart = @psHStart,
			HorizontalStop = @psHStop,
			HorizontalStep = @psHStep,	
			VerticalColID = @piVColID,
			VerticalStart = @psVStart,
			VerticalStop = @psVStop,
			VerticalStep = @psVStep,	
			PageBreakColID = @piPColID,
			PageBreakStart = @psPStart,
			PageBreakStop = @psPStop,
			PageBreakStep = @psPStep,	
			IntersectionType = @piIType,
			IntersectionColID = @piIColID,
			Percentage = @pfPercentage,
			PercentageofPage = @pfPerPage,
			SuppressZeros = @pfSuppress,
			ThousandSeparators = @pfUse1000Separator,
			OutputPreview = @pfOutputPreview,
			OutputFormat = @piOutputFormat,
			OutputScreen = @pfOutputScreen,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName,
			OutputSave = @pfOutputSave,
			OutputSaveExisting = @piOutputSaveExisting,
			OutputEmail = @pfOutputEmail,
			OutputEmailAddr = @piOutputEmailAddr,
			OutputEmailSubject = @psOutputEmailSubject,
			OutputEmailAttachAs = @psOutputEmailAttachAs,
			OutputFileName = @psOutputFilename
		WHERE CrossTabID = @piID
	END

	DELETE FROM ASRSysCrossTabAccess WHERE ID = @piID

	INSERT INTO ASRSysCrossTabAccess (ID, groupName, access)
		(SELECT @piID, sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.name <> 'ASRSysGroup'
			AND sysusers.uid <> 0)

	SET @sTemp = @psAccess
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			IF EXISTS (SELECT * FROM ASRSysCrossTabAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysCrossTabAccess
					SET access = @sAccess
					WHERE ID = @piID
						AND groupName = @sGroup
		END
	END

	IF (@fIsNew = 1)
	BEGIN
		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (1, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 1

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (1, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 1
		END
	END
	
	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + ''''
		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' + @psJobsToHide + ')
				AND groupName IN (' + @psJobsToHideGroups + ')'
		EXEC sp_executesql @sSQL

		SET @sSQL = 'INSERT INTO ASRSysBatchJobAccess
			(ID, groupName, access)
			(SELECT ASRSysBatchJobName.ID, 
				sysusers.name,
				CASE
					WHEN (SELECT count(*)
						FROM ASRSysGroupPermissions
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
							OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
						WHERE sysusers.Name = ASRSysGroupPermissions.groupname
							AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
					ELSE ''HD''
				END
			FROM sysusers,
				ASRSysBatchJobName
			WHERE sysusers.uid = sysusers.gid
				AND sysusers.uid <> 0
				AND sysusers.name IN (' + @psJobsToHideGroups + ')
				AND ASRSysBatchJobName.ID IN (' + @psJobsToHide + '))'
		EXEC sp_executesql @sSQL
	END
	
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntSaveCustomReport]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntSaveCustomReport];
GO
CREATE PROCEDURE [dbo].[spASRIntSaveCustomReport] (
	@psName						varchar(255),
	@psDescription				varchar(MAX),
	@piBaseTableID				integer,
	@pfAllRecords				bit,
	@piPicklistID				integer,
	@piFilterID					integer,
	@piParent1TableID			integer,
	@piParent1FilterID			integer,
	@piParent2TableID			integer,
	@piParent2FilterID			integer,
	@pfSummary					bit,
	@pfPrintFilterHeader		bit,
	@psUserName					varchar(255),
	@pfOutputPreview			bit,
	@piOutputFormat				integer,
	@pfOutputScreen				bit,
	@pfOutputPrinter			bit,
	@psOutputPrinterName		varchar(MAX),
	@pfOutputSave				bit,
	@piOutputSaveExisting		integer,
	@pfOutputEmail				bit,
	@piOutputEmailAddr			integer,
	@psOutputEmailSubject		varchar(MAX),
	@psOutputEmailAttachAs		varchar(MAX),
	@psOutputFilename			varchar(MAX),
	@pfParent1AllRecords		bit,
	@piParent1Picklist			integer,
	@pfParent2AllRecords		bit,
	@piParent2Picklist			integer,
	@psAccess					varchar(MAX),
	@psJobsToHide				varchar(MAX),
	@psJobsToHideGroups			varchar(MAX),
	@psColumns					varchar(MAX),
	@psColumns2					varchar(MAX),
	@psChildString				varchar(MAX),
	@piID						integer					OUTPUT,
	@pfIgnoreZeros				bit
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sTemp					varchar(MAX),
			@sColumnDefn			varchar(MAX),
			@sColumnParam			varchar(MAX),
			@iSequence				integer,
			@sType					varchar(MAX),
			@iColExprID				integer,
			@sHeading				varchar(MAX),
			@iSize					integer,
			@iDP					integer,
			@fIsNumeric				bit,
			@fAvge					bit,
			@fCnt					bit,
			@fTot					bit,
			@fHidden				bit,
			@fGroupWithNext			bit,
			@iSortOrderSequence		integer,
			@sSortOrder				varchar(MAX),
			@fBOC					bit,
			@fPOC					bit,
			@fVOC					bit,
			@fSRV					bit,
			@fRepetition 			integer,
			@iCount					integer,
			@fIsNew					bit,
			@iChildTableID			integer,
			@iChildFilterID			integer,
			@iChildOrderID			integer,
			@iChildMaxRecords		integer,
			@sChildDefn				varchar(MAX),
			@sChildParam			varchar(MAX),
			@sGroup					varchar(255),
			@sAccess				varchar(MAX),
			@sSQL					nvarchar(MAX);

	/* Clean the input string parameters. */
	IF len(@psJobsToHide) > 0 SET @psJobsToHide = replace(@psJobsToHide, '''', '''''')
	IF len(@psJobsToHideGroups) > 0 SET @psJobsToHideGroups = replace(@psJobsToHideGroups, '''', '''''')

	SET @fIsNew = 0

	/* Insert/update the report header. */
	IF (@piID = 0)
	BEGIN
		/* Creating a new report. */
		INSERT ASRSysCustomReportsName (
			Name, 
			[Description], 
			BaseTable, 
			AllRecords, 
			Picklist, 
			Filter, 
 			Parent1Table, 
 			Parent1Filter, 
 			Parent2Table, 
 			Parent2Filter, 
 			Summary,
 			IgnoreZeros, 
 			PrintFilterHeader, 
 			UserName, 
 			OutputPreview, 
 			OutputFormat, 
 			OutputScreen, 
 			OutputPrinter, 
 			OutputPrinterName, 
 			OutputSave, 
 			OutputSaveExisting, 
 			OutputEmail, 
 			OutputEmailAddr, 
 			OutputEmailSubject, 
 			OutputEmailAttachAs, 
 			OutputFileName, 
 			Parent1AllRecords, 
 			Parent1Picklist, 
 			Parent2AllRecords, 
 			Parent2Picklist)
 		VALUES (
 			@psName,
 			@psDescription,
 			@piBaseTableID,
 			@pfAllRecords,
 			@piPicklistID,
 			@piFilterID,
 			@piParent1TableID,
 			@piParent1FilterID,
 			@piParent2TableID,
 			@piParent2FilterID,
 			@pfSummary,
 			@pfIgnoreZeros,
 			@pfPrintFilterHeader,
 			@psUserName,
 			@pfOutputPreview,
 			@piOutputFormat,
 			@pfOutputScreen,
 			@pfOutputPrinter,
 			@psOutputPrinterName,
 			@pfOutputSave,
 			@piOutputSaveExisting,
 			@pfOutputEmail,
 			@piOutputEmailAddr,
 			@psOutputEmailSubject,
 			@psOutputEmailAttachAs,
 			@psOutputFilename,
 			@pfParent1AllRecords,
 			@piParent1Picklist,
 			@pfParent2AllRecords,
 			@piParent2Picklist
		)

		SET @fIsNew = 1
		/* Get the ID of the inserted record.*/
		SELECT @piID = MAX(ID) FROM ASRSysCustomReportsName
	END
	ELSE
	BEGIN
		/* Updating an existing report. */
		UPDATE ASRSYSCustomReportsName SET 
			Name = @psName,
			Description = @psDescription,
			BaseTable = @piBaseTableID,
			AllRecords = @pfAllRecords,
			Picklist = @piPicklistID,
			Filter = @piFilterID,
			Parent1Table = @piParent1TableID,
			Parent1Filter = @piParent1FilterID,
			Parent2Table = @piParent2TableID,
			Parent2Filter = @piParent2FilterID,
			Summary = @pfSummary,
			IgnoreZeros = @pfIgnoreZeros,
			PrintFilterHeader = @pfPrintFilterHeader,
			OutputPreview = @pfOutputPreview,
			OutputFormat = @piOutputFormat,
			OutputScreen = @pfOutputScreen,
			OutputPrinter = @pfOutputPrinter,
			OutputPrinterName = @psOutputPrinterName,
			OutputSave = @pfOutputSave,
			OutputSaveExisting = @piOutputSaveExisting,
			OutputEmail = @pfOutputEmail,
			OutputEmailAddr = @piOutputEmailAddr,
			OutputEmailSubject = @psOutputEmailSubject,
			OutputEmailAttachAs = @psOutputEmailAttachAs,
			OutputFileName = @psOutputFilename,
			Parent1AllRecords = @pfParent1AllRecords,
			Parent1Picklist = @piParent1Picklist,
			Parent2AllRecords = @pfParent2AllRecords,
			Parent2Picklist = @piParent2Picklist
		WHERE ID = @piID

		/* Delete existing report details. */
		DELETE FROM ASRSysCustomReportsDetails 
		WHERE customReportID = @piID
	END

	/* Create the details records. */
	SET @sTemp = @psColumns

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sColumnDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)

			IF len(@sTemp) <= 7000
			BEGIN
				SET @sTemp = @sTemp + LEFT(@psColumns2, 1000)
				IF len(@psColumns2) > 1000
				BEGIN
					SET @psColumns2 = SUBSTRING(@psColumns2, 1001, len(@psColumns2) - 1000)
				END
				ELSE
				BEGIN
					SET @psColumns2 = ''
				END
			END
		END
		ELSE
		BEGIN
			SET @sColumnDefn = @sTemp
			SET @sTemp = ''
		END

		/* Rip out the column definition parameters. */
		SET @iSequence = 0
		SET @sType = ''
		SET @iColExprID = 0
		SET @sHeading = ''
		SET @iSize = 0
		SET @iDP = 0
		SET @fIsNumeric = 0
		SET @fAvge = 0
		SET @fCnt = 0
		SET @fTot = 0
		SET @fHidden = 0
		SET @fGroupWithNext = 0
		SET @iSortOrderSequence = 0
		SET @sSortOrder = ''
		SET @fBOC = 0
		SET @fPOC = 0
		SET @fVOC = 0
		SET @fSRV = 0
		SET @fRepetition = 0
		SET @iCount = 0
		
		WHILE LEN(@sColumnDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sColumnDefn) > 0
			BEGIN
				SET @sColumnParam = LEFT(@sColumnDefn, CHARINDEX('||', @sColumnDefn) - 1)
				SET @sColumnDefn = RIGHT(@sColumnDefn, LEN(@sColumnDefn) - CHARINDEX('||', @sColumnDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sColumnParam = @sColumnDefn
				SET @sColumnDefn = ''
			END

			IF @iCount = 0 SET @iSequence = convert(integer, @sColumnParam)
			IF @iCount = 1 SET @sType = @sColumnParam
			IF @iCount = 2 SET @iColExprID = convert(integer, @sColumnParam)
			IF @iCount = 3 SET @sHeading = @sColumnParam
			IF @iCount = 4 SET @iSize = convert(integer, @sColumnParam)
			IF @iCount = 5 SET @iDP = convert(integer, @sColumnParam)
			IF @iCount = 6 SET @fIsNumeric = convert(bit, @sColumnParam)
			IF @iCount = 7 SET @fAvge = convert(bit, @sColumnParam)
			IF @iCount = 8 SET @fCnt = convert(bit, @sColumnParam)
			IF @iCount = 9 SET @fTot = convert(bit, @sColumnParam)
			IF @iCount = 10 SET @fHidden = convert(bit, @sColumnParam)
			IF @iCount = 11 SET @fGroupWithNext = convert(bit, @sColumnParam)
			IF @iCount = 12 SET @iSortOrderSequence = convert(integer, @sColumnParam)
			IF @iCount = 13 SET @sSortOrder = @sColumnParam
			IF @iCount = 14 SET @fBOC = convert(bit, @sColumnParam)
			IF @iCount = 15 SET @fPOC = convert(bit, @sColumnParam)
			IF @iCount = 16 SET @fVOC = convert(bit, @sColumnParam)
			IF @iCount = 17 SET @fSRV = convert(bit, @sColumnParam)
			IF @iCount = 18 SET @fRepetition = convert(integer, @sColumnParam)

			SET @iCount = @iCount + 1
		END

		INSERT ASRSysCustomReportsDetails 
			(customReportID, sequence, type, colExprID, heading, size, dp, isNumeric, avge, 
			cnt, tot, hidden, GroupWithNextColumn, 	sortOrderSequence, sortOrder, boc, poc, voc, srv, repetition) 
		VALUES (@piID, @iSequence, @sType, @iColExprID, @sHeading, @iSize, @iDP, @fIsNumeric, @fAvge, 
			@fCnt, @fTot, @fHidden, @fGroupWithNext, @iSortOrderSequence, @sSortOrder, @fBOC, @fPOC, @fVOC, @fSRV, @fRepetition)

	END

	/* Create the table records. */

	IF (@fIsNew = 0)
	BEGIN
		/* Delete existing report child tables. */
		DELETE FROM ASRSysCustomReportsChildDetails 
		WHERE customReportID = @piID
	END

	SET @sTemp = @psChildString

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX('**', @sTemp) > 0
		BEGIN
			SET @sChildDefn = LEFT(@sTemp, CHARINDEX('**', @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX('**', @sTemp) - 1)
		END
		ELSE
		BEGIN
			SET @sChildDefn = @sTemp
			SET @sTemp = ''
		END

		/* Rip out the column definition parameters. */
		SET @iChildTableID = 0
		SET @iChildFilterID = 0
		SET @iChildOrderID = 0
		SET @iChildMaxRecords = 0
		SET @iCount = 0
		
		WHILE LEN(@sChildDefn) > 0
		BEGIN
			IF CHARINDEX('||', @sChildDefn) > 0
			BEGIN
				SET @sChildParam = LEFT(@sChildDefn, CHARINDEX('||', @sChildDefn) - 1)
				SET @sChildDefn = RIGHT(@sChildDefn, LEN(@sChildDefn) - CHARINDEX('||', @sChildDefn) - 1)
			END
			ELSE
			BEGIN
				SET @sChildParam = @sChildDefn
				SET @sChildDefn = ''
			END

			IF @iCount = 0 SET @iChildTableID = convert(integer, @sChildParam)
			IF @iCount = 1 SET @iChildFilterID = convert(integer, @sChildParam)
			IF @iCount = 2 SET @iChildOrderID = convert(integer, @sChildParam)
			IF @iCount = 3 SET @iChildMaxRecords = convert(integer, @sChildParam)
	
			SET @iCount = @iCount + 1
		END

		INSERT ASRSysCustomReportsChildDetails 
			(customReportID, childtable, childfilter, childorder, childmaxrecords) 
		VALUES (@piID, @iChildTableID, @iChildFilterID, @iChildOrderID, @iChildMaxRecords)

	END

	DELETE FROM ASRSysCustomReportAccess WHERE ID = @piID
	INSERT INTO ASRSysCustomReportAccess (ID, groupName, access)
		(SELECT @piID, sysusers.name,
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.Name = ASRSysGroupPermissions.groupname
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
				ELSE 'HD'
			END
		FROM sysusers
		WHERE sysusers.uid = sysusers.gid
			AND sysusers.name <> 'ASRSysGroup'
			AND sysusers.uid <> 0)

	SET @sTemp = @psAccess
	
	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sGroup = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			SET @sAccess = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = SUBSTRING(@sTemp, CHARINDEX(char(9), @sTemp) + 1, LEN(@sTemp) - (CHARINDEX(char(9), @sTemp)))
	
			IF EXISTS (SELECT * FROM ASRSysCustomReportAccess
				WHERE ID = @piID
				AND groupName = @sGroup
				AND access <> 'RW')
				UPDATE ASRSysCustomReportAccess
					SET access = @sAccess
					WHERE ID = @piID
						AND groupName = @sGroup
		END
	END

	IF (@fIsNew = 1)
	BEGIN
		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (2, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 2

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (2, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 2
		END
	END

	IF LEN(@psJobsToHide) > 0 
	BEGIN
		SET @psJobsToHideGroups = '''' + REPLACE(SUBSTRING(LEFT(@psJobsToHideGroups, LEN(@psJobsToHideGroups) - 1), 2, LEN(@psJobsToHideGroups)-1), char(9), ''',''') + ''''

		SET @sSQL = 'DELETE FROM ASRSysBatchJobAccess 
			WHERE ID IN (' +@psJobsToHide + ')
				AND groupName IN (' + @psJobsToHideGroups + ')'
		EXEC sp_executesql @sSQL

		SET @sSQL = 'INSERT INTO ASRSysBatchJobAccess
			(ID, groupName, access)
			(SELECT ASRSysBatchJobName.ID, 
				sysusers.name,
				CASE
					WHEN (SELECT count(*)
						FROM ASRSysGroupPermissions
						INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
							AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
							OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
						INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
							AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
						WHERE sysusers.Name = ASRSysGroupPermissions.groupname
							AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
					ELSE ''HD''
				END
			FROM sysusers,
				ASRSysBatchJobName
			WHERE sysusers.uid = sysusers.gid
				AND sysusers.uid <> 0
				AND sysusers.name IN (' + @psJobsToHideGroups + ')
				AND ASRSysBatchJobName.ID IN (' + @psJobsToHide + '))'
		EXEC sp_executesql @sSQL
	END
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntValidateCrossTab]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntValidateCrossTab];
GO
CREATE PROCEDURE [dbo].[spASRIntValidateCrossTab] (
	@psUtilName 		varchar(255), 
	@piUtilID 			integer, 
	@piTimestamp 		integer, 
	@piBasePicklistID	integer, 
	@piBaseFilterID 	integer, 
	@piEmailGroupID 	integer, 
	@psHiddenGroups 	varchar(MAX), 
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT, /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								3 = definition changed by someone else, overwrite ? */
	@psDeletedFilters 	varchar(MAX)	OUTPUT,
	@psHiddenFilters 	varchar(MAX)	OUTPUT,
	@psJobIDsToHide		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp				integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,
			@sTemp					varchar(MAX),
			@sCurrentID				varchar(MAX),
			@sParameter				varchar(MAX),
			@sExprName  			varchar(MAX),
			@sBatchJobName			varchar(MAX),
			@iBatchJobID			integer,
			@iBatchJobScheduled		integer,
			@sBatchJobRoleToPrompt	varchar(MAX),
			@iNonHiddenCount		integer,
			@sBatchJobUserName		sysname,
			@sJobName				varchar(MAX),
			@sCurrentUserGroup		sysname,
			@fBatchJobsOK			bit,
			@sScheduledUserGroups	varchar(MAX),
			@sScheduledJobDetails	varchar(MAX),
			@sCurrentUserAccess		varchar(MAX),
			@iOwnedJobCount 		integer,
			@sOwnedJobDetails		varchar(MAX),
			@sOwnedJobIDs			varchar(MAX),
			@sNonOwnedJobDetails	varchar(MAX),
			@sHiddenGroupsList		varchar(MAX),
			@sHiddenGroup			varchar(MAX),
			@fSysSecMgr				bit,
			@sActualUserName		sysname,
			@iUserGroupID			integer;

	SET @fBatchJobsOK = 1
	SET @sScheduledUserGroups = ''
	SET @sScheduledJobDetails = ''
	SET @iOwnedJobCount = 0
	SET @sOwnedJobDetails = ''
	SET @sOwnedJobIDs = ''
	SET @sNonOwnedJobDetails = ''

	SELECT @sCurrentUser = SYSTEM_USER
	SET @psErrorMsg = ''
	SET @piErrorCode = 0
	SET @psDeletedFilters = ''
	SET @psHiddenFilters = ''

	exec spASRIntSysSecMgr @fSysSecMgr OUTPUT
	
 	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The cross tab has been deleted by another user. Save as a new definition ?'
			SET @piErrorCode = 2
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysCrossTab
			WHERE CrossTabID = @piUtilID

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					1, 
					@piUtilID,
					@sAccess	OUTPUT
		
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The cross tab has been amended by another user and is now Read Only. Save as a new definition ?'
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The cross tab has been amended by another user. Would you like to overwrite this definition ?'
					SET @piErrorCode = 3
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the cross tab name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCrossTab
			WHERE name = @psUtilName
				AND CrossTabID <> @piUtilID
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCrossTab
			WHERE name = @psUtilName
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'A cross tab called ''' + @psUtilName + ''' already exists.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piBasePicklistID > 0)
	BEGIN
		/* Check that the Base table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piBasePicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table picklist has been deleted by another user.'
			SET @piErrorCode = 1
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piBasePicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table picklist has been made hidden by another user.'
				SET @piErrorCode = 1
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseFilterID > 0)
	BEGIN
		/* Check that the Base table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piBaseFilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table filter has been deleted by another user.'
			SET @piErrorCode = 1
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piBaseFilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table filter has been made hidden by another user.'
				SET @piErrorCode = 1
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piEmailGroupID > 0)
	BEGIN
		/* Check that the email group exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysEmailGroupName 
		WHERE emailGroupID = @piEmailGroupID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The email group has been deleted by another user.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piUtilID > 0) AND (len(@psHiddenGroups) > 0)
	BEGIN
		SELECT @sOwner = userName
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID

		IF (@sOwner = @sCurrentUser) 
		BEGIN
			EXEC spASRIntGetActualUserDetails
				@sActualUserName OUTPUT,
				@sCurrentUserGroup OUTPUT,
				@iUserGroupID OUTPUT

			DECLARE @HiddenGroups TABLE(groupName sysname, groupID integer)
			SET @sHiddenGroupsList = substring(@psHiddenGroups, 2, len(@psHiddenGroups)-2)
			WHILE LEN(@sHiddenGroupsList) > 0
			BEGIN
				IF CHARINDEX(char(9), @sHiddenGroupsList) > 0
				BEGIN
					SET @sHiddenGroup = LEFT(@sHiddenGroupsList, CHARINDEX(char(9), @sHiddenGroupsList) - 1)
					SET @sHiddenGroupsList = RIGHT(@sHiddenGroupsList, LEN(@sHiddenGroupsList) - CHARINDEX(char(9), @sHiddenGroupsList))
				END
				ELSE
				BEGIN
					SET @sHiddenGroup = @sHiddenGroupsList
					SET @sHiddenGroupsList = ''
				END

				INSERT INTO @HiddenGroups (groupName, groupID) (SELECT @sHiddenGroup, uid FROM sysusers WHERE name = @sHiddenGroup)
			END

			DECLARE batchjob_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
				ASRSysBatchJobName.Username,
				ASRSysCrossTab.Name AS 'JobName'
	 		FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID 
			INNER JOIN ASRSysCrossTab ON ASRSysCrossTab.CrossTabID = ASRSysBatchJobDetails.JobID
			LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.access <> 'HD'
				AND ASRSysBatchJobAccess.groupName IN (SELECT name FROM sysusers WHERE uid IN (SELECT groupID FROM @HiddenGroups))
				AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysBatchJobDetails.JobType = 'Cross Tab'
				AND ASRSysBatchJobDetails.JobID IN (@piUtilID)
			GROUP BY ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				ASRSysBatchJobName.Username,
				ASRSysCrossTab.Name

			OPEN batchjob_cursor
			FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
				@iBatchJobID,
				@iBatchJobScheduled,
				@sBatchJobRoleToPrompt,
				@iNonHiddenCount,
				@sBatchJobUserName,
				@sJobName	
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @sCurrentUserAccess = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
								OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		 						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
						WHEN ASRSysBatchJobName.userName = system_user THEN 'RW'
						ELSE
							CASE
								WHEN ASRSysBatchJobAccess.access IS null THEN 'HD'
								ELSE ASRSysBatchJobAccess.access
							END
					END 
				FROM sysusers b
				INNER JOIN sysusers a ON b.uid = a.gid
				LEFT OUTER JOIN ASRSysBatchJobAccess ON (b.name = ASRSysBatchJobAccess.groupName
					AND ASRSysBatchJobAccess.id = @iBatchJobID)
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobAccess.ID = ASRSysBatchJobName.ID
				WHERE a.Name = @sActualUserName

				IF @sBatchJobUserName = @sOwner
				BEGIN
					/* Found a Batch Job whose owner is the same. */
					IF (@iBatchJobScheduled = 1) AND
						(len(@sBatchJobRoleToPrompt) > 0) AND
						(@sBatchJobRoleToPrompt <> @sCurrentUserGroup) AND
						(CHARINDEX(char(9) + @sBatchJobRoleToPrompt + char(9), @psHiddenGroups) > 0)
					BEGIN
						/* Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sBatchJobRoleToPrompt + '<BR>'

						IF @sCurrentUserAccess = 'HD'
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0 
						BEGIN
							SET @iOwnedJobCount = @iOwnedJobCount + 1
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains Cross Tab ' + @sJobName + ')' + '<BR>'
							SET @sOwnedJobIDs = @sOwnedJobIDs +
								CASE 
									WHEN Len(@sOwnedJobIDs) > 0 THEN ', '
									ELSE ''
								END +  convert(varchar(100), @iBatchJobID)
						END
					END
				END			
				ELSE
				BEGIN
					/* Found a Batch Job whose owner is not the same. */
					SET @fBatchJobsOK = 0
	    
					IF @sCurrentUserAccess = 'HD'
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
					END
				END

				FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
					@iBatchJobID,
					@iBatchJobScheduled,
					@sBatchJobRoleToPrompt,
					@iNonHiddenCount,
					@sBatchJobUserName,
					@sJobName	
			END

			CLOSE batchjob_cursor
			DEALLOCATE batchjob_cursor	
		END
	END

	IF @fBatchJobsOK = 0
	BEGIN
		SET @piErrorCode = 1

		IF Len(@sScheduledJobDetails) > 0 
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden from the following user groups :'  + '<BR><BR>' +
				@sScheduledUserGroups  +
				'<BR>as it is used in the following batch jobs which are scheduled to be run by these user groups :<BR><BR>' +
				@sScheduledJobDetails
		END
		ELSE
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden as it is used in the following batch jobs of which you are not the owner :<BR><BR>' +
				@sNonOwnedJobDetails
	      	END
	END
	ELSE
	BEGIN
	    	IF (@iOwnedJobCount > 0) 
		BEGIN
			SET @piErrorCode = 4
			SET @psErrorMsg = 'Making this definition hidden to user groups will automatically make the following definition(s), of which you are the owner, hidden to the same user groups:<BR><BR>' +
				@sOwnedJobDetails + '<BR><BR>' +
				'Do you wish to continue ?'
		END
	END

	SET @psJobIDsToHide = @sOwnedJobIDs

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntValidateCustomReport]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntValidateCustomReport];
GO
CREATE PROCEDURE [dbo].[spASRIntValidateCustomReport] (
	@psUtilName 				varchar(255), 
	@piUtilID 					integer, 
	@piTimestamp 				integer, 
	@piBasePicklistID			integer, 
	@piBaseFilterID 			integer, 
	@piEmailGroupID 			integer, 
	@piParent1PicklistID		integer, 
	@piParent1FilterID 			integer, 
	@piParent2PicklistID		integer, 
	@piParent2FilterID 			integer, 
	@piChildFilterID 			varchar(100),			/* tab delimited string of child filter ids */ 
	@psCalculations 			varchar(MAX), 
	@psHiddenGroups 			varchar(MAX), 
	@psErrorMsg					varchar(MAX)	OUTPUT,
	@piErrorCode				varchar(MAX)	OUTPUT, /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								3 = definition changed by someone else, overwrite ? */
	@psDeletedCalcs 			varchar(MAX)	OUTPUT, 
	@psHiddenCalcs 				varchar(MAX)	OUTPUT,
	@psDeletedFilters 			varchar(MAX)	OUTPUT,
	@psHiddenFilters 			varchar(MAX)	OUTPUT,
	@psDeletedOrders			varchar(MAX)	OUTPUT,
	@psJobIDsToHide				varchar(MAX)	OUTPUT,
	@psDeletedPicklists 		varchar(MAX)	OUTPUT,
	@psHiddenPicklists 			varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp	integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,
			@sTemp					varchar(MAX),
			@sCurrentID				varchar(100),
			@sParameter				varchar(MAX),
			@sExprName  			varchar(255),
			@sBatchJobName			varchar(255),
			@iBatchJobID			integer,
			@iBatchJobScheduled		integer,
			@sBatchJobRoleToPrompt	varchar(MAX),
			@iNonHiddenCount		integer,
			@sBatchJobUserName		sysname,
			@sJobName				varchar(255),
			@sCurrentUserGroup		sysname,
			@fBatchJobsOK			bit,
			@sScheduledUserGroups	varchar(MAX),
			@sScheduledJobDetails	varchar(MAX),
			@sCurrentUserAccess		varchar(MAX),
			@iOwnedJobCount			integer,
			@sOwnedJobDetails		varchar(MAX),
			@sOwnedJobIDs			varchar(MAX),
			@sNonOwnedJobDetails	varchar(MAX),
			@sHiddenGroupsList		varchar(MAX),
			@sHiddenGroup			varchar(MAX),
			@fSysSecMgr				bit,
			@sActualUserName		sysname,
			@iUserGroupID			integer;

	SET @fBatchJobsOK = 1
	SET @sScheduledUserGroups = ''
	SET @sScheduledJobDetails = ''
	SET @iOwnedJobCount = 0
	SET @sOwnedJobDetails = ''
	SET @sOwnedJobIDs = ''
	SET @sNonOwnedJobDetails = ''

	SELECT @sCurrentUser = SYSTEM_USER
	SET @psErrorMsg = ''
	SET @piErrorCode = 0
	SET @psDeletedCalcs = ''
	SET @psHiddenCalcs = ''
	SET @psDeletedOrders = ''
	SET @psDeletedFilters = ''
	SET @psHiddenFilters = ''
	SET @psDeletedPicklists = ''
	SET @psHiddenPicklists = ''

	EXEC spASRIntSysSecMgr @fSysSecMgr OUTPUT
	
 	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCustomReportsName
		WHERE ID = @piUtilID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The report has been deleted by another user. Save as a new definition ?'
			SET @piErrorCode = 2
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysCustomReportsName
			WHERE ID = @piUtilID

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					2, 
					@piUtilID,
					@sAccess	OUTPUT

				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The report has been amended by another user and is now Read Only. Save as a new definition ?'
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The report has been amended by another user. Would you like to overwrite this definition ?'
					SET @piErrorCode = 3
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the report name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCustomReportsName
			WHERE name = @psUtilName
				AND ID <> @piUtilID
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSCustomReportsName
			WHERE name = @psUtilName
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'A report called ''' + @psUtilName + ''' already exists.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piBasePicklistID > 0)
	BEGIN
		/* Check that the Base table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piBasePicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piBasePicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piBasePicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piBasePicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseFilterID > 0)
	BEGIN
		/* Check that the Base table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piBaseFilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piBaseFilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piBaseFilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piBaseFilterID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piEmailGroupID > 0)
	BEGIN
		/* Check that the email group exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysEmailGroupName 
		WHERE emailGroupID = @piEmailGroupID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The email group has been deleted by another user.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piParent1PicklistID > 0)
	BEGIN
		/* Check that the Parent1 table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piParent1PicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The first parent table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent1PicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piParent1PicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The first parent table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent1PicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent1FilterID > 0)
	BEGIN
		/* Check that the Parent 1 table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piParent1FilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The parent 1 filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent1FilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piParent1FilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The parent 1 table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent1FilterID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent2PicklistID > 0)
	BEGIN
		/* Check that the Parent1 table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piParent2PicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The second parent table picklist has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedPicklists = @psDeletedPicklists +
			CASE
				WHEN LEN(@psDeletedPicklists) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent2PicklistID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piParent2PicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The second parent table picklist has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenPicklists = @psHiddenPicklists +
				CASE
					WHEN LEN(@psHiddenPicklists) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent2PicklistID)
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piParent2FilterID > 0)
	BEGIN
		/* Check that the Parent 2 table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piParent2FilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The parent 2 filter has been deleted by another user, and will be automatically removed from the report.'
			SET @piErrorCode = 1

			SET @psDeletedFilters = @psDeletedFilters +
			CASE
				WHEN LEN(@psDeletedFilters) > 0 THEN ','
				ELSE ''
			END + convert(varchar(100), @piParent2FilterID)
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piParent2FilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The parent 2 table filter has been made hidden by another user, and will be automatically removed from the report.'
				SET @piErrorCode = 1

				SET @psHiddenFilters = @psHiddenFilters +
				CASE
					WHEN LEN(@psHiddenFilters) > 0 THEN ','
					ELSE ''
				END + convert(varchar(100), @piParent2FilterID)
			END
		END
	END

	/* Check that the selected child filters exist and are not hidden. */
	IF (@piErrorCode = 0) AND (LEN(@piChildFilterID) > 0)
	BEGIN
		SET @sTemp = @piChildFilterID

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(char(9), @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			IF @sCurrentID > 0 
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM ASRSysExpressions
				WHERE exprID = convert(integer, @sCurrentID)

				IF @iCount = 0
				BEGIN
					SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedFilters) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							 'One or more of the child filters have been deleted by another user. They will be automatically removed from the report.'
					END
					SET @psDeletedFilters = @psDeletedFilters +
					CASE
						WHEN LEN(@psDeletedFilters) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
					SET @piErrorCode = 1
			 	END
				ELSE
			  	BEGIN
					SELECT @sOwner = userName,
						@sAccess = access
					FROM ASRSysExpressions
					WHERE exprID = convert(integer, @sCurrentID)

					IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
					BEGIN
						SET @psErrorMsg = 
							@psErrorMsg + 
							CASE
								WHEN LEN(@psHiddenFilters) > 0 THEN ''
								ELSE 
									CASE 
										WHEN LEN(@psErrorMsg) > 0 THEN char(13)
										ELSE ''
									END +
									'One or more of the child filters have been made hidden by another user. They will be automatically removed from the report.'
							END
						SET @psHiddenFilters = @psHiddenFilters +
						CASE
							WHEN LEN(@psHiddenFilters) > 0 THEN ','
							ELSE ''
						END + @sCurrentID
						
						SET @piErrorCode = 1
					END
			  	END
			END
		END
	END

	/* Check that the selected child filters exist and are not hidden. */
	IF (@piErrorCode = 0) AND (LEN(@psDeletedOrders) > 0)
	BEGIN
		SET @sTemp = @psDeletedOrders

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(char(9), @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			IF @sCurrentID > 0 
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM ASRSysOrders
				WHERE OrderID = convert(integer, @sCurrentID)

				IF @iCount = 0
				BEGIN
					SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedOrders) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							 'One or more of the child orders have been deleted by another user. They will be automatically removed from the report.'
					END
					SET @psDeletedOrders = @psDeletedOrders +
					CASE
						WHEN LEN(@psDeletedOrders) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
					SET @piErrorCode = 1
			 	END
			END
		END
	END
	
	/* Check that the selected runtime calculations exists. */
	IF (@piErrorCode = 0) AND (LEN(@psCalculations) > 0)
	BEGIN
		SET @sTemp = @psCalculations

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(',', @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(',', @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(',', @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			SELECT @iCount = COUNT(*)
			FROM ASRSysExpressions
			 WHERE exprID = convert(integer, @sCurrentID)

			IF @iCount = 0
			BEGIN
				SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							'One or more runtime calculations have been deleted by another user. They will be automatically removed from the report.'
					END
				SET @psDeletedCalcs = @psDeletedCalcs +
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
				SET @piErrorCode = 1
			END
			ELSE
			BEGIN
				SELECT @sOwner = userName,
					@sAccess = access
				FROM ASRSysExpressions
				WHERE exprID = convert(integer, @sCurrentID)

				IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @psErrorMsg = 
						@psErrorMsg + 
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ''
							ELSE 
								CASE 
									WHEN LEN(@psErrorMsg) > 0 THEN char(13)
									ELSE ''
								END +
								'One or more runtime calculations have been made hidden by another user. They will be automatically removed from the report.'
						END
					SET @psHiddenCalcs = @psHiddenCalcs +
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ','
							ELSE ''
						END + @sCurrentID
						
					SET @piErrorCode = 1
				END
			END
		END
	END
	
	IF (@piErrorCode = 0) AND (@piUtilID > 0) AND (len(@psHiddenGroups) > 0)
	BEGIN
		SELECT @sOwner = userName
		FROM ASRSysCustomReportsName
		WHERE ID = @piUtilID

		IF (@sOwner = @sCurrentUser) 
		BEGIN
			EXEC spASRIntGetActualUserDetails
				@sActualUserName OUTPUT,
				@sCurrentUserGroup OUTPUT,
				@iUserGroupID OUTPUT

			DECLARE @HiddenGroups TABLE(groupName sysname, groupID integer)
			SET @sHiddenGroupsList = substring(@psHiddenGroups, 2, len(@psHiddenGroups)-2)
			WHILE LEN(@sHiddenGroupsList) > 0
			BEGIN
				IF CHARINDEX(char(9), @sHiddenGroupsList) > 0
				BEGIN
					SET @sHiddenGroup = LEFT(@sHiddenGroupsList, CHARINDEX(char(9), @sHiddenGroupsList) - 1)
					SET @sHiddenGroupsList = RIGHT(@sHiddenGroupsList, LEN(@sHiddenGroupsList) - CHARINDEX(char(9), @sHiddenGroupsList))
				END
				ELSE
				BEGIN
					SET @sHiddenGroup = @sHiddenGroupsList
					SET @sHiddenGroupsList = ''
				END

				INSERT INTO @HiddenGroups (groupName, groupID) (SELECT @sHiddenGroup, uid FROM sysusers WHERE name = @sHiddenGroup)
			END

			DECLARE batchjob_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
				ASRSysBatchJobName.Username,
				ASRSysCustomReportsName.Name AS 'JobName'
	 		FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID 
			INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsName.ID = ASRSysBatchJobDetails.JobID
			LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.access <> 'HD'
				AND ASRSysBatchJobAccess.groupName IN (SELECT name FROM sysusers WHERE uid IN (SELECT groupID FROM @HiddenGroups))
				AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysBatchJobDetails.JobType = 'Custom Report'
				AND ASRSysBatchJobDetails.JobID IN (@piUtilID)
			GROUP BY ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				ASRSysBatchJobName.Username,
				ASRSysCustomReportsName.Name

			OPEN batchjob_cursor
			FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
				@iBatchJobID,
				@iBatchJobScheduled,
				@sBatchJobRoleToPrompt,
				@iNonHiddenCount,
				@sBatchJobUserName,
				@sJobName	
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @sCurrentUserAccess = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
								OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		 						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
						WHEN ASRSysBatchJobName.userName = system_user THEN 'RW'
						ELSE
							CASE
								WHEN ASRSysBatchJobAccess.access IS null THEN 'HD'
								ELSE ASRSysBatchJobAccess.access
							END
					END 
				FROM sysusers b
				INNER JOIN sysusers a ON b.uid = a.gid
				LEFT OUTER JOIN ASRSysBatchJobAccess ON (b.name = ASRSysBatchJobAccess.groupName
					AND ASRSysBatchJobAccess.id = @iBatchJobID)
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobAccess.ID = ASRSysBatchJobName.ID
				WHERE a.Name = @sActualUserName

				IF @sBatchJobUserName = @sOwner
				BEGIN
					/* Found a Batch Job whose owner is the same. */
					IF (@iBatchJobScheduled = 1) AND
						(len(@sBatchJobRoleToPrompt) > 0) AND
						(@sBatchJobRoleToPrompt <> @sCurrentUserGroup) AND
						(CHARINDEX(char(9) + @sBatchJobRoleToPrompt + char(9), @psHiddenGroups) > 0)
					BEGIN
						/* Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sBatchJobRoleToPrompt + '<BR>'

						IF @sCurrentUserAccess = 'HD'
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0 
						BEGIN
							SET @iOwnedJobCount = @iOwnedJobCount + 1
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains Custom Report ' + @sJobName + ')' + '<BR>'
							SET @sOwnedJobIDs = @sOwnedJobIDs +
								CASE 
									WHEN Len(@sOwnedJobIDs) > 0 THEN ', '
									ELSE ''
								END +  convert(varchar(100), @iBatchJobID)
						END
					END
				END			
				ELSE
				BEGIN
					/* Found a Batch Job whose owner is not the same. */
					SET @fBatchJobsOK = 0
	    
					IF @sCurrentUserAccess = 'HD'
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
					END
				END

				FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
					@iBatchJobID,
					@iBatchJobScheduled,
					@sBatchJobRoleToPrompt,
					@iNonHiddenCount,
					@sBatchJobUserName,
					@sJobName	
			END
			CLOSE batchjob_cursor
			DEALLOCATE batchjob_cursor	

		END
	END

	IF @fBatchJobsOK = 0
	BEGIN
		SET @piErrorCode = 1

		IF Len(@sScheduledJobDetails) > 0 
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden from the following user groups :'  + '<BR><BR>' +
				@sScheduledUserGroups  +
				'<BR>as it is used in the following batch jobs which are scheduled to be run by these user groups :<BR><BR>' +
				@sScheduledJobDetails
		END
		ELSE
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden as it is used in the following batch jobs of which you are not the owner :<BR><BR>' +
				@sNonOwnedJobDetails
	      	END
	END
	ELSE
	BEGIN
	    	IF (@iOwnedJobCount > 0) 
		BEGIN
			SET @piErrorCode = 4
			SET @psErrorMsg = 'Making this definition hidden to user groups will automatically make the following definition(s), of which you are the owner, hidden to the same user groups:<BR><BR>' +
				@sOwnedJobDetails + '<BR><BR>' +
				'Do you wish to continue ?'
		END
	END

	SET @psJobIDsToHide = @sOwnedJobIDs
	
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntValidateMailMerge]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntValidateMailMerge];
GO
CREATE PROCEDURE [dbo].[spASRIntValidateMailMerge] (
	@psUtilName 		varchar(255), 
	@piUtilID 			integer, 
	@piTimestamp 		integer, 
	@piBasePicklistID	integer, 
	@piBaseFilterID 	integer, 
	@psCalculations 	varchar(MAX), 
	@psHiddenGroups 	varchar(MAX), 
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT, /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								3 = definition changed by someone else, overwrite ?
								4 = saving will cause batch jobs to be made hiiden. Prompt to continue */
	@psDeletedCalcs 	varchar(MAX)	OUTPUT, 
	@psHiddenCalcs 		varchar(MAX)	OUTPUT,
	@psJobIDsToHide		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp				integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,
			@sTemp					varchar(MAX),
			@sCurrentID				varchar(MAX),
			@sBatchJobName			varchar(255),
			@iBatchJobID			integer,
			@iBatchJobScheduled		integer,
			@sBatchJobRoleToPrompt	varchar(MAX),
			@iNonHiddenCount		integer,
			@sBatchJobUserName		sysname,
			@sJobName				varchar(255),
			@sCurrentUserGroup		sysname,
			@fBatchJobsOK			bit,
			@sScheduledUserGroups	varchar(MAX),
			@sScheduledJobDetails	varchar(MAX),
			@sCurrentUserAccess		varchar(MAX),
			@iOwnedJobCount 		integer,
			@sOwnedJobDetails		varchar(MAX),
			@sOwnedJobIDs			varchar(MAX),
			@sNonOwnedJobDetails	varchar(MAX),
			@sHiddenGroupsList		varchar(MAX),
			@sHiddenGroup			varchar(MAX),
			@fSysSecMgr				bit,
			@sActualUserName		sysname,
			@iUserGroupID			integer;

	SET @fBatchJobsOK = 1
	SET @sScheduledUserGroups = ''
	SET @sScheduledJobDetails = ''
	SET @iOwnedJobCount = 0
	SET @sOwnedJobDetails = ''
	SET @sOwnedJobIDs = ''
	SET @sNonOwnedJobDetails = ''

	SELECT @sCurrentUser = SYSTEM_USER
	SET @psErrorMsg = ''
	SET @piErrorCode = 0
	SET @psDeletedCalcs = ''
	SET @psHiddenCalcs = ''

	exec spASRIntSysSecMgr @fSysSecMgr OUTPUT
	
	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysMailMergeName
		WHERE MailMergeID = @piUtilID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The mail merge has been deleted by another user. Save as a new definition ?'
			SET @piErrorCode = 2
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysMailMergeName
			WHERE MailMergeID = @piUtilID

			IF (@iTimestamp <> @piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					9, 
					@piUtilID,
					@sAccess	OUTPUT
		
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The mail merge has been amended by another user and is now Read Only. Save as a new definition ?'
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The mail merge has been amended by another user. Would you like to overwrite this definition ?'
					SET @piErrorCode = 3
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the report name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSMailMergeName
			WHERE name = @psUtilName
				AND MailMergeID <> @piUtilID AND IsLabel = 0
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSYSMailMergeName
			WHERE name = @psUtilName AND IsLabel = 0
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'A mail merge called ''' + @psUtilName + ''' already exists.'
			SET @piErrorCode = 1
		END
	END

	IF (@piErrorCode = 0) AND (@piBasePicklistID > 0)
	BEGIN
		/* Check that the Base table picklist exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysPicklistName 
		WHERE picklistID = @piBasePicklistID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table picklist has been deleted by another user.'
			SET @piErrorCode = 1
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysPicklistName 
			WHERE picklistID = @piBasePicklistID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table picklist has been made hidden by another user.'
				SET @piErrorCode = 1
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseFilterID > 0)
	BEGIN
		/* Check that the Base table filter exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions 
		WHERE exprID = @piBaseFilterID

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base table filter has been deleted by another user.'
			SET @piErrorCode = 1
		END
		ELSE
		BEGIN
			SELECT @sOwner = userName,
				@sAccess = access
			FROM ASRSysExpressions 
			WHERE exprID = @piBaseFilterID

			IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @psErrorMsg = 'The base table filter has been made hidden by another user.'
				SET @piErrorCode = 1
			END
		END
	END

	/* Check that the selected runtime calculations exists. */
	IF (@piErrorCode = 0) AND (LEN(@psCalculations) > 0)
	BEGIN
		SET @sTemp = @psCalculations

		WHILE LEN(@sTemp) > 0
		BEGIN
			IF CHARINDEX(',', @sTemp) > 0
			BEGIN
				SET @sCurrentID = LEFT(@sTemp, CHARINDEX(',', @sTemp) - 1)
				SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(',', @sTemp))
			END
			ELSE
			BEGIN
				SET @sCurrentID = @sTemp
				SET @sTemp = ''
			END
			
			SELECT @iCount = COUNT(*)
			FROM ASRSysExpressions
			 WHERE exprID = convert(integer, @sCurrentID)

			IF @iCount = 0
			BEGIN
				SET @psErrorMsg = 
					@psErrorMsg + 
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ''
						ELSE 
							CASE 
								WHEN LEN(@psErrorMsg) > 0 THEN char(13)
								ELSE ''
							END +
							'One or more runtime calculations have been deleted by another user. They will be automatically removed from the mail merge.'
					END
				SET @psDeletedCalcs = @psDeletedCalcs +
					CASE
						WHEN LEN(@psDeletedCalcs) > 0 THEN ','
						ELSE ''
					END + @sCurrentID
				SET @piErrorCode = 1
			END
			ELSE
			BEGIN
				SELECT @sOwner = userName,
					@sAccess = access
				FROM ASRSysExpressions
				WHERE exprID = convert(integer, @sCurrentID)

				IF (@sOwner <> @sCurrentUser) AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
				BEGIN
					SET @psErrorMsg = 
						@psErrorMsg + 
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ''
							ELSE 
								CASE 
									WHEN LEN(@psErrorMsg) > 0 THEN char(13)
									ELSE ''
								END +
								'One or more runtime calculations have been made hidden by another user. They will be automatically removed from the mail merge.'
						END
					SET @psHiddenCalcs = @psHiddenCalcs +
						CASE
							WHEN LEN(@psHiddenCalcs) > 0 THEN ','
							ELSE ''
						END + @sCurrentID
						
					SET @piErrorCode = 1
				END
			END
		END
	END

	IF (@piErrorCode = 0) AND (@piUtilID > 0) AND (len(@psHiddenGroups) > 0)
	BEGIN
		SELECT @sOwner = userName
		FROM ASRSysMailMergeName
		WHERE MailMergeID = @piUtilID

		IF (@sOwner = @sCurrentUser) 
		BEGIN
			EXEC spASRIntGetActualUserDetails
				@sActualUserName OUTPUT,
				@sCurrentUserGroup OUTPUT,
				@iUserGroupID OUTPUT

			DECLARE @HiddenGroups TABLE(groupName sysname, groupID integer)
			SET @sHiddenGroupsList = substring(@psHiddenGroups, 2, len(@psHiddenGroups)-2)
			WHILE LEN(@sHiddenGroupsList) > 0
			BEGIN
				IF CHARINDEX(char(9), @sHiddenGroupsList) > 0
				BEGIN
					SET @sHiddenGroup = LEFT(@sHiddenGroupsList, CHARINDEX(char(9), @sHiddenGroupsList) - 1)
					SET @sHiddenGroupsList = RIGHT(@sHiddenGroupsList, LEN(@sHiddenGroupsList) - CHARINDEX(char(9), @sHiddenGroupsList))
				END
				ELSE
				BEGIN
					SET @sHiddenGroup = @sHiddenGroupsList
					SET @sHiddenGroupsList = ''
				END

				INSERT INTO @HiddenGroups (groupName, groupID) (SELECT @sHiddenGroup, uid FROM sysusers WHERE name = @sHiddenGroup)
			END

			DECLARE batchjob_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				COUNT (ASRSysBatchJobAccess.Access) AS [nonHiddenCount],
				ASRSysBatchJobName.Username,
				ASRSysMailMergeName.Name AS 'JobName'
	 		FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobName.ID = ASRSysBatchJobDetails.BatchJobNameID 
			INNER JOIN ASRSysMailMergeName ON ASRSysMailMergeName.MailMergeID = ASRSysBatchJobDetails.JobID
			LEFT OUTER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.access <> 'HD'
				AND ASRSysBatchJobAccess.groupName IN (SELECT name FROM sysusers WHERE uid IN (SELECT groupID FROM @HiddenGroups))
				AND ASRSysBatchJobAccess.groupName NOT IN (SELECT sysusers.name
					FROM sysusers
					INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName
						AND ASRSysGroupPermissions.permitted = 1
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE sysusers.uid = sysusers.gid
						AND sysusers.uid <> 0)
			WHERE ASRSysBatchJobDetails.JobType = 'Mail Merge'
				AND ASRSysBatchJobDetails.JobID IN (@piUtilID)
			GROUP BY ASRSysBatchJobName.Name,
				ASRSysBatchJobName.ID,
				convert(integer, ASRSysBatchJobName.scheduled),
				ASRSysBatchJobName.roleToPrompt,
				ASRSysBatchJobName.Username,
				ASRSysMailMergeName.Name

			OPEN batchjob_cursor
			FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
				@iBatchJobID,
				@iBatchJobScheduled,
				@sBatchJobRoleToPrompt,
				@iNonHiddenCount,
				@sBatchJobUserName,
				@sJobName	
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @sCurrentUserAccess = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
								OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		 						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 'RW'
						WHEN ASRSysBatchJobName.userName = system_user THEN 'RW'
						ELSE
							CASE
								WHEN ASRSysBatchJobAccess.access IS null THEN 'HD'
								ELSE ASRSysBatchJobAccess.access
							END
					END 
				FROM sysusers b
				INNER JOIN sysusers a ON b.uid = a.gid
				LEFT OUTER JOIN ASRSysBatchJobAccess ON (b.name = ASRSysBatchJobAccess.groupName
					AND ASRSysBatchJobAccess.id = @iBatchJobID)
				INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobAccess.ID = ASRSysBatchJobName.ID
				WHERE a.Name = @sActualUserName

				IF @sBatchJobUserName = @sOwner
				BEGIN
					/* Found a Batch Job whose owner is the same. */
					IF (@iBatchJobScheduled = 1) AND
						(len(@sBatchJobRoleToPrompt) > 0) AND
						(@sBatchJobRoleToPrompt <> @sCurrentUserGroup) AND
						(CHARINDEX(char(9) + @sBatchJobRoleToPrompt + char(9), @psHiddenGroups) > 0)
					BEGIN
						/* Found a Batch Job which is scheduled for another user group to run. */
						SET @fBatchJobsOK = 0
						SET @sScheduledUserGroups = @sScheduledUserGroups + @sBatchJobRoleToPrompt + '<BR>'

						IF @sCurrentUserAccess = 'HD'
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
						END
						ELSE
						BEGIN
							SET @sScheduledJobDetails = @sScheduledJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
						END
					END
					ELSE
					BEGIN
						IF @iNonHiddenCount > 0 
						BEGIN
							SET @iOwnedJobCount = @iOwnedJobCount + 1
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains Mail Merge ' + @sJobName + ')' + '<BR>'
							SET @sOwnedJobIDs = @sOwnedJobIDs +
								CASE 
									WHEN Len(@sOwnedJobIDs) > 0 THEN ', '
									ELSE ''
								END +  convert(varchar(100), @iBatchJobID)
						END
					END
				END			
				ELSE
				BEGIN
					/* Found a Batch Job whose owner is not the same. */
					SET @fBatchJobsOK = 0
	    
					IF @sCurrentUserAccess = 'HD'
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : <Hidden> by ' + @sBatchJobUserName + '<BR>'
					END
					ELSE
					BEGIN
						SET @sNonOwnedJobDetails = @sNonOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + '<BR>'
					END
				END

				FETCH NEXT FROM batchjob_cursor INTO @sBatchJobName, 
					@iBatchJobID,
					@iBatchJobScheduled,
					@sBatchJobRoleToPrompt,
					@iNonHiddenCount,
					@sBatchJobUserName,
					@sJobName	
			END
			CLOSE batchjob_cursor
			DEALLOCATE batchjob_cursor	

		END
	END

	IF @fBatchJobsOK = 0
	BEGIN
		SET @piErrorCode = 1

		IF Len(@sScheduledJobDetails) > 0 
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden from the following user groups :'  + '<BR><BR>' +
				@sScheduledUserGroups  +
				'<BR>as it is used in the following batch jobs which are scheduled to be run by these user groups :<BR><BR>' +
				@sScheduledJobDetails
		END
		ELSE
		BEGIN
			SET @psErrorMsg = 'This definition cannot be made hidden as it is used in the following batch jobs of which you are not the owner :<BR><BR>' +
				@sNonOwnedJobDetails
	      	END
	END
	ELSE
	BEGIN
	    	IF (@iOwnedJobCount > 0) 
		BEGIN
			SET @piErrorCode = 4
			SET @psErrorMsg = 'Making this definition hidden to user groups will automatically make the following definition(s), of which you are the owner, hidden to the same user groups:<BR><BR>' +
				@sOwnedJobDetails + '<BR><BR>' +
				'Do you wish to continue ?'
		END
	END

	SET @psJobIDsToHide = @sOwnedJobIDs
	
END
GO



GO
DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
		 INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
		OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
		OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
		AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
		IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
		BEGIN
				SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END
		ELSE
		BEGIN
				SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
				EXEC(@sSQL)
		END

		FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects

GO

		
DECLARE @sVersion varchar(10) = '8.1.0'

EXEC spsys_setsystemsetting 'database', 'version', '8.1';
EXEC spsys_setsystemsetting 'intranet', 'version', @sVersion;
EXEC spsys_setsystemsetting 'ssintranet', 'version', @sVersion;
