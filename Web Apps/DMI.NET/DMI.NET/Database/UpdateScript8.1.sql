

---- Drop redundant functions (or renamed)
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetMailMergeDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetMailMergeDefinition];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetEventLogRecords]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetEventLogRecords];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetEventLogBatchDetails]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].spASRIntGetEventLogBatchDetails;
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

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetMessages]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetMessages];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetWorkflowParameters]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetWorkflowParameters];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetUtilityBaseTable]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetUtilityBaseTable];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetUtilityName]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetUtilityName];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetExprFunctions]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetExprFunctions];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntActivateModule]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntActivateModule];
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetFindRecords3]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetFindRecords3];
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_ASRIntGetSystemSetting]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetSystemSetting]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_ASRIntGetSetting]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetSetting]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_ASRIntGetUtilityPromptedValues]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetUtilityPromptedValues]
GO




-- Functions we do want to keep

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntPopulateDefsel]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntPopulateDefsel];
GO

CREATE PROCEDURE [dbo].[sp_ASRIntPopulateDefsel] (
	@intType int, 
	@blnOnlyMine bit,
	@intTableID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the details with which to populate the intranet defsel grid. */
	DECLARE 
		@strSQL 			nvarchar(MAX),
		@strExplicitSQL 	varchar(MAX),
		@strTableName		varchar(255),
		@strIDName 			varchar(255),
		@sExtraWhereSQL		varchar(MAX),
		@fNewAccess			bit,
		@sRecordSourceWhere	varchar(MAX),
		@sAccessTableName	varchar(255),
		@sRoleName			varchar(255),
		@fSysSecMgr			bit,
		@fDoneWhere			bit,
		@sActualUserName	varchar(250),
		@iActualUserGroupID	integer

	SET @fNewAccess = 0;
	SET @sExtraWhereSQL = '';
	SET @fDoneWhere = 0;
	SET @strExplicitSQL = '';
	
	IF ((@intTableID <=0) OR (@intTableID IS null)) AND (@intType <> 17) AND (@intType <> 9)
	BEGIN
		/* No table ID passed in, so use the first table alphabetically. */
		SELECT TOP 1 @intTableID = tableID
		FROM [dbo].[ASRSysTables]
		ORDER BY tableName;
	END

	IF @intType = 1 /*'crosstabs'*/
	BEGIN
		SET @strTableName = 'ASRSysCrossTab';
		SET @strIDName = 'CrossTabID';
		SET @fNewAccess = 1;
		SET @sAccessTableName= 'ASRSysCrossTabAccess';
		SET @sExtraWhereSQL = ' CrossTabType = 0';
	END

	IF @intType = 2 /*'customreports'*/
	BEGIN
		SET @strTableName = 'ASRSysCustomReportsName';
		SET @strIDName = 'ID';
		SET @fNewAccess = 1;
		SET @sAccessTableName= 'ASRSysCustomReportAccess';
	END

	IF @intType = 9 /*'mailmerge'*/
	BEGIN
		SET @strTableName = 'ASRSysMailMergeName';
		SET @strIDName = 'MailMergeID';
		SET @fNewAccess = 1;
		SET @sRecordSourceWhere = 'ASRSysMailMergeName.IsLabel = 0';
		SET @sAccessTableName= 'ASRSysMailMergeAccess';
		if (@intTableID > 0)
		BEGIN
			SET @sExtraWhereSQL = 'ASRSysMailMergeName.TableID = ' + convert(varchar(255), @intTableID);
		END
	END

	IF @intType = 10 /*'picklists'*/
	BEGIN
		SET @strTableName = 'ASRSysPickListName';
		SET @strIDName = 'picklistID';
		SET @sExtraWhereSQL = ' TableID = ' + convert(varchar(255), @intTableID);
	END

	IF @intType = 11 /*'filters'*/
	BEGIN
		SET @strTableName = 'ASRSysExpressions';
		SET @strIDName = 'exprID';
		SET @sExtraWhereSQL = ' type = 11 AND (returnType = 3 OR type = 10) AND parentComponentID = 0	AND TableID = ' + convert(varchar(255), @intTableID);
	END

	IF @intType = 12 /*'calculations'*/
	BEGIN
		SET @strTableName = 'ASRSysExpressions';
		SET @strIDName = 'exprID';
		SET @sExtraWhereSQL = ' type = 10 AND (returnType = 0 OR type = 10) AND parentComponentID = 0	AND TableID = ' + convert(varchar(255), @intTableID);
	END
	
	IF @intType = 17 /*'calendarreports'*/
	BEGIN
		SET @strTableName = 'ASRSysCalendarReports';
		SET @strIDName = 'ID';
		SET @fNewAccess = 1;
		SET @sAccessTableName= 'ASRSysCalendarReportAccess';
		if (@intTableID > 0)
		BEGIN
			SET @sExtraWhereSQL = 'ASRSysCalendarReports.BaseTable = ' + convert(varchar(255), @intTableID);
		END
	END

	IF @intType = 25 /*'workflow'*/
	BEGIN
		SET @strExplicitSQL = 'SELECT 
			Name, 
			replace(ASRSysWorkflows.description, char(9), '''') AS [description],
			'''' AS [Username],
			''rw'' AS [Access],
			ID
			FROM ASRSysWorkflows
			WHERE ASRSysWorkflows.enabled = 1
				AND ISNULL(ASRSysWorkflows.initiationType, 0) = 0
			ORDER BY ASRSysWorkflows.name';
	END

	IF @intType = 35 /*'nineboxgridreport'*/
	BEGIN
		SET @strTableName = 'AsrSysCrossTab';
		SET @strIDName = 'CrossTabID';
		SET @fNewAccess = 1;
		SET @sAccessTableName= 'ASRSysCrossTabAccess';
		SET @sExtraWhereSQL = ' CrossTabType = 4';
	END
		
	IF len(@strExplicitSQL) > 0 
	BEGIN
		SET @strSQL = @strExplicitSQL;
	END
	ELSE
	BEGIN
		EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sRoleName OUTPUT,
			@iActualUserGroupID OUTPUT;

		SELECT @fSysSecMgr = 
			CASE
				WHEN (SELECT count(*)
					FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
						AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
						OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))
					INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')
					WHERE ASRSysGroupPermissions.groupname = @sRoleName
						AND ASRSysGroupPermissions.permitted = 1) > 0 THEN 1
				ELSE 0
			END;
			
		IF @fNewAccess = 1
		BEGIN
			SET @strSQL = 'SELECT ' + @strTableName + '.Name, ' +
				'replace(' + @strTableName + '.Description, char(9), '''') AS [description], ' +
				'lower(' +@strTableName + '.Username) as ''Username'', ';
				
			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					'lower(' + @sAccessTableName + '.Access) as ''Access'', ';
			END
			ELSE
			BEGIN
				SET @strSQL = @strSQL +
					'''rw'' as ''Access'', ';
			END
								
			SET @strSQL = @strSQL +
				@strTableName + '.' + @strIDName + '  as ''ID'' 
				FROM ' + @strTableName + 
				' INNER JOIN ' + @sAccessTableName + ' ON ' + @strTableName + '.' + @strIDName +  ' = ' + @sAccessTableName + '.ID
				AND ' + @sAccessTableName + '.groupname = ''' + @sRoleName + '''';

			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					 ' WHERE ([Username] = SYSTEM_USER';
				IF @blnOnlyMine = 0 SET @strSQL = @strSQL + ' OR [Access] <> ''HD''';
				SET @strSQL = @strSQL + ')';
				SET @fDoneWhere = 1;
			END
			ELSE
			BEGIN
				IF @blnOnlyMine = 1
				BEGIN
					SET @strSQL = @strSQL +
						 ' WHERE ([Username] = SYSTEM_USER)';
					SET @fDoneWhere = 1;
				END
			END

			IF LEN(@sRecordSourceWhere) > 0 
			BEGIN
				IF @fDoneWhere = 0
				BEGIN
					SET @strSQL = @strSQL  + ' WHERE';
					SET @fDoneWhere = 1;
				END
				ELSE
				BEGIN
					SET @strSQL = @strSQL  + ' AND';
				END

				SET @strSQL = @strSQL  + ' (' + @sRecordSourceWhere + ')';
			END
			
			IF LEN(@sExtraWhereSQL) > 0 
			BEGIN
				IF @fDoneWhere = 0
				BEGIN
					SET @strSQL = @strSQL  + ' WHERE';
					SET @fDoneWhere = 1;
				END
				ELSE
				BEGIN
					SET @strSQL = @strSQL  + ' AND';
				END

				SET @strSQL = @strSQL  + ' (' + @sExtraWhereSQL + ')';
			END
			
			SET @strSQL = @strSQL + ' ORDER BY ' + @strTableName + '.Name';
		END
		ELSE
		BEGIN
			SET @strSQL = 'SELECT Name, replace(Description, char(9), '''') AS [description], lower(Username) as ''Username'', ';
			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					'lower(Access) as ''Access'', ';
			END
			ELSE
			BEGIN
				SET @strSQL = @strSQL +
					'''rw'' as ''Access'', ';
			END
			SET @strSQL = @strSQL +
				@strIDName + '  as ''ID'' FROM ' + @strTableName;

			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					 ' WHERE ([Username] = SYSTEM_USER';
				IF @blnOnlyMine = 0 SET @strSQL = @strSQL + ' OR [Access] <> ''HD''';
				SET @strSQL = @strSQL + ')';
				SET @fDoneWhere = 1;
			END
			ELSE
			BEGIN
				IF @blnOnlyMine = 1
				BEGIN
					SET @strSQL = @strSQL +
						 ' WHERE ([Username] = SYSTEM_USER)';
					SET @fDoneWhere = 1;
				END
			END
			
			IF LEN(@sRecordSourceWhere) > 0 
			BEGIN
				IF @fDoneWhere = 0
				BEGIN
					SET @strSQL = @strSQL  + ' WHERE';
					SET @fDoneWhere = 1;
				END
				ELSE
				BEGIN
					SET @strSQL = @strSQL  + ' AND';
				END

				SET @strSQL = @strSQL  + ' (' + @sRecordSourceWhere + ')';
			END
			
			IF LEN(@sExtraWhereSQL) > 0 
			BEGIN
				IF @fDoneWhere = 0
				BEGIN
					SET @strSQL = @strSQL  + ' WHERE';
					SET @fDoneWhere = 1;
				END
				ELSE
				BEGIN
					SET @strSQL = @strSQL  + ' AND';
				END

				SET @strSQL = @strSQL  + ' (' + @sExtraWhereSQL + ')';
			END

			SET @strSQL = @strSQL + ' ORDER BY Name';
		END
	END
	
	/* Return the resultset. */
	EXECUTE sp_executeSQL @strSQL;
	
END

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
	SELECT @psReportName AS [Name], [description], userName AS [owner],		
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

	
IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntDeleteCheck]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntDeleteCheck];
GO

CREATE PROCEDURE [dbo].[spASRIntDeleteCheck] (
	@piUtilityType	integer,
	@plngID			integer,
	@pfDeleted		bit				OUTPUT,
	@psAccess		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@sTableName			sysname,
		@sAccessTableName	sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@fNewAccess			bit,
		@iCount				integer,
		@sAccess			varchar(MAX),
		@fSysSecMgr			bit;

	SET @sTableName = '';
	SET @psAccess = 'HD';
	SET @pfDeleted = 0;
	SET @fNewAccess = 0;

	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sAccessTableName = 'ASRSysBatchJobAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sAccessTableName = 'ASRSysCalendarReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END

	IF @piUtilityType = 1 OR @piUtilityType = 35 /* Cross Tab or 9-Box Grid*/
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sAccessTableName = 'ASRSysCrossTabAccess';
		SET @sIDColumnName = 'CrossTabID';
		SET @fNewAccess = 1;
 	END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sAccessTableName = 'ASRSysCustomReportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
 	END
    
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sAccessTableName = 'ASRSysDataTransferAccess';
		SET @sIDColumnName = 'DataTransferID';
		SET @fNewAccess = 1;
  END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sAccessTableName = 'ASRSysExportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sAccessTableName = 'ASRSysGlobalAccess';
		SET @sIDColumnName = 'functionID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sAccessTableName = 'ASRSysImportAccess';
		SET @sIDColumnName = 'ID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sAccessTableName = 'ASRSysMailMergeAccess';
		SET @sIDColumnName = 'mailMergeID';
		SET @fNewAccess = 1;
  END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sAccessTableName = 'ASRSysRecordProfileAccess';
		SET @sIDColumnName = 'recordProfileID';
		SET @fNewAccess = 1
  END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sAccessTableName = 'ASRSysMatchReportAccess';
		SET @sIDColumnName = 'matchReportID';
		SET @fNewAccess = 1;
  END

	IF (@piUtilityType = 11) OR (@piUtilityType = 12)  /* Filters/Calcs */
	BEGIN
		SET @sTableName = 'ASRSysExpressions';
		SET @sIDColumnName = 'exprID';
  END

	IF (@piUtilityType = 10)  /* Picklists */
	BEGIN
		SET @sTableName = 'ASRSysPicklistName';
		SET @sIDColumnName = 'picklistID';
  END

	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @iCount = COUNT(*)
				FROM ' + @sTableName + 
				' WHERE ' + @sTableName + '.' + @sIDColumnName + ' = ' + convert(nvarchar(255), @plngID);
		SET @sParamDefinition = N'@iCount integer OUTPUT';
		EXEC sp_executesql @sSQL,  @sParamDefinition, @iCount OUTPUT;

		IF @iCount = 0 
		BEGIN
			SET @pfDeleted = 1;
		END
		ELSE
		BEGIN
			IF @fNewAccess = 1
			BEGIN
				exec [dbo].[spASRIntCurrentUserAccess] @piUtilityType,	@plngID, @psAccess OUTPUT;
			END
			ELSE
			BEGIN
				exec [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
				
				IF @fSysSecMgr = 1 
				BEGIN
					SET @psAccess = 'RW';
				END
				ELSE
				BEGIN
					SET @sSQL = 'SELECT @sAccess = CASE 
								WHEN userName = system_user THEN ''RW''
								ELSE access
							END
							FROM ' + @sTableName + 
							' WHERE ' + @sTableName + '.' + @sIDColumnName + ' = ' + convert(nvarchar(255), @plngID);
					SET @sParamDefinition = N'@sAccess varchar(MAX) OUTPUT';
					EXEC sp_executesql @sSQL,  @sParamDefinition, @sAccess OUTPUT;

					SET @psAccess = @sAccess;
				END
			END
		END
	END
END

GO

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntDeleteUtility]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntDeleteUtility];
GO

CREATE PROCEDURE [dbo].[sp_ASRIntDeleteUtility] (
	@piUtilType	integer,
	@piUtilID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iExprID	integer;

	IF @piUtilType = 0
	BEGIN
		/* Batch Jobs */
		DELETE FROM ASRSysBatchJobName WHERE ID = @piUtilID;
		DELETE FROM ASRSysBatchJobDetails WHERE BatchJobNameID = @piUtilID;
		DELETE FROM ASRSysBatchJobAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 1 OR @piUtilType = 35
	BEGIN
		/* Cross Tabs or 9-Box Grid*/
		DELETE FROM ASRSysCrossTab WHERE CrossTabID = @piUtilID;
		DELETE FROM ASRSysCrossTabAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 2
	BEGIN
		/* Custom Reports. */
		DELETE FROM ASRSysCustomReportsName WHERE id = @piUtilID;
		DELETE FROM ASRSysCustomReportsDetails WHERE customReportID= @piUtilID;
		DELETE FROM ASRSysCustomReportAccess WHERE ID = @piUtilID;
	END
	
	IF @piUtilType = 3
	BEGIN
		/* Data Transfer. */
		DELETE FROM ASRSysDataTransferName WHERE DataTransferID = @piUtilID;
		DELETE FROM ASRSysDataTransferAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 4
	BEGIN
		/* Export. */
		DELETE FROM ASRSysExportName WHERE ID = @piUtilID;
		DELETE FROM ASRSysExportDetails WHERE ExportID = @piUtilID;
		DELETE FROM ASRSysExportAccess WHERE ID = @piUtilID;
	END

	IF (@piUtilType = 5) OR (@piUtilType = 6) OR (@piUtilType = 7)
	BEGIN
		/* Globals. */
		DELETE FROM ASRSysGlobalFunctions  WHERE FunctionID = @piUtilID;
		DELETE FROM ASRSysGlobalAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 8
	BEGIN
		/* Import. */
		DELETE FROM ASRSysImportName  WHERE ID = @piUtilID;
		DELETE FROM ASRSysImportDetails WHERE ImportID = @piUtilID;
		DELETE FROM ASRSysImportAccess WHERE ID = @piUtilID;
	END

	IF (@piUtilType = 9) OR (@piUtilType = 18)
	BEGIN
		/* Mail Merge/ Envelopes & Labels. */
		DELETE FROM ASRSysMailMergeName  WHERE MailMergeID = @piUtilID;
		DELETE FROM ASRSysMailMergeColumns  WHERE MailMergeID = @piUtilID;
		DELETE FROM ASRSysMailMergeAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 10
	BEGIN
		/* Picklists. */
		DELETE FROM ASRSysPickListName WHERE picklistID = @piUtilID;
		DELETE FROM ASRSysPickListItems WHERE picklistID = @piUtilID;
	END
	
	IF @piUtilType = 11 OR @piUtilType = 12
	BEGIN
		/* Filters and Calculations. */
		DECLARE subExpressions_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysExpressions.exprID
			FROM ASRSysExpressions
			INNER JOIN ASRSysExprComponents ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID
			AND ASRSysExprComponents.exprID = @piUtilID;
		OPEN subExpressions_cursor;
		FETCH NEXT FROM subExpressions_cursor INTO @iExprID;
		WHILE (@@fetch_status = 0)
		BEGIN
			exec [dbo].[sp_ASRIntDeleteUtility] @piUtilType, @iExprID;
			
			FETCH NEXT FROM subExpressions_cursor INTO @iExprID;
		END
		CLOSE subExpressions_cursor;
		DEALLOCATE subExpressions_cursor;

		DELETE FROM ASRSysExprComponents
		WHERE exprID = @piUtilID;

		DELETE FROM ASRSysExpressions WHERE exprID = @piUtilID;
	END	

	IF (@piUtilType = 14) OR (@piUtilType = 23) OR (@piUtilType = 24)
	BEGIN
		/* Match Reports/Succession Planning/Career Progression. */
		DELETE FROM ASRSysMatchReportName WHERE MatchReportID = @piUtilID;
		DELETE FROM ASRSysMatchReportAccess WHERE ID = @piUtilID;
	END

	IF @piUtilType = 17 
	BEGIN
		/*Calendar Reports*/
		DELETE FROM ASRSysCalendarReports WHERE ID = @piUtilID;
		DELETE FROM ASRSysCalendarReportEvents WHERE CalendarReportID = @piUtilID;
		DELETE FROM ASRSysCalendarReportOrder WHERE CalendarReportID = @piUtilID;
		DELETE FROM ASRSysCalendarReportAccess WHERE ID = @piUtilID;
	END
	
	IF @piUtilType = 20 
	BEGIN
		/*Record Profile*/
		DELETE FROM ASRSysRecordProfileName WHERE recordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileDetails WHERE RecordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileTables WHERE RecordProfileID = @piUtilID;
		DELETE FROM ASRSysRecordProfileAccess WHERE ID = @piUtilID;
	END
	
END

GO

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntSaveCrossTab]') AND xtype in (N'P'))
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
			OutputFileName,
			CrossTabType)
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
			@psOutputFilename,
			0 -- Cross tab
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

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetUtilityAccessRecords]') AND xtype in (N'P'))
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

	IF @piUtilityType = 1 OR @piUtilityType = 35
	BEGIN
		/* Cross Tabs or 9-box Grid*/
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

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntValidateNineBoxGrid]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntValidateNineBoxGrid];
GO

CREATE PROCEDURE [dbo].[spASRIntValidateNineBoxGrid] (
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
			SET @psErrorMsg = 'The 9-Box Grid has been deleted by another user. Save as a new definition ?'
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
					SET @psErrorMsg = 'The 9-Box Grid has been amended by another user and is now Read Only. Save as a new definition ?'
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The 9-Box Grid has been amended by another user. Would you like to overwrite this definition ?'
					SET @piErrorCode = 3
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the 9-Box Grid name is unique. */
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
			SET @psErrorMsg = 'A 9-Box Grid called ''' + @psUtilName + ''' already exists.'
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
			WHERE ASRSysBatchJobDetails.JobType = '9-BOX GRID REPORT'
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
							SET @sOwnedJobDetails = @sOwnedJobDetails + 'Batch Job : ' + @sBatchJobName + ' (Contains 9-Box Grid ' + @sJobName + ')' + '<BR>'
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

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntSaveNineBoxGrid]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntSaveNineBoxGrid];
GO

CREATE PROCEDURE [dbo].[spASRIntSaveNineBoxGrid] (
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
	@piVColID			integer,
	@psVStart			varchar(100),
	@psVStop			varchar(100),
	@piPColID			integer,
	@psPStart			varchar(100),
	@psPStop			varchar(100),
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
	@XAxisLabel varchar(255),
	@XAxisSubLabel1 varchar(255),
	@XAxisSubLabel2 varchar(255),
	@XAxisSubLabel3 varchar(255),
	@YAxisLabel varchar(255),
	@YAxisSubLabel1 varchar(255),
	@YAxisSubLabel2 varchar(255),
	@YAxisSubLabel3 varchar(255),
	@Description1 varchar(255),
	@ColorDesc1 varchar(6),
	@Description2 varchar(255),
	@ColorDesc2 varchar(6),
	@Description3 varchar(255),
	@ColorDesc3 varchar(6),
	@Description4 varchar(255),
	@ColorDesc4 varchar(6),
	@Description5 varchar(255),
	@ColorDesc5 varchar(6),
	@Description6 varchar(255),
	@ColorDesc6 varchar(6),
	@Description7 varchar(255),
	@ColorDesc7 varchar(6),
	@Description8 varchar(255),
	@ColorDesc8 varchar(6),
	@Description9 varchar(255),
	@ColorDesc9 varchar(6),
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
			OutputFileName,
			CrossTabType,
			XAxisLabel,
			XAxisSubLabel1,
			XAxisSubLabel2,
			XAxisSubLabel3,
			YAxisLabel,
			YAxisSubLabel1,
			YAxisSubLabel2,
			YAxisSubLabel3,
			Description1,
			ColorDesc1,
			Description2,
			ColorDesc2,
			Description3,
			ColorDesc3,
			Description4,
			ColorDesc4,
			Description5,
			ColorDesc5,
			Description6,
			ColorDesc6,
			Description7,
			ColorDesc7,
			Description8,
			ColorDesc8,
			Description9,
			ColorDesc9)
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
			0,
			@piVColID,
			@psVStart,
			@psVStop,
			0,
			@piPColID,
			@psPStart,
			@psPStop,
			0,
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
			@psOutputFilename,
			4, -- Nine box grid
			@XAxisLabel,
			@XAxisSubLabel1,
			@XAxisSubLabel2,
			@XAxisSubLabel3,
			@YAxisLabel,
			@YAxisSubLabel1,
			@YAxisSubLabel2,
			@YAxisSubLabel3,
			@Description1,
			@ColorDesc1,
			@Description2,
			@ColorDesc2,
			@Description3,
			@ColorDesc3,
			@Description4,
			@ColorDesc4,
			@Description5,
			@ColorDesc5,
			@Description6,
			@ColorDesc6,
			@Description7,
			@ColorDesc7,
			@Description8,
			@ColorDesc8,
			@Description9,
			@ColorDesc9)

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
			VerticalColID = @piVColID,
			VerticalStart = @psVStart,
			VerticalStop = @psVStop,
			PageBreakColID = @piPColID,
			PageBreakStart = @psPStart,
			PageBreakStop = @psPStop,
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
			OutputFileName = @psOutputFilename,
			XAxisLabel = @XAxisLabel,
			XAxisSubLabel1 = @XAxisSubLabel1,
			XAxisSubLabel2 = @XAxisSubLabel2,
			XAxisSubLabel3 = @XAxisSubLabel3,
			YAxisLabel = @YAxisLabel,
			YAxisSubLabel1 = @YAxisSubLabel1,
			YAxisSubLabel2 = @YAxisSubLabel2,
			YAxisSubLabel3 = @YAxisSubLabel3,
			Description1 = @Description1,
			ColorDesc1 = @ColorDesc1,
			Description2 = @Description2,
			ColorDesc2 = @ColorDesc2,
			Description3 = @Description3,
			ColorDesc3 = @ColorDesc3,
			Description4 = @Description4,
			ColorDesc4 = @ColorDesc4,
			Description5 = @Description5,
			ColorDesc5 = @ColorDesc5,
			Description6 = @Description6,
			ColorDesc6 = @ColorDesc6,
			Description7 = @Description7,
			ColorDesc7 = @ColorDesc7,
			Description8 = @Description8,
			ColorDesc8 = @ColorDesc8,
			Description9 = @Description9,
			ColorDesc9 = @ColorDesc9
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
		VALUES (35, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 35

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (35, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 35
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

IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetNineBoxGridDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetNineBoxGridDefinition];
GO

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
			@ColorDesc9 varchar(6);


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

GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetLinks]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetLinks];
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[spASRIntGetLinks] 
(
		@plngTableID	integer,
		@plngViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iCount				integer,
		@iUtilType			integer, 
		@iUtilID			integer,
		@iScreenID			integer,
		@iTableID			integer,
		@sTableName			sysname,
		@iTableType			integer,
		@sRealSource		sysname,
		@iChildViewID		integer,
		@sAccess			varchar(MAX),
		@sGroupName			varchar(255),
		@pfPermitted		bit,
		@sActualUserName	sysname,
		@iActualUserGroupID integer,
		@fBaseTableReadable bit,
		@iBaseTableID		integer,
		@sURL				varchar(MAX), 
		@fUtilOK			bit,
		@fDrillDownHidden bit,
		@iLinkType			integer,		-- 0 = Hypertext, 1 = Button, 2 = Dropdown List
		@iElement_Type		integer;		-- 2 = chart

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
	
	IF @plngViewID < 1 
	BEGIN 
		SET @plngViewID = -1;
	END
	SET @fBaseTableReadable = 1;
	
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA'
	BEGIN
		EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sGroupName OUTPUT,
			@iActualUserGroupID OUTPUT;
		
		DECLARE @Phase1 TABLE([ID] INT);
		INSERT INTO @Phase1
			SELECT Object_ID(ASRSysViews.ViewName) 
			FROM ASRSysViews 
			WHERE NOT Object_ID(ASRSysViews.ViewName) IS null
			UNION
			SELECT Object_ID(ASRSysTables.TableName) 
			FROM ASRSysTables 
			WHERE NOT Object_ID(ASRSysTables.TableName) IS null
			UNION
			SELECT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255))
			FROM ASRSysChildViews2
			INNER JOIN ASRSysTables 
				ON ASRSysChildViews2.tableID = ASRSysTables.tableID
			WHERE NOT OBJECT_ID(left('ASRSysCV' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ '#' + replace(ASRSysTables.tableName, ' ', '_')
				+ '#' + replace(@sGroupName, ' ', '_'), 255)) IS null;
		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int PRIMARY KEY CLUSTERED);
		INSERT INTO @SysProtects
			SELECT p.[ID] 
			FROM ASRSysProtectsCache p
						INNER JOIN SysColumns c ON (c.id = p.id
							AND c.[Name] = 'timestamp'
							AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
							OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0)))
			WHERE p.UID = @iUserGroupID
				AND p.[ProtectType] IN (204, 205)
				AND p.[Action] = 193			
				AND p.id IN (SELECT ID FROM @Phase1);
		-- Readable tables
		DECLARE @ReadableTables TABLE([Name] sysname PRIMARY KEY CLUSTERED);
		INSERT INTO @ReadableTables
			SELECT OBJECT_NAME(p.ID)
			FROM @SysProtects p;
		
		SET @sRealSource = '';
		IF @plngViewID > 0
		BEGIN
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @plngViewID;
		END
		ELSE
		BEGIN
			SELECT @sRealSource = tableName
			FROM ASRSysTables
			WHERE tableID = @plngTableID;
		END
		SET @fBaseTableReadable = 0
		IF len(@sRealSource) > 0
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM @ReadableTables
			WHERE name = @sRealSource;
		
			IF @iCount > 0
			BEGIN
				SET @fBaseTableReadable = 1;
			END
		END
	END
	DECLARE @Links TABLE([ID]						integer PRIMARY KEY CLUSTERED,
											 [utilityType]	integer,
											 [utilityID]		integer,
											 [screenID]			integer,
											 [LinkType]			integer,
											 [Element_Type]	integer,
											 [DrillDownHidden]				bit);
	INSERT INTO @Links ([ID],[utilityType],[utilityID],[screenID], [LinkType], [Element_Type], [DrillDownHidden])
	SELECT ASRSysSSIntranetLinks.ID,
					ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType,
					ASRSysSSIntranetLinks.Element_Type,
					0
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
			AND tableid = @plngTableID)
			AND (id NOT IN (SELECT linkid 
								FROM ASRSysSSIHiddenGroups
								WHERE groupName = @sGroupName));
	/* Remove any utility links from the temp table where the utility has been deleted or hidden from the current user.*/
	/* Or if the user does not permission to run them. */	
	DECLARE utilitiesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType,
					ASRSysSSIntranetLinks.Element_Type
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
				AND tableid = @plngTableID)
			AND (utilityID > 0 
				OR screenID > 0);
	OPEN utilitiesCursor;
	FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType, @iElement_Type;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iUtilID > 0
		BEGIN
			SET @fUtilOK = 1	;			
			/* Check if the utility is deleted or hidden from the user. */
			EXECUTE [dbo].[spASRIntCurrentAccessForRole]
								@sGroupName,
								@iUtilType,
								@iUtilID,
								@sAccess	OUTPUT;
			IF @sAccess = 'HD' 
			BEGIN
				/* Report/utility is hidden from the user. */
				--HERE FOR CHARTs **************************************************************************************************************************************
				IF @iElement_Type = 2
				BEGIN
					SET @fUtilOK = 1;				
					SET @fDrillDownHidden = 1;
				END
				ELSE
				BEGIN
					SET @fUtilOK = 0;
				END
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has system permission to run this type of report/utility. */
				IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA'
				BEGIN
					SELECT @pfPermitted = ASRSysGroupPermissions.permitted
					FROM ASRSysPermissionItems
					INNER JOIN ASRSysPermissionCategories 
					ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 					
							CASE 
								WHEN @iUtilType = 17 THEN 'CALENDARREPORTS'
								WHEN @iUtilType = 9 THEN 'MAILMERGE'
								WHEN @iUtilType = 2 THEN 'CUSTOMREPORTS'
								WHEN @iUtilType = 25 THEN 'WORKFLOW'
								WHEN @iUtilType = 35 THEN 'NINEBOXGRID'
								ELSE ''
							END
					LEFT OUTER JOIN ASRSysGroupPermissions 
					ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
						AND ASRSysGroupPermissions.groupName = @sGroupName
					WHERE ASRSysPermissionItems.itemKey = 'RUN';
					IF (@pfPermitted IS null) OR (@pfPermitted = 0)
					BEGIN
						/* User does not have system permission to run this type of report/utility. */
						--HERE FOR CHARTS**************************************************************************************************************************************
						IF @iElement_Type = 2
						BEGIN
							SET @fUtilOK = 1;
							SET @fDrillDownHidden = 1;
						END
						ELSE
						BEGIN
							SET @fUtilOK = 0;
						END
					END
				END
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has read permission on the report/utility base table or any views on it. */
				SET @iBaseTableID = 0;
				IF @iUtilType = 17 /* Calendar Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCalendarReports
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 2 /* Custom Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCustomReportsName
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 9 /* Mail Merge */
				BEGIN
					SELECT @iBaseTableID = TableID
					FROM ASRSysMailMergeName
					WHERE MailMergeID = @iUtilID;
				END
				IF @iUtilType = 35 /* 9-Box Grid Reports */
				BEGIN				
					SELECT @iBaseTableID = TableID
					FROM ASRSysCrossTab
					WHERE CrossTabID = @iUtilID
					AND CrossTabType = 4;
				END
				/* Not check required for reports/utilities without a base table.
				OR reports/utilities based on the top-level table if the user has read permission on the current view. */
				IF (@iBaseTableID > 0)
					AND((@fBaseTableReadable = 0)
						OR (@iBaseTableID <> @plngTableID))
				BEGIN
					IF (@iLinkType <> 0) -- Hypertext link
						AND (@fBaseTableReadable = 0)
						AND (@iBaseTableID = @plngTableID)
					BEGIN
						/* The report/utility is based on the top-level table, and the user does NOT have read permission
						on the current view (on which Button & DropdownList links are scoped). */
						SET @fUtilOK = 0;
					END
					ELSE
					BEGIN
						SELECT @iCount = COUNT(p.ID)
						FROM @SysProtects p
						WHERE OBJECT_NAME(p.ID) IN (SELECT ASRSysTables.tableName
							FROM ASRSysTables
							WHERE ASRSysTables.tableID = @iBaseTableID
						UNION 
							SELECT ASRSysViews.viewName
								FROM ASRSysViews
								WHERE ASRSysViews.viewTableID = @iBaseTableID
						UNION
							SELECT
								left('ASRSysCV' 
									+ convert(varchar(1000), ASRSysChildViews2.childViewID) 
									+ '#'
									+ replace(ASRSysTables.tableName, ' ', '_')
									+ '#'
									+ replace(@sGroupName, ' ', '_'), 255)
								FROM ASRSysChildViews2
								INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID
								WHERE ASRSysChildViews2.role = @sGroupName
									AND ASRSysChildViews2.tableID = @iBaseTableID);
						IF @iCount = 0 
						BEGIN
							SET @fUtilOK = 0;
						END
					END
				END
			END
			/* For some reason the user cannot use this report/utility, so remove it from the temp table of links. */
			IF @fUtilOK = 0 
			BEGIN
				DELETE FROM @Links
				WHERE utilityType = @iUtilType
					AND utilityID = @iUtilID;
			END
			IF @fDrillDownHidden = 1
			BEGIN
				UPDATE @Links
				SET DrillDownHidden = 1 
				WHERE utilityType = @iUtilType
					AND utilityID = @iUtilID;
			END
			
		END
		
		IF (@iScreenID > 0) AND (UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> 'SA')
		BEGIN
			/* Do not display the link if the user does not have permission to read the defined view/table for the screen. */
			SELECT @iTableID = ASRSysTables.tableID, 
				@sTableName = ASRSysTables.tableName,
				@iTableType = ASRSysTables.tableType
			FROM ASRSysScreens
						INNER JOIN ASRSysTables 
						ON ASRSysScreens.tableID = ASRSysTables.tableID
			WHERE screenID = @iScreenID;
			SET @sRealSource = '';
			IF @iTableType  = 2
			BEGIN
				SET @iChildViewID = 0;
				/* Child table - check child views. */
				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTableID
					AND [role] = @sGroupName;
				
				IF @iChildViewID IS null SET @iChildViewID = 0;
				
				IF (@iChildViewID > 0) AND (@fBaseTableReadable = 1)
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTableName, ' ', '_') +
						'#' + replace(@sGroupName, ' ', '_');
				
					SET @sRealSource = left(@sRealSource, 255);
				END
				ELSE
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
			ELSE
			BEGIN
				/* Not a child table - must be the top-level table. Check if the user has 'read' permission on the defined view. */
				SET @sRealSource = '';
				IF @plngViewID > 0
				BEGIN
					SELECT @sRealSource = viewName
					FROM ASRSysViews
					WHERE viewID = @plngViewID;
				END
				ELSE
				BEGIN
					SELECT @sRealSource = tableName
					FROM ASRSysTables
					WHERE tableID = @plngTableID;
				END
			END
			IF len(@sRealSource) > 0
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM @ReadableTables
				WHERE name = @sRealSource;
			
				IF @iCount = 0
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
		END
		FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType, @iElement_Type;
	END
	CLOSE utilitiesCursor;
	DEALLOCATE utilitiesCursor;
	/* Remove the Workflow links if the URL has not been configured. */
	SELECT @sURL = isnull(settingValue , '')
	FROM ASRSysSystemSettings
	WHERE section = 'MODULE_WORKFLOW'		
		AND settingKey = 'Param_URL';	
	
	
	IF LEN(@sURL) = 0
	BEGIN
		DELETE FROM @Links
		WHERE utilityType = 25;
	END
	SELECT ASRSysSSIntranetLinks.*, 
		CASE 
			WHEN ASRSysSSIntranetLinks.utilityType = 9 THEN ASRSysMailMergeName.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 2 THEN ASRSysCustomReportsName.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 17 THEN ASRSysCalendarReports.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 35 THEN ASRSysCrossTab.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 25 THEN 0
			ELSE null
		END AS [baseTable],
		ASRSysColumns.ColumnName as [Chart_ColumnName],
		tvL.DrillDownHidden as [DrillDownHidden]
	FROM ASRSysSSIntranetLinks
			LEFT OUTER JOIN ASRSysMailMergeName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysMailMergeName.MailMergeID
				AND ASRSysSSIntranetLinks.utilityType = 9
			LEFT OUTER JOIN ASRSysCalendarReports 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCalendarReports.ID
				AND ASRSysSSIntranetLinks.utilityType = 17
			LEFT OUTER JOIN ASRSysCrossTab 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCrossTab.CrossTabID
				AND ASRSysSSIntranetLinks.utilityType = 35
			LEFT OUTER JOIN ASRSysCustomReportsName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCustomReportsName.ID
				AND ASRSysSSIntranetLinks.utilityType = 2
			LEFT OUTER JOIN ASRSysColumns
			ON ASRSysSSIntranetLinks.Chart_ColumnID = ASRSysColumns.columnId		
			LEFT OUTER JOIN @Links tvL
			ON ASRSysSSIntranetLinks.ID = tvL.ID
	WHERE ASRSysSSIntranetLinks.ID IN (SELECT ID FROM @Links)
	ORDER BY ASRSysSSIntranetLinks.linkOrder;
	
END

GO



IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntDefUsage]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntDefUsage];
GO
CREATE PROCEDURE [dbo].[sp_ASRIntDefUsage] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sExecSQL		nvarchar(MAX),
		@sJobTypeName		varchar(255),
		@sCurrentUser		sysname,
		@sDescription		varchar(MAX),
		@sName				varchar(255), 
		@sUserName			varchar(255), 
		@sAccess			varchar(MAX),
		@fIsBatch			bit,
		@sUtilType			varchar(255),
		@iCompID			integer,
		@iRootExprID		integer,
		@sRoleName			varchar(255),
		@fSysSecMgr			bit,
		@iCount				integer,
		@sActualUserName	sysname,
		@iUserGroupID		integer;
		
	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
	
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	SET @sExecSQL = '';
	SET @sCurrentUser = SYSTEM_USER;

	DECLARE @results TABLE([description] varchar(MAX));
	DECLARE @rootExprs TABLE(exprID integer);

	IF @intType = 11 OR @intType = 12
	BEGIN
		/* Create a table of IDs of the expressions that use the given filter or calc. */
		DECLARE expr_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT componentID 
			FROM ASRSysExprComponents
			WHERE calculationID = @intID
				OR filterID = @intID
				OR (fieldSelectionFilter = @intID AND type = 1)
		OPEN expr_cursor
		FETCH NEXT FROM expr_cursor INTO @iCompID
		WHILE (@@fetch_status = 0)
		BEGIN
			execute sp_ASRIntGetRootExpressionIDs @iCompID, @iRootExprID OUTPUT
			IF @iRootExprID > 0
			BEGIN
				INSERT INTO @rootExprs (exprID) VALUES (@iRootExprID)
			END
			FETCH NEXT FROM expr_cursor INTO @iCompID
		END
		CLOSE expr_cursor
		DEALLOCATE expr_cursor
	END

	IF @intType = 1 OR @intType = 2 OR @intType = 9 OR @intType = 17 OR @intType = 35
	BEGIN
		/* Reports & Utilities
		Check for usage in Batch Jobs */
		IF @intType = 1 SET @sJobTypeName = 'CROSS TAB'
		IF @intType = 2 SET @sJobTypeName = 'CUSTOM REPORT'
		IF @intType = 9 SET @sJobTypeName = 'MAIL MERGE' 
		IF @intType = 17 SET @sJobTypeName = 'CALENDAR REPORT'
		IF @intType = 35 SET @sJobTypeName = '9-BOX GRID REPORT'
		
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT ASRSysBatchJobName.Name, 
				ASRSysBatchJobName.UserName, 
				ASRSysBatchJobAccess.Access,
				AsrSysBatchJobName.IsBatch
			FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID
			INNER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.groupname = @sRoleName
			WHERE ASRSysBatchJobDetails.JobType = @sJobTypeName
				AND ASRSysBatchJobDetails.JobID = @intID
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @fIsBatch = 1 BEGIN
				SET @sDescription = 'Batch Job: '
			END ELSE BEGIN
				SET @sDescription = 'Report Pack: '
			END

			IF (@sUserName <> @sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor

		SELECT @iCount = COUNT(*) 
		FROM [ASRSysSSIntranetLinks]
		WHERE [ASRSysSSIntranetLinks].[utilityID] = @intID
			AND [ASRSysSSIntranetLinks].[utilityType] = @intType
		IF @iCount > 0
		BEGIN
		   	INSERT INTO @results (description) VALUES ('Self-service intranet link')
		END
	END

	IF @intType = 10
	BEGIN
		/* Picklists 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports, Calendar Reports and Mail Merges*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Cross Tab', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName					
				WHERE PickListID =@intID
					AND ASRSysCrossTab.CrossTabType <> 4
			UNION
				SELECT DISTINCT '9-Box Grid Report', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName					
				WHERE PickListID =@intID
					AND ASRSysCrossTab.CrossTabType = 4
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.pickListID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.pickList = @intID OR ASRSysExportName.Parent1Picklist = @intID OR ASRSysExportName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.pickListID = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access 
				FROM ASRSysCustomReportsName
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.PickList = @intID 
					OR ASRSysCustomReportsName.Parent1Picklist = @intID 
					OR ASRSysCustomReportsName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END, 
					ASRSysMailMergeName.Name, 
					ASRSysMailMergeName.UserName, 
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName 
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE PickListID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning'
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report'
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Picklist = @intID
					OR ASRSysMatchReportName.Table2Picklist = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSysCalendarReportAccess.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.PickList = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				WHERE ASRSysRecordProfileName.pickListID = @intID
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 11
	BEGIN
		/* Filters 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Cross Tab',
					ASRSysCrossTab.Name,
					ASRSysCrossTab.UserName,
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE ASRSysCrossTab.FilterID = @intID
					AND ASRSysCrossTab.CrossTabType <> 4
			UNION
				SELECT DISTINCT '9-Box Grid Report', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName					
				WHERE ASRSysCrossTab.FilterID = @intID
					AND ASRSysCrossTab.CrossTabType = 4
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsName
				LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.Filter = @intID
					OR ASRSysCustomReportsName.Parent1Filter = @intID
					OR ASRSysCustomReportsName.Parent2Filter = @intID
					OR ASRSYSCustomReportsChildDetails.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.FilterID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.Filter = @intID 
					OR ASRSysExportName.Parent1Filter = @intID
					OR ASRSysExportName.Parent2Filter = @intID
					OR ASRSysExportName.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.FilterID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeName.FilterID = @intID
			UNION
				SELECT DISTINCT
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning' 
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report' 
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Filter = @intID
					OR ASRSysMatchReportName.Table2Filter = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.filter = @intID
					OR ASRSysCalendarReportEvents.filterID = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID
				WHERE ASRSysRecordProfileName.FilterID = @intID
					OR ASRSYSRecordProfileTables.FilterID = @intID		

		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 12
	BEGIN
		/* Calculation.
		Check for usage in Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports 
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.DescriptionExpr =@intID 
					OR ASRSysCalendarReports.StartDateExpr = @intID 
					OR ASRSysCalendarReports.EndDateExpr = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name,
					ASRSysCustomReportsName.UserName,
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsDetails
				INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.CustomReportID = ASRSysCustomReportsName.ID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' 
					AND ASRSysCustomReportsDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportDetails
				INNER JOIN ASRSysExportName ON ASRSysExportDetails.ID = ASRSysExportName.ID 
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysExportDetails.type) = 'X' 
					AND ASRSysExportDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalItems
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID 
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalItems.ValueType = 4 
					AND ASRSysGlobalItems.ExprID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeColumns.ColumnID = @intID
					AND upper(ASRSysMailMergeColumns.type) = 'E'
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + '''';
			END
    
    		INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess;
		END
		
		CLOSE usage_cursor;
		DEALLOCATE usage_cursor;
	END

	/* Return the usage records. */
	SELECT * FROM @results ORDER BY description;

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

	IF @piUtilityType = 35 /* 9-Box Grid Report*/
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

	SET NOCOUNT ON;

	SELECT TableID, TableName, TableType, DefaultOrderID, RecordDescExprID FROM dbo.ASRSysTables;

	SELECT ColumnID, TableID, ColumnName, DataType, ColumnType, Use1000Separator, Size, Decimals, LookupTableID, LookupColumnID FROM dbo.ASRSysColumns;

	SELECT ParentID, ChildID FROM dbo.ASRSysRelations;

	SELECT ModuleKey, ParameterKey, ISNULL(ParameterValue,'') AS ParameterValue, ParameterType FROM dbo.ASRSysModuleSetup;

	SELECT * FROM dbo.ASRSysUserSettings WHERE Username = @Username;

	SELECT functionID, functionName, returnType FROM dbo.ASRSysFunctions;

	SELECT * FROM dbo.ASRSysFunctionParameters ORDER BY functionID, parameterIndex;

	SELECT * FROM dbo.ASRSysOperators;

	SELECT * FROM dbo.ASRSysOperatorParameters ORDER BY OperatorID, parameterIndex;
	
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

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetPersonnelParameters]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetPersonnelParameters];
GO
CREATE PROCEDURE [dbo].[sp_ASRIntGetPersonnelParameters] (
	@piEmployeeTableID	integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	SET @piEmployeeTableID = 0;

	-- Get the EMPLOYEE table information.
	SELECT @piEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel';
	IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0;

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetTrainingBookingParameters]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetTrainingBookingParameters];
GO
CREATE PROCEDURE [dbo].[sp_ASRIntGetTrainingBookingParameters] (
	@piEmployeeTableID			integer	OUTPUT,
	@piCourseTableID			integer	OUTPUT,
	@piCourseCancelDateColumnID	integer	OUTPUT,
	@piTBTableID				integer	OUTPUT,
	@pfTBTableSelect			bit		OUTPUT,
	@pfTBTableInsert			bit		OUTPUT,
	@pfTBTableUpdate			bit		OUTPUT,
	@piTBStatusColumnID			integer	OUTPUT,
	@pfTBStatusColumnUpdate		bit		OUTPUT,
	@piTBCancelDateColumnID		integer	OUTPUT,
	@pfTBCancelDateColumnUpdate	bit		OUTPUT,
	@pfTBProvisionalStatusExists	bit	OUTPUT,
	@piWaitListTableID			integer	OUTPUT,
	@pfWaitListTableInsert			bit	OUTPUT,
	@pfWaitListTableDelete			bit	OUTPUT,
	@piWaitListCourseTitleColumnID		integer	OUTPUT,
	@pfWaitListCourseTitleColumnUpdate	bit	OUTPUT,
	@pfWaitListCourseTitleColumnSelect	bit	OUTPUT,
	@piBulkBookingDefaultViewID		integer	OUTPUT,
	@piBulkBookingDefaultOrderID		integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @fOK			bit,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sTempName			sysname,
		@iTempAction		integer,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@sRealSource		sysname,
		@iStatusCount		integer,
		@iChildViewID		integer,
		@sTBTableName		sysname,
		@sWLTableName		sysname,
		@sActualUserName	sysname;
		
	/* Training Booking information. */
	SET @fOK = 1

	SET @piEmployeeTableID = 0

	SET @piCourseTableID = 0
	SET @piCourseCancelDateColumnID = 0

	SET @piTBTableID = 0
	SET @pfTBTableSelect = 0
	SET @pfTBTableInsert = 0
	SET @pfTBTableUpdate = 0
	SET @piTBStatusColumnID = 0
	SET @pfTBStatusColumnUpdate = 0
	SET @piTBCancelDateColumnID = 0
	SET @pfTBCancelDateColumnUpdate = 0
	SET @pfTBProvisionalStatusExists = 0

	SET @piWaitListTableID = 0
	SET @pfWaitListTableInsert = 0
	SET @pfWaitListTableDelete = 0
	SET @piWaitListCourseTitleColumnID = 0
	SET @pfWaitListCourseTitleColumnUpdate = 0
	SET @pfWaitListCourseTitleColumnSelect = 0

	SET @piBulkBookingDefaultViewID = 0
	
	/* Get the current user's group id. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'
	OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

		/* Get the EMPLOYEE table information. */
		SELECT @piEmployeeTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_EmployeeTable'
		IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0

		/* Get the COURSE table information. */
		SELECT @piCourseTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseTable'
		IF @piCourseTableID IS NULL SET @piCourseTableID = 0

		IF @piCourseTableID > 0
		BEGIN
			SELECT @piCourseCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_CourseCancelDate'
			IF @piCourseCancelDateColumnID IS NULL SET @piCourseCancelDateColumnID = 0
		END

		/* Get the TRAINING BOOKING table information. */
		SELECT @piTBTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_TrainBookTable'
		IF @piTBTableID IS NULL SET @piTBTableID = 0


		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID]				int,
								   [columns]		varbinary(8000),
								   [Action]			tinyint,
								   [ProtectType]	tinyint)
		INSERT INTO @SysProtects
		SELECT [ID], [Columns], [Action], [ProtectType] FROM ASRSysProtectsCache
			WHERE [UID] = @iUserGroupID AND [Action] IN (193, 195, 196, 197)

		IF @piTBTableID > 0
		BEGIN
			SELECT @sTBTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piTBTableID

			SELECT @piTBStatusColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookStatus'
			IF @piTBStatusColumnID IS NULL SET @piTBStatusColumnID = 0

			SELECT @piTBCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookCancelDate'
			IF @piTBCancelDateColumnID IS NULL SET @piTBCancelDateColumnID = 0

			SET @sCommand = 'SELECT @iStatusCount = COUNT(*)' +
				' FROM ASRSysColumnControlValues' +
				' WHERE columnID = ' + convert(nvarchar(100), @piTBStatusColumnID) +
				' AND value = ''P'''
			SET @sParamDefinition = N'@iStatusCount integer OUTPUT'
			EXEC sp_executesql @sCommand, @sParamDefinition, @iStatusCount OUTPUT
			IF @iStatusCount > 0 SET @pfTBProvisionalStatusExists = 1

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				/* System/Security managers must have all permissions granted. */
				SET @pfTBTableSelect = 1
				SET @pfTBTableInsert = 1
				SET @pfTBTableUpdate = 1
				SET @pfTBStatusColumnUpdate = 1
				SET @pfTBCancelDateColumnUpdate = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piTBTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN

					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTBTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (193, 195, 197)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN

						IF @iTempAction = 193
						BEGIN
							SET @pfTBTableSelect = 1
						END
						IF @iTempAction = 195
						BEGIN
							SET @pfTBTableInsert = 1
						END
						IF @iTempAction = 197
						BEGIN
							SET @pfTBTableUpdate = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @tbColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @tbColumnPermissions
					SELECT 
						ASRSysColumns.columnId,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piTBTableID
						AND (ASRSysColumns.columnId = @piTBStatusColumnID
							OR ASRSysColumns.columnId = @piTBCancelDateColumnID))
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfTBStatusColumnUpdate = granted

					FROM @tbColumnPermissions
					WHERE columnID = @piTBStatusColumnID
						AND action = 197
					IF @pfTBStatusColumnUpdate IS NULL SET @pfTBStatusColumnUpdate = 0

					SELECT @pfTBCancelDateColumnUpdate = granted
					FROM @tbColumnPermissions
					WHERE columnID = @piTBCancelDateColumnID
						AND action = 197
					IF @pfTBCancelDateColumnUpdate IS NULL SET @pfTBCancelDateColumnUpdate = 0

				END
			END
		END

		/* Get the waiting list table information. */
		SELECT @piWaitListTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_WaitListTable'
		IF @piWaitListTableID IS NULL SET @piWaitListTableID = 0

		IF @piWaitListTableID > 0
		BEGIN
			SELECT @sWLTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piWaitListTableID

			SELECT @piWaitListCourseTitleColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_WaitListCourseTitle'
			IF @piWaitListCourseTitleColumnID IS NULL SET @piWaitListCourseTitleColumnID = 0

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				SET @pfWaitListTableInsert = 1
				SET @pfWaitListTableDelete = 1
				SET @pfWaitListCourseTitleColumnUpdate = 1
				SET @pfWaitListCourseTitleColumnSelect = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piWaitListTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sWLTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (195, 196)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iTempAction = 195
						BEGIN
							SET @pfWaitListTableInsert = 1
						END
						IF @iTempAction = 196
						BEGIN
							SET @pfWaitListTableDelete = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction

					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @waitListColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @waitListColumnPermissions
					SELECT 
						ASRSysColumns.columnId,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piWaitListTableID
						AND ASRSysColumns.columnId = @piWaitListCourseTitleColumnID)
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfWaitListCourseTitleColumnUpdate = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 197
					IF @pfWaitListCourseTitleColumnUpdate IS NULL SET @pfWaitListCourseTitleColumnUpdate = 0

					SELECT @pfWaitListCourseTitleColumnSelect = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 193
					IF @pfWaitListCourseTitleColumnSelect IS NULL SET @pfWaitListCourseTitleColumnSelect = 0

				END
			END
		END

		/* Get the Bulk Booking default view. */
		SELECT @piBulkBookingDefaultViewID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_BulkBookingDefaultView'
		IF @piBulkBookingDefaultViewID IS NULL SET @piBulkBookingDefaultViewID = 0

			/* Get the Bulk Booking default order. */
		SELECT @piBulkBookingDefaultOrderID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_EmployeeOrder'
		IF @piBulkBookingDefaultOrderID IS NULL SET @piBulkBookingDefaultOrderID = 0
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntDefUsage]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntDefUsage];
GO

CREATE PROCEDURE [dbo].[sp_ASRIntDefUsage] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sExecSQL		nvarchar(MAX),
		@sJobTypeName		varchar(255),
		@sCurrentUser		sysname,
		@sDescription		varchar(MAX),
		@sName				varchar(255), 
		@sUserName			varchar(255), 
		@sAccess			varchar(MAX),
		@fIsBatch			bit,
		@sUtilType			varchar(255),
		@iCompID			integer,
		@iRootExprID		integer,
		@sRoleName			varchar(255),
		@fSysSecMgr			bit,
		@iCount				integer,
		@sActualUserName	sysname,
		@iUserGroupID		integer;
		
	EXEC [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;
	
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	SET @sExecSQL = '';
	SET @sCurrentUser = SYSTEM_USER;

	DECLARE @results TABLE([description] varchar(MAX));
	DECLARE @rootExprs TABLE(exprID integer);

	IF @intType = 11 OR @intType = 12
	BEGIN
		/* Create a table of IDs of the expressions that use the given filter or calc. */
		DECLARE expr_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT componentID 
			FROM ASRSysExprComponents
			WHERE calculationID = @intID
				OR filterID = @intID
				OR (fieldSelectionFilter = @intID AND type = 1)
		OPEN expr_cursor
		FETCH NEXT FROM expr_cursor INTO @iCompID
		WHILE (@@fetch_status = 0)
		BEGIN
			execute sp_ASRIntGetRootExpressionIDs @iCompID, @iRootExprID OUTPUT
			IF @iRootExprID > 0
			BEGIN
				INSERT INTO @rootExprs (exprID) VALUES (@iRootExprID)
			END
			FETCH NEXT FROM expr_cursor INTO @iCompID
		END
		CLOSE expr_cursor
		DEALLOCATE expr_cursor
	END

	IF @intType = 1 OR @intType = 2 OR @intType = 9 OR @intType = 17
	BEGIN
		/* Reports & Utilities
		Check for usage in Batch Jobs */
		IF @intType = 1 SET @sJobTypeName = 'CROSS TAB'
		IF @intType = 2 SET @sJobTypeName = 'CUSTOM REPORT'
		IF @intType = 9 SET @sJobTypeName = 'MAIL MERGE' 
		IF @intType = 17 SET @sJobTypeName = 'CALENDAR REPORT'
		
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT ASRSysBatchJobName.Name, 
				ASRSysBatchJobName.UserName, 
				ASRSysBatchJobAccess.Access,
				AsrSysBatchJobName.IsBatch
			FROM ASRSysBatchJobDetails
			INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID
			INNER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobName.ID = ASRSysBatchJobAccess.ID
				AND ASRSysBatchJobAccess.groupname = @sRoleName
			WHERE ASRSysBatchJobDetails.JobType = @sJobTypeName
				AND ASRSysBatchJobDetails.JobID = @intID
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @fIsBatch = 1 BEGIN
				SET @sDescription = 'Batch Job: '
			END ELSE BEGIN
				SET @sDescription = 'Report Pack: '
			END

			IF (@sUserName <> @sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sName, @sUserName, @sAccess, @fIsBatch
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor

		SELECT @iCount = COUNT(*) 
		FROM [ASRSysSSIntranetLinks]
		WHERE [ASRSysSSIntranetLinks].[utilityID] = @intID
			AND [ASRSysSSIntranetLinks].[utilityType] = @intType
		IF @iCount > 0
		BEGIN
		   	INSERT INTO @results (description) VALUES ('Self-service intranet link')
		END
	END

	IF @intType = 10
	BEGIN
		/* Picklists 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports, Calendar Reports and Mail Merges*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Cross Tab', 
					ASRSysCrossTab.Name, 
					ASRSysCrossTab.UserName, 
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE PickListID =@intID
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.pickListID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.pickList = @intID OR ASRSysExportName.Parent1Picklist = @intID OR ASRSysExportName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.pickListID = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access 
				FROM ASRSysCustomReportsName
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.PickList = @intID 
					OR ASRSysCustomReportsName.Parent1Picklist = @intID 
					OR ASRSysCustomReportsName.Parent2Picklist = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END, 
					ASRSysMailMergeName.Name, 
					ASRSysMailMergeName.UserName, 
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName 
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE PickListID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning'
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report'
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Picklist = @intID
					OR ASRSysMatchReportName.Table2Picklist = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSysCalendarReportAccess.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.PickList = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				WHERE ASRSysRecordProfileName.pickListID = @intID
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 11
	BEGIN
		/* Filters 
		Check for usage in Cross Tabs, Data Transfers, Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Cross Tab',
					ASRSysCrossTab.Name,
					ASRSysCrossTab.UserName,
					ASRSysCrossTabAccess.Access
				FROM ASRSysCrossTab
				INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID
					AND ASRSysCrossTabAccess.groupname = @sRoleName
				WHERE ASRSysCrossTab.FilterID = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name, 
					ASRSysCustomReportsName.UserName, 
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsName
				LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE ASRSysCustomReportsName.Filter = @intID
					OR ASRSysCustomReportsName.Parent1Filter = @intID
					OR ASRSysCustomReportsName.Parent2Filter = @intID
					OR ASRSYSCustomReportsChildDetails.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Data Transfer',
					ASRSysDataTransferName.Name,
					ASRSysDataTransferName.UserName,
					ASRSysDataTransferAccess.Access
				FROM ASRSysDataTransferName
				INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferAccess.ID
					AND ASRSysDataTransferAccess.groupname = @sRoleName
				WHERE ASRSysDataTransferName.FilterID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportName
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE ASRSysExportName.Filter = @intID 
					OR ASRSysExportName.Parent1Filter = @intID
					OR ASRSysExportName.Parent2Filter = @intID
					OR ASRSysExportName.ChildFilter = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalFunctions
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalFunctions.FilterID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeName.FilterID = @intID
			UNION
				SELECT DISTINCT
					CASE 
						WHEN ASRSysMatchReportName.matchReportType = 0 THEN 'Match Report'
						WHEN ASRSysMatchReportName.matchReportType = 1 THEN 'Succession Planning' 
						WHEN ASRSysMatchReportName.matchReportType = 2 THEN 'Career Progression' 
						ELSE 'Match Report' 
					END,
					ASRSysMatchReportName.Name,
					ASRSysMatchReportName.UserName,
					ASRSysMatchReportAccess.Access
				FROM ASRSysMatchReportName
				INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID
					AND ASRSysMatchReportAccess.groupname = @sRoleName
				WHERE ASRSysMatchReportName.Table1Filter = @intID
					OR ASRSysMatchReportName.Table2Filter = @intID
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports
				LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.filter = @intID
					OR ASRSysCalendarReportEvents.filterID = @intID
			UNION
				SELECT DISTINCT 'Record Profile',
					ASRSysRecordProfileName.Name,
					ASRSysRecordProfileName.UserName,
					ASRSysRecordProfileAccess.Access
				FROM ASRSysRecordProfileName
				INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID
					AND ASRSysRecordProfileAccess.groupname = @sRoleName
				LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID
				WHERE ASRSysRecordProfileName.FilterID = @intID
					OR ASRSYSRecordProfileTables.FilterID = @intID		

		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + ''''
			END
    
    			INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		END
		CLOSE usage_cursor
		DEALLOCATE usage_cursor
	END

	IF @intType = 12
	BEGIN
		/* Calculation.
		Check for usage in Globals, Exports, Custom Reports and Mail Merges. Also other Filters and Calculations*/
		DECLARE usage_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT 'Calculation', Name, UserName, Access 
				FROM ASRSysExpressions
				WHERE Type = 10
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 'Calendar Report', 
					ASRSysCalendarReports.Name, 
					ASRSysCalendarReports.UserName, 
					ASRSysCalendarReportAccess.Access
				FROM ASRSysCalendarReports 
				INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID
					AND ASRSysCalendarReportAccess.groupname = @sRoleName
				WHERE ASRSysCalendarReports.DescriptionExpr =@intID 
					OR ASRSysCalendarReports.StartDateExpr = @intID 
					OR ASRSysCalendarReports.EndDateExpr = @intID
			UNION
				SELECT DISTINCT 'Custom Report', 
					ASRSysCustomReportsName.Name,
					ASRSysCustomReportsName.UserName,
					ASRSysCustomReportAccess.Access
				FROM ASRSysCustomReportsDetails
				INNER JOIN ASRSysCustomReportsName ON ASRSysCustomReportsDetails.CustomReportID = ASRSysCustomReportsName.ID
				INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID
					AND ASRSysCustomReportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' 
					AND ASRSysCustomReportsDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Export',
					ASRSysExportName.Name,
					ASRSysExportName.UserName,
					ASRSysExportAccess.Access
				FROM ASRSysExportDetails
				INNER JOIN ASRSysExportName ON ASRSysExportDetails.ID = ASRSysExportName.ID 
				INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID
					AND ASRSysExportAccess.groupname = @sRoleName
				WHERE UPPER(ASRSysExportDetails.type) = 'X' 
					AND ASRSysExportDetails.colExprID = @intID
			UNION
				SELECT DISTINCT 'Filter', Name, UserName, Access
				FROM ASRSysExpressions
				WHERE Type = 11
					AND ExprID IN (SELECT exprID FROM @rootExprs)
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysGlobalFunctions.Type = 'A' THEN 'Global Add'
						WHEN ASRSysGlobalFunctions.Type = 'D' THEN 'Global Delete'
						WHEN ASRSysGlobalFunctions.Type = 'U' THEN 'Global Update'
						ELSE 'Global Function' 
					END, 
					ASRSysGlobalFunctions.Name,
					ASRSysGlobalFunctions.UserName,
					ASRSysGlobalAccess.Access
				FROM ASRSysGlobalItems
				INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID 
				INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID
					AND ASRSysGlobalAccess.groupname = @sRoleName
				WHERE ASRSysGlobalItems.ValueType = 4 
					AND ASRSysGlobalItems.ExprID = @intID
			UNION
				SELECT DISTINCT 
					CASE 
						WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'
						ELSE 'Mail Merge' 
					END,
					ASRSysMailMergeName.Name,
					ASRSysMailMergeName.UserName,
					ASRSysMailMergeAccess.Access
				FROM ASRSysMailMergeName
				INNER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID
				INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID
					AND ASRSysMailMergeAccess.groupname = @sRoleName
				WHERE ASRSysMailMergeColumns.ColumnID = @intID
					AND upper(ASRSysMailMergeColumns.type) = 'E'
			
		OPEN usage_cursor
		FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sDescription = @sUtilType + ': '

			IF (@sUserName <>@sCurrentUser)
				AND (@sAccess = 'HD') AND (@fSysSecMgr = 0)
			BEGIN
				SET @sDescription = @sDescription + '<Hidden by ' + @sUserName + '>'
			END
			ELSE
			BEGIN
				SET @sDescription = @sDescription + '''' + @sName + '''';
			END
    
    		INSERT INTO @results (description) VALUES (@sDescription)

			FETCH NEXT FROM usage_cursor INTO @sUtilType, @sName, @sUserName, @sAccess;
		END
		
		CLOSE usage_cursor;
		DEALLOCATE usage_cursor;
	END

	/* Return the usage records. */
	SELECT * FROM @results ORDER BY description;

END

GO

CREATE PROCEDURE [dbo].[spASRIntGetEventLogRecords] (
	@pfError 						bit 				OUTPUT, 
	@psFilterUser					varchar(MAX),
	@piFilterType					integer,
	@piFilterStatus					integer,
	@piFilterMode					integer,
	@psOrderColumn					varchar(MAX),
	@psOrderOrder					varchar(MAX),
	@piRecordsRequired				integer,
	@pfFirstPage					bit					OUTPUT,
	@pfLastPage						bit					OUTPUT,
	@psAction						varchar(100),
	@piTotalRecCount				integer				OUTPUT,
	@piFirstRecPos					integer				OUTPUT,
	@piCurrentRecCount				integer
)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE	@sRealSource 			sysname,
			@sSelectSQL				varchar(MAX),
			@iTempCount 			integer,
			@sExecString			nvarchar(MAX),
			@sTempExecString		nvarchar(MAX),
			@sTempParamDefinition	nvarchar(500),
			@iCount					integer,
			@iGetCount				integer,
			@sFilterSQL				varchar(MAX),
			@sOrderSQL				varchar(MAX),
			@sReverseOrderSQL		varchar(MAX);
			
	/* Clean the input string parameters. */
	IF len(@psAction) > 0 SET @psAction = replace(@psAction, '''', '''''');
	IF len(@psFilterUser) > 0 SET @psFilterUser = replace(@psFilterUser, '''', '''''');
	IF len(@psOrderColumn) > 0 SET @psOrderColumn = replace(@psOrderColumn, '''', '''''');
	IF len(@psOrderOrder) > 0 SET @psOrderOrder = replace(@psOrderOrder, '''', '''''');

	/* Initialise variables. */
	SET @pfError = 0;
	SET @sExecString = '';
	SET @sRealSource = 'ASRSysEventLog';
	SET @psAction = UPPER(@psAction);

	IF (@psAction <> 'MOVEPREVIOUS') AND (@psAction <> 'MOVENEXT') AND (@psAction <> 'MOVELAST') 
		BEGIN
			SET @psAction = 'MOVEFIRST';
		END

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 50;

	/* Construct the filter SQL from ther input parameters. */
	SET @sFilterSQL = '';
	
	SET @sFilterSQL = @sFilterSQL + ' Type NOT IN (23, 24) ';

	IF @psFilterUser <> '-1' 
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' LOWER(username) = ''' + lower(@psFilterUser) + '''';
	END
	IF @piFilterType <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' Type = ' + convert(varchar(MAX), @piFilterType) + ' ';
	END
	IF @piFilterStatus <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		SET @sFilterSQL = @sFilterSQL + ' Status = ' + convert(varchar(MAX), @piFilterStatus) + ' ';
	END
	IF @piFilterMode <> -1
	BEGIN
		IF len(@sFilterSQL) > 0 SET @sFilterSQL = @sFilterSQL + ' AND ';
		--SET @sFilterSQL = @sFilterSQL + ' Mode = ' + convert(varchar(MAX), @piFilterMode) + ' ';
		SET @sFilterSQL = @sFilterSQL + 
			CASE @piFilterMode 
				WHEN 1 THEN '[Mode] = 1 AND [ReportPack] = 0'
				WHEN 2 THEN '[ReportPack] = 1'
				WHEN 0 THEN '[Mode] = 0 AND [ReportPack] = 0'
			END 
	END
	
	/* Construct the order SQL from ther input parameters. */
	SET @sOrderSQL = '';
	IF @psOrderColumn = 'Type'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Type] 
					WHEN 1 THEN ''Cross Tab''
					WHEN 2 THEN ''Custom Report''
					WHEN 3 THEN ''Data Transfer''
					WHEN 4 THEN ''Export''
					WHEN 5 THEN ''Global Add''
					WHEN 6 THEN ''Global Delete''
					WHEN 7 THEN ''Global Update''
					WHEN 8 THEN ''Import''
					WHEN 9 THEN ''Mail Merge''
					WHEN 10 THEN ''Diary Delete''
					WHEN 11 THEN ''Diary Rebuild''
					WHEN 12 THEN ''Email Rebuild''
					WHEN 13 THEN ''Standard Report''
					WHEN 14 THEN ''Record Editing''
					WHEN 15 THEN ''System Error''
					WHEN 16 THEN ''Match Report''
					WHEN 17 THEN ''Calendar Report''
					WHEN 18 THEN ''Envelopes & Labels''
					WHEN 19 THEN ''Label Definition''
					WHEN 20 THEN ''Record Profile''
					WHEN 21 THEN ''Succession Planning''
					WHEN 22 THEN ''Career Progression''
					WHEN 25 THEN ''Workflow Rebuild''
					WHEN 35 THEN ''9-Box Grid Report''
					ELSE ''Unknown''
				END ';
	END
	ELSE
	BEGIN
		IF @psOrderColumn = 'Mode'
		BEGIN
			SET @sOrderSQL =	
				' CASE ' + @piFilterMode + '
						WHEN 1 THEN ''Batch''
						WHEN 0 THEN ''Manual''
						WHEN 2 THEN ''Pack''
					END ';
		END
		ELSE 
		BEGIN
			IF @psOrderColumn = 'Status'
			BEGIN
				SET @sOrderSQL =	
					' CASE [Status]
							WHEN 0 THEN ''Pending''
							WHEN 1 THEN ''Cancelled''
							WHEN 2 THEN ''Failed''
							WHEN 3 THEN ''Successful''
							WHEN 4 THEN ''Skipped''
							WHEN 5 THEN ''Error''
							ELSE ''Unknown''
						END ';
			END
			ELSE
			BEGIN
				SET @sOrderSQL = @psOrderColumn;
			END
		END
	END
	
	SET @sReverseOrderSQL = @sOrderSQL;
	if @psOrderOrder = 'DESC'
	BEGIN
		SET @sReverseOrderSQL = @sReverseOrderSQL + ' ASC ';
	END
	ELSE
	BEGIN
		SET @sReverseOrderSQL = @sReverseOrderSQL + ' DESC ';
	END

	SET @sOrderSQL = @sOrderSQL + ' ' + @psOrderOrder + ' ';


	SET @sSelectSQL = '[DateTime],
					[EndTime],
					IsNull([Duration],-1) AS ''Duration'', 
		 			CASE [Type] 
						WHEN 0 THEN ''Unknown''
						WHEN 1 THEN ''Cross Tab'' 
						WHEN 2 THEN ''Custom Report'' 
						WHEN 3 THEN ''Data Transfer'' 
						WHEN 4 THEN ''Export'' 
						WHEN 5 THEN ''Global Add'' 
						WHEN 6 THEN ''Global Delete'' 
						WHEN 7 THEN ''Global Update'' 
						WHEN 8 THEN ''Import'' 
						WHEN 9 THEN ''Mail Merge'' 
						WHEN 10 THEN ''Diary Delete'' 
						WHEN 11 THEN ''Diary Rebuild''
						WHEN 12 THEN ''Email Rebuild''
						WHEN 13 THEN ''Standard Report''
						WHEN 14 THEN ''Record Editing''
						WHEN 15 THEN ''System Error''
						WHEN 16 THEN ''Match Report''
						WHEN 17 THEN ''Calendar Report''
						WHEN 18 THEN ''Envelopes & Labels''
						WHEN 19 THEN ''Label Definition''
						WHEN 20 THEN ''Record Profile''
						WHEN 21	THEN ''Succession Planning''
						WHEN 22 THEN ''Career Progression''
						WHEN 25 THEN ''Workflow Rebuild''
						WHEN 35 THEN ''9-Box Grid Report''
						ELSE ''Unknown''  
					END + char(9) + 
				 	[Name] + char(9) + 
		 			CASE Status 
						WHEN 0 THEN ''Pending''
					  WHEN 1 THEN ''Cancelled'' 
						WHEN 2 THEN ''Failed'' 
						WHEN 3 THEN ''Successful'' 
						WHEN 4 THEN ''Skipped'' 
						WHEN 5 THEN ''Error''
						ELSE ''Unknown'' 
					END + char(9) +
					CASE 
						WHEN [Mode] = 1 AND ([ReportPack] = 0 OR [ReportPack] IS NULL) THEN ''Batch''
						WHEN [Mode] = 0 AND ([ReportPack] = 0 OR [ReportPack] IS NULL) THEN ''Manual''
						ELSE ''Pack''
				 	END + char(9) + 
					[Username] + char(9) + 
					IsNull(convert(varchar, [BatchJobID]), ''0'') + char(9) +
					IsNull(convert(varchar, [BatchRunID]), ''0'') + char(9) +
					IsNull([BatchName],'''') + char(9) +
					IsNull(convert(varchar, [SuccessCount]),''0'') + char(9) +
					IsNull(convert(varchar, [FailCount]), ''0'') AS EventInfo ';

		
	
	/****************************************************************************************************************************************/
	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.ID) FROM ' + @sRealSource;

	IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;

	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;
	SET @piTotalRecCount = @iCount;
	/****************************************************************************************************************************************/
	
	IF len(@sSelectSQL) > 0 
		BEGIN
			SET @sSelectSQL = @sRealSource + '.ID, ' + @sSelectSQL;
			SET @sExecString = 'SELECT ' ;

			IF @psAction = 'MOVEFIRST'
				BEGIN
					SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
					
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource ;

					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = 1;
					SET @pfFirstPage = 1;
					SET @pfLastPage = 
					CASE 
						WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
						ELSE 0
					END;
				END
		
			IF (@psAction = 'MOVELAST')
				BEGIN
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;
					
					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL;
						END

					SET @sExecString = @sExecString + ')'

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1;
					IF @piFirstRecPos < 1 SET @piFirstRecPos = 1;
					SET @pfFirstPage = 	CASE 
									WHEN @piFirstRecPos = 1 THEN 1
									ELSE 0
								END;
					SET @pfLastPage = 1;

				END

			IF (@psAction = 'MOVENEXT') 
				BEGIN
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
						BEGIN
							SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
						END
					ELSE
						BEGIN
							SET @iGetCount = @piRecordsRequired;
						END

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID ' + 
						' FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sRealSource + '.ID ' + 
						' FROM ' + @sRealSource;

					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
					
					SET @sExecString = @sExecString + ')';

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL;
						END

					SET @sExecString = @sExecString + ')';

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END

					/* Set the position variables */
					SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount;
					SET @pfFirstPage = 0
					SET @pfLastPage = 	CASE 
									WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
									ELSE 0
								END;
				END

			IF @psAction = 'MOVEPREVIOUS'
				BEGIN	
					SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sRealSource;

					IF @piFirstRecPos <= @piRecordsRequired
						BEGIN
							SET @iGetCount = @piFirstRecPos - 1;
						END
					ELSE
						BEGIN
							SET @iGetCount = @piRecordsRequired;
						END
		
					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;

					SET @sExecString = @sExecString + 
						' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID FROM ' + @sRealSource;
				
					/* Add the filter code. */
					IF len(@sFilterSQL) > 0
						BEGIN
							SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL;
						END
					
					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
					
					SET @sExecString = @sExecString + ')';

					/* Add the reverse order code */
					IF len(@sReverseOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL + ')';
						END
					
					SET @sExecString = @sExecString

					/* Add the order code */
					IF len(@sOrderSQL) > 0 
						BEGIN
							SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;
						END
		
					/* Set the position variables */
					SET @piFirstRecPos = @piFirstRecPos - @iGetCount;
					IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1;
					SET @pfFirstPage = CASE WHEN @piFirstRecPos = 1 
															THEN 1
															ELSE 0
														 END;
					SET @pfLastPage = CASE WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount 
															THEN 1
															ELSE 0
														END;
				END

		END

	EXECUTE sp_executeSQL @sExecString;
END

GO

CREATE PROCEDURE [dbo].[spASRIntGetEventLogBatchDetails] (
	@piBatchRunID 	integer,
	@piEventID		integer
)
AS
BEGIN

	SET NOCOUNT ON;
	
	DECLARE @sExecString		nvarchar(MAX),
			@sSelectString 		varchar(MAX),
			@sFromString		varchar(MAX),
			@sWhereString		varchar(MAX),
			@sOrderString 		varchar(MAX);

	SET @sSelectString = '';
	SET @sFromString = '';
	SET @sWhereString = '';
	SET @sOrderString = '';

	/* create SELECT statment string */
	SET @sSelectString = 'SELECT 
		 ID, 
		 DateTime,
		 EndTime,
		 IsNull(Duration,-1) AS Duration,
		 Username,
		 CASE Type 
						WHEN 0 THEN ''Unknown''
						WHEN 1 THEN ''Cross Tab'' 
						WHEN 2 THEN ''Custom Report'' 
						WHEN 3 THEN ''Data Transfer'' 
						WHEN 4 THEN ''Export'' 
						WHEN 5 THEN ''Global Add'' 
						WHEN 6 THEN ''Global Delete'' 
						WHEN 7 THEN ''Global Update'' 
						WHEN 8 THEN ''Import'' 
						WHEN 9 THEN ''Mail Merge'' 
						WHEN 10 THEN ''Diary Delete'' 
						WHEN 11 THEN ''Diary Rebuild''
						WHEN 12 THEN ''Email Rebuild''
						WHEN 13 THEN ''Standard Report''
						WHEN 14 THEN ''Record Editing''
						WHEN 15 THEN ''System Error''
						WHEN 16 THEN ''Match Report''
						WHEN 17 THEN ''Calendar Report''
						WHEN 18 THEN ''Envelopes & Labels''
						WHEN 19 THEN ''Label Definition''
						WHEN 20 THEN ''Record Profile''
						WHEN 21	THEN ''Succession Planning''
						WHEN 22 THEN ''Career Progression''
						WHEN 25 THEN ''Workflow Rebuild''
						WHEN 35 THEN ''9-Box Grid Report''
						ELSE ''Unknown''  
		 END AS Type,
		 Name,
		 CASE Mode 
			WHEN 1 THEN ''Batch''
			WHEN 0 THEN ''Manual''
			ELSE ''Unknown''
		 END AS Mode, 
		 CASE Status 
				WHEN 0 THEN ''Pending''
		   	WHEN 1 THEN ''Cancelled'' 
				WHEN 2 THEN ''Failed'' 
				WHEN 3 THEN ''Successful'' 
				WHEN 4 THEN ''Skipped'' 
				WHEN 5 THEN ''Error''
				ELSE ''Unknown'' 
		 END AS Status,
		 IsNull(BatchName,'''') AS BatchName,
		 IsNull(convert(varchar,SuccessCount), ''N/A'') AS SuccessCount,
		 IsNull(convert(varchar,FailCount), ''N/A'') AS FailCount,
		 IsNull(convert(varchar,BatchJobID), ''N/A'') AS BatchJobID,
		 IsNull(convert(varchar,BatchRunID), ''N/A'') AS BatchRunID';

	SET @sFromString = ' FROM ASRSysEventLog ';

	IF @piBatchRunID > 0
		BEGIN
			SET @sWhereString = ' WHERE BatchRunID = ' + convert(varchar, @piBatchRunID);
		END
	ELSE
		BEGIN
			SET @sWhereString = ' WHERE ID = ' + convert(varchar, @piEventID);
		END

	SET @sOrderString = ' ORDER BY DateTime ASC ';

	SET @sExecString = @sSelectString + @sFromString + @sWhereString + @sOrderString;

	-- Run generated statement
	EXEC sp_executeSQL @sExecString;
	
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetEventLogEmailInfo]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].spASRIntGetEventLogEmailInfo;
GO

CREATE PROCEDURE [dbo].[spASRIntGetEventLogEmailInfo] (
	@psSelectedIDs	varchar(MAX),
	@psSubject		varchar(MAX) OUTPUT,
	@psOrderColumn	varchar(MAX),
	@psOrderOrder	varchar(MAX)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sSelectString 	nvarchar(MAX),
			@sOrderSQL		varchar(MAX);
	
	/* Clean the input string parameters. */
	IF len(@psSelectedIDs) > 0 SET @psSelectedIDs = replace(@psSelectedIDs, '''', '''''');
	IF len(@psOrderColumn) > 0 SET @psOrderColumn = replace(@psOrderColumn, '''', '''''');
	IF len(@psOrderOrder) > 0 SET @psOrderOrder = replace(@psOrderOrder, '''', '''''');

	SELECT @psSubject = IsNull(SettingValue,'<<Unknown Customer>>') + ' - Event Log' 
	FROM ASRSysSystemSettings 
	WHERE Lower(Section) = 'licence' 
		AND Lower(SettingKey) = 'customer name';

	SET @sSelectString = '';

	/* create SELECT statment string */
	SET @sSelectString = 'SELECT 	A.ID, 
		A.Name, 
		A.DateTime,
		A.EndTime,
		IsNull(A.Duration,-1) AS Duration, 
		A.Username, 
		CASE A.Mode 
			WHEN 1 THEN ''Batch'' 
			ELSE ''Manual'' 
		END AS ''Mode'', 
		CASE A.Status 
			WHEN 0 THEN ''Pending''
		  WHEN 1 THEN ''Cancelled'' 
			WHEN 2 THEN ''Failed'' 
			WHEN 3 THEN ''Successful'' 
			WHEN 4 THEN ''Skipped'' 
			WHEN 5 THEN ''Error''
			ELSE ''Unknown'' 
		END AS Status, 
		CASE A.Type 
			WHEN 0 THEN ''Unknown''
			WHEN 1 THEN ''Cross Tab'' 
			WHEN 2 THEN ''Custom Report'' 
			WHEN 3 THEN ''Data Transfer'' 
			WHEN 4 THEN ''Export'' 
			WHEN 5 THEN ''Global Add'' 
			WHEN 6 THEN ''Global Delete'' 
			WHEN 7 THEN ''Global Update'' 
			WHEN 8 THEN ''Import'' 
			WHEN 9 THEN ''Mail Merge'' 
			WHEN 10 THEN ''Diary Delete'' 
			WHEN 11 THEN ''Diary Rebuild''
			WHEN 12 THEN ''Email Rebuild''
			WHEN 13 THEN ''Standard Report''
			WHEN 14 THEN ''Record Editing''
			WHEN 15 THEN ''System Error''
			WHEN 16 THEN ''Match Report''
			WHEN 17 THEN ''Calendar Report''
			WHEN 18 THEN ''Envelopes & Labels''
			WHEN 19 THEN ''Label Definition''
			WHEN 20 THEN ''Record Profile''
			WHEN 21	THEN ''Succession Planning''
			WHEN 22 THEN ''Career Progression''
			WHEN 25 THEN ''Workflow Rebuild''
			WHEN 35 THEN ''9-Box Grid Report''
			ELSE ''Unknown''  
		END AS Type,
		CASE 
			WHEN A.SuccessCount IS NULL THEN ''N/A''
			ELSE CONVERT(varchar, A.SuccessCount)
		END AS SuccessCount,
		CASE
			WHEN A.FailCount IS NULL THEN ''N/A''
			ELSE CONVERT(varchar, A.FailCount)
		END AS FailCount,
		A.BatchName AS BatchName,
		A.BatchJobID AS BatchJobID,
		A.BatchRunID AS BatchRunID,
		B.Notes, 
		B.ID AS ''DetailsID'' ,
		(SELECT count(ID) 
			FROM ASRSysEventLogDetails C 
			WHERE C.EventLogID = A.ID) as ''count''
		FROM ASRSysEventLog A
		LEFT OUTER JOIN ASRSysEventLogDetails B
			ON A.ID = B.EventLogID
		WHERE A.ID IN (' + @psSelectedIDs + ')';

	SET @sOrderSQL = '';
	
	IF @psOrderColumn = 'Type'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Type] 
				WHEN 1 THEN ''Cross Tab''
				WHEN 2 THEN ''Custom Report''
				WHEN 3 THEN ''Data Transfer''
				WHEN 4 THEN ''Export''
				WHEN 5 THEN ''Global Add''
				WHEN 6 THEN ''Global Delete''
				WHEN 7 THEN ''Global Update''
				WHEN 8 THEN ''Import''
				WHEN 9 THEN ''Mail Merge''
				WHEN 10 THEN ''Diary Delete''
				WHEN 11 THEN ''Diary Rebuild''
				WHEN 12 THEN ''Email Rebuild''
				WHEN 13 THEN ''Standard Report''
				WHEN 14 THEN ''Record Editing''
				WHEN 15 THEN ''System Error''
				WHEN 16 THEN ''Match Report''
				WHEN 17 THEN ''Calendar Report''
				WHEN 18 THEN ''Envelopes & Labels''
				WHEN 19 THEN ''Label Definition''
				WHEN 20 THEN ''Record Profile''
				WHEN 21 THEN ''Succession Planning''
				WHEN 22 THEN ''Career Progression''
				WHEN 25 THEN ''Workflow Rebuild''
				WHEN 35 THEN ''9-Box Grid Report''
				ELSE ''Unknown''
			END ';
	END

	IF @psOrderColumn = 'Mode'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Mode] 
				WHEN 1 THEN ''Batch''
				WHEN 0 THEN ''Manual''
				ELSE ''Unknown''
			END ';
	END
	
	IF @psOrderColumn = 'Status'
	BEGIN
		SET @sOrderSQL = 
			' CASE [Status] 
				WHEN 0 THEN ''Pending''
				WHEN 1 THEN ''Cancelled''
				WHEN 2 THEN ''Failed''
				WHEN 3 THEN ''Successful''
				WHEN 4 THEN ''Skipped''
				WHEN 5 THEN ''Error''
				ELSE ''Unknown''
			END ';
	END
	
	IF len(@sOrderSQL) = 0
	BEGIN
		SET @sOrderSQL = @psOrderColumn;
	END
	
	SET @sOrderSQL = @sOrderSQL + ' ' + @psOrderOrder;

	IF LEN(LTRIM(RTRIM(@sOrderSQL))) > 0 
	BEGIN
		SET @sSelectString = @sSelectString + ' ORDER BY ' + @sOrderSQL;
	END

	EXEC sp_executeSQL @sSelectString;
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetExpressionDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetExpressionDefinition];
GO

CREATE PROCEDURE [dbo].[sp_ASRIntGetExpressionDefinition] (
	@piExprID		integer,
	@psAction		varchar(100),
	@psErrMsg		varchar(MAX)	OUTPUT,
	@piTimestamp	integer			OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return the defintions of each component and expression in the given expression. */
	DECLARE @sExprIDs		varchar(MAX),
		@sComponentIDs		varchar(MAX),
		@sTempExprIDs		varchar(MAX),
		@sTempComponentIDs	varchar(MAX),
		@sCurrentUser		sysname,
		@iCount				integer,
		@sOwner				varchar(255),
		@sAccess			varchar(MAX),
		@iBaseTableID		integer,
		@sBaseTableID		varchar(100),
		@fSysSecMgr			bit,
		@sExecString		nvarchar(MAX);
	
	SET @psErrMsg = '';
	SET @sCurrentUser = SYSTEM_USER;

	/* Check the expressions exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysExpressions
	WHERE exprID = @piExprID;

	IF @iCount = 0
	BEGIN
		SET @psErrMsg = 'expression has been deleted by another user.';
		RETURN;
	END

	SELECT @sOwner = userName,
		@sAccess = access,
		@iBaseTableID = tableID,
		@piTimestamp = convert(integer, timestamp)
	FROM ASRSysExpressions
	WHERE exprID = @piExprID;

	IF @sAccess <> 'RW'
	BEGIN
		exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;
	
		IF @fSysSecMgr = 1 SET @sAccess = 'RW';
	END
	
	IF @iBaseTableID IS null 
	BEGIN
		SET @sBaseTableID = '0';
	END
	ELSE
	BEGIN
		SET @sBaseTableID = convert(varchar(100), @iBaseTableID);
	END

	/* Check the current user can view the expression. */
	IF (@sAccess = 'HD') AND (@sOwner <> @sCurrentUser) 
	BEGIN
		SET @psErrMsg = 'expression has been made hidden by another user.';
		RETURN;
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@sOwner <> @sCurrentUser) 
	BEGIN
		SET @psErrMsg = 'expression has been made read only by another user.';
		RETURN;
	END

	SET @sExprIDs = convert(varchar(MAX), @piExprID);
	SET @sComponentIDs = '0';

	/* Get a list of the components and sub-expressions in the given expression. */
	exec sp_ASRIntGetSubExpressionsAndComponents @piExprID, @sTempExprIDs OUTPUT, @sTempComponentIDs OUTPUT;

	IF len(@sTempExprIDs) > 0 SET @sExprIDs = @sExprIDs + ',' + @sTempExprIDs;
	IF len(@sTempComponentIDs) > 0 SET @sComponentIDs = @sComponentIDs + ',' + @sTempComponentIDs;

	SET @sExecString = 'SELECT
		''C'' as [type],
		ASRSysExprComponents.componentID AS [id],
		convert(varchar(100), ASRSysExprComponents.componentID)+ char(9) +
		convert(varchar(100), ASRSysExprComponents.exprID)+ char(9) +
		convert(varchar(100), ASRSysExprComponents.type)+ char(9) +
		CASE WHEN ASRSysExprComponents.fieldColumnID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldColumnID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldPassBy IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldPassBy) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionTableID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionTableID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionRecord IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionRecord) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionLine IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionLine) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionOrderID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionOrderID) END + char(9) +
		CASE WHEN ASRSysExprComponents.fieldSelectionFilter IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.fieldSelectionFilter) END + char(9) +
		CASE WHEN ASRSysExprComponents.functionID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.functionID) END + char(9) +
		CASE WHEN ASRSysExprComponents.calculationID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.calculationID) END + char(9) +
		CASE WHEN ASRSysExprComponents.operatorID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.operatorID) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueType) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueCharacter IS null THEN '''' ELSE ASRSysExprComponents.valueCharacter END + char(9) +
		CASE WHEN ASRSysExprComponents.valueNumeric IS null THEN '''' ELSE convert(varchar(100), convert(numeric(38, 8), ASRSysExprComponents.valueNumeric)) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueLogic IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueLogic) END + char(9) +
		CASE WHEN ASRSysExprComponents.valueDate IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.valueDate, 101) END + char(9) +
		CASE WHEN ASRSysExprComponents.promptDescription IS null THEN '''' ELSE ASRSysExprComponents.promptDescription END + char(9) +
		CASE WHEN ASRSysExprComponents.promptMask IS null THEN '''' ELSE ASRSysExprComponents.promptMask END + char(9) +
		CASE WHEN ASRSysExprComponents.promptSize IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptSize) END + char(9) +
		CASE WHEN ASRSysExprComponents.promptDecimals IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptDecimals) END + char(9) +
		CASE WHEN ASRSysExprComponents.functionReturnType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.functionReturnType) END + char(9) +
		CASE WHEN ASRSysExprComponents.lookupTableID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.lookupTableID) END + char(9) +
		CASE WHEN ASRSysExprComponents.lookupColumnID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.lookupColumnID) END + char(9) +
		CASE WHEN ASRSysExprComponents.filterID IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.filterID) END + char(9) +
		CASE WHEN ASRSysExprComponents.expandedNode IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.expandedNode) END + char(9) + 
		CASE WHEN ASRSysExprComponents.promptDateType IS null THEN '''' ELSE convert(varchar(100), ASRSysExprComponents.promptDateType) END + char(9) + 
		CASE 
			WHEN ASRSysExprComponents.type = 1 THEN fldtabs.tablename + 
				CASE 
					WHEN (ASRSysExprComponents.fieldPassBy = 2) OR (ASRSysExprComponents.fieldSelectionRecord <> 5) then '' : '' + fldcols.columnname
					ELSE ''''
				END +
				CASE 
					WHEN ASRSysExprComponents.fieldPassBy = 2 then ''''
					ELSE
						CASE 
							WHEN fldrelations.parentID IS null THEN ''''
							ELSE
								CASE 
									WHEN ASRSysExprComponents.fieldSelectionRecord = 1 THEN '' (first record''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 2 THEN '' (last record''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 3 THEN '' (line '' + convert(varchar(100), ASRSysExprComponents.fieldSelectionLine)
									WHEN ASRSysExprComponents.fieldSelectionRecord = 4 THEN '' (total''
									WHEN ASRSysExprComponents.fieldSelectionRecord = 5 THEN '' (record count''
									ELSE '' (''
								END +
								CASE 
									WHEN fldorders.name IS null THEN ''''
									ELSE '', order by '''''' + fldorders.name + ''''''''
								END  +
								CASE 
									WHEN fldfilters.name IS null then ''''
									ELSE '', filter by '''''' + fldfilters.name + ''''''''
								END + 
								'')''
						END
				END
			WHEN ASRSysExprComponents.type = 2 THEN ASRSysFunctions.functionName
			WHEN ASRSysExprComponents.type = 3 THEN calcexprs.name
			WHEN ASRSysExprComponents.type = 5 THEN ASRSysOperators.name
			WHEN ASRSysExprComponents.type = 10 THEN filtexprs.name
			ELSE ''''
		END + char(9) +
		CASE WHEN fldcols.tableID IS null THEN '''' ELSE convert(varchar(100), fldcols.tableID) END + char(9) + 
		CASE WHEN fldorders.name IS null THEN '''' ELSE fldorders.name END + char(9) + 
		CASE WHEN fldfilters.name IS null THEN '''' ELSE fldfilters.name END
		AS [definition]
	FROM ASRSysExprComponents
	LEFT OUTER JOIN ASRSysExpressions calcexprs ON ASRSysExprComponents.calculationID = calcexprs.exprID
	LEFT OUTER JOIN ASRSysExpressions filtexprs ON ASRSysExprcomponents.filterID = filtexprs.exprID
	LEFT OUTER JOIN ASRSysColumns fldcols ON ASRSysExprComponents.FieldColumnID = fldcols.columnID
	LEFT OUTER JOIN ASRSysTables fldtabs ON fldcols.tableID = fldtabs.tableID
	LEFT OUTER JOIN ASRSysFunctions ON ASRSysExprComponents.functionID = asrsysfunctions.functionID 
	LEFT OUTER JOIN ASRSysOperators ON ASRSysExprComponents.operatorID = asrsysoperators.operatorID 
	LEFT OUTER JOIN ASRSysRelations fldrelations ON (ASRSysExprComponents.fieldTableID = fldrelations.childID and fldrelations.parentID = ' + @sBaseTableID + ')
	LEFT OUTER JOIN ASRSysOrders fldorders ON ASRSysExprComponents.fieldSelectionOrderID = fldorders.orderID
	LEFT OUTER JOIN ASRSysExpressions fldfilters ON ASRSysExprComponents.fieldSelectionFilter = fldfilters.exprID	
	WHERE ASRSysExprComponents.componentID IN (' + @sComponentIDs + ')
	UNION
	SELECT 	
		''E'' as [type],
		ASRSysExpressions.exprID AS [id],
		convert(varchar(100), ASRSysExpressions.exprID)+ char(9) +
		ASRSysExpressions.name + char(9) +
		convert(varchar(100), ASRSysExpressions.tableID) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnType) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnSize) + char(9) +
		convert(varchar(100), ASRSysExpressions.returnDecimals) + char(9) +
		convert(varchar(100), ASRSysExpressions.type) + char(9) +
		convert(varchar(100), ASRSysExpressions.parentComponentID) + char(9) +
		ASRSysExpressions.userName + char(9) +
		ASRSysExpressions.access + char(9) +
		CASE WHEN ASRSysExpressions.description IS null THEN '''' ELSE ASRSysExpressions.description END + char(9) +
		convert(varchar(100), convert(integer, ASRSysExpressions.timestamp)) + char(9) + 
		convert(varchar(100), isnull(ASRSysExpressions.viewInColour, 0)) + char(9) +
		convert(varchar(100), isnull(ASRSysExpressions.expandedNode, 0)) AS [definition]
	FROM ASRSysExpressions
	WHERE ASRSysExpressions.exprID IN (' + @sExprIDs + ')
	ORDER BY [id]';
	
	EXECUTE sp_EXecuteSQL @sExecString;
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntCurrentUserAccess]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntCurrentUserAccess];
GO

CREATE PROCEDURE [dbo].[spASRIntCurrentUserAccess] (
	@piUtilityType	integer,
	@plngID			integer,
	@psAccess		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@sTableName			sysname,
		@sAccessTableName	sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@sRoleName			varchar(255),
		@sActualUserName	sysname,
		@iActualUserGroupID	integer,
		@fEnabled			bit

	SET @sTableName = '';
	SET @psAccess = 'HD';

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iActualUserGroupID OUTPUT;
					
	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sAccessTableName = 'ASRSysBatchJobAccess';
		SET @sIDColumnName = 'ID';
 	END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sAccessTableName = 'ASRSysCalendarReportAccess';
		SET @sIDColumnName = 'ID';
 	END

	IF @piUtilityType = 1 OR @piUtilityType = 35 /* Cross Tab */
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sAccessTableName = 'ASRSysCrossTabAccess';
		SET @sIDColumnName = 'CrossTabID';
 	END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sAccessTableName = 'ASRSysCustomReportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
    
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sAccessTableName = 'ASRSysDataTransferAccess';
		SET @sIDColumnName = 'DataTransferID';
 	END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sAccessTableName = 'ASRSysExportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sAccessTableName = 'ASRSysGlobalAccess';
		SET @sIDColumnName = 'functionID';
 	END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sAccessTableName = 'ASRSysImportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sAccessTableName = 'ASRSysMailMergeAccess';
		SET @sIDColumnName = 'mailMergeID';
 	END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sAccessTableName = 'ASRSysRecordProfileAccess';
		SET @sIDColumnName = 'recordProfileID';
 	END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sAccessTableName = 'ASRSysMatchReportAccess';
		SET @sIDColumnName = 'matchReportID';
 	END

	IF (@piUtilityType = 25) /* Workflow */
	BEGIN
		SELECT @fEnabled = enabled
		FROM [dbo].[ASRSysWorkflows]
		WHERE ID = @plngID;
		
		IF @fEnabled = 1
		BEGIN
			SET @psAccess = 'RW';
		END
	END

	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @sValue = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
								OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
								AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
						WHEN ' + @sTableName + '.userName = system_user THEN ''RW''
						ELSE
							CASE
								WHEN ' + @sAccessTableName + '.access IS null THEN ''HD''
								ELSE ' + @sAccessTableName + '.access 
							END
						END
					FROM sysusers b
					INNER JOIN sysusers a ON b.uid = a.gid
					LEFT OUTER JOIN ' + @sAccessTableName + ' ON (b.name = ' + @sAccessTableName + '.groupName
						AND ' + @sAccessTableName + '.id = ' + convert(nvarchar(100), @plngID) + ')
					INNER JOIN ' + @sTableName + ' ON ' + @sAccessTableName + '.ID = ' + @sTableName + '.' + @sIDColumnName + '
					WHERE b.Name = ''' + @sRoleName + '''';

		SET @sParamDefinition = N'@sValue varchar(MAX) OUTPUT';
		EXEC sp_executesql @sSQL,  @sParamDefinition, @psAccess OUTPUT;
	END

	IF @psAccess IS null SET @psAccess = 'HD';
END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntCurrentAccessForRole]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntCurrentAccessForRole];
GO

CREATE PROCEDURE [dbo].[spASRIntCurrentAccessForRole] (
	@psRoleName		sysname,
	@piUtilityType	integer,
	@plngID			integer,
	@psAccess		varchar(2)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@sTableName			sysname,
		@sAccessTableName	sysname,
		@sIDColumnName		sysname,
		@sSQL				nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@fEnabled			bit;

	SET @sTableName = '';
	SET @psAccess = 'HD';

	IF @piUtilityType = 0 /* Batch Job */
	BEGIN
		SET @sTableName = 'ASRSysBatchJobName';
		SET @sAccessTableName = 'ASRSysBatchJobAccess';
		SET @sIDColumnName = 'ID';
 	END

	IF @piUtilityType = 17 /* Calendar Report */
	BEGIN
		SET @sTableName = 'ASRSysCalendarReports';
		SET @sAccessTableName = 'ASRSysCalendarReportAccess';
		SET @sIDColumnName = 'ID';
 	END

	IF @piUtilityType = 1 OR @piUtilityType = 35 /* Cross Tab */
	BEGIN
		SET @sTableName = 'ASRSysCrossTab';
		SET @sAccessTableName = 'ASRSysCrossTabAccess';
		SET @sIDColumnName = 'CrossTabID';
 	END
    
	IF @piUtilityType = 2 /* Custom Report */
	BEGIN
		SET @sTableName = 'ASRSysCustomReportsName';
		SET @sAccessTableName = 'ASRSysCustomReportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
    
	IF @piUtilityType = 3 /* Data Transfer */
	BEGIN
		SET @sTableName = 'ASRSysDataTransferName';
		SET @sAccessTableName = 'ASRSysDataTransferAccess';
		SET @sIDColumnName = 'DataTransferID';
 	END
    
	IF @piUtilityType = 4 /* Export */
	BEGIN
		SET @sTableName = 'ASRSysExportName';
		SET @sAccessTableName = 'ASRSysExportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
	IF (@piUtilityType = 5) OR (@piUtilityType = 6) OR (@piUtilityType = 7) /* Globals */
	BEGIN
		SET @sTableName = 'ASRSysGlobalFunctions';
		SET @sAccessTableName = 'ASRSysGlobalAccess';
		SET @sIDColumnName = 'functionID';
 	END
    
	IF (@piUtilityType = 8) /* Import */
	BEGIN
		SET @sTableName = 'ASRSysImportName';
		SET @sAccessTableName = 'ASRSysImportAccess';
		SET @sIDColumnName = 'ID';
 	END
    
	IF (@piUtilityType = 9) OR (@piUtilityType = 18) /* Label or Mail Merge */
	BEGIN
		SET @sTableName = 'ASRSysMailMergeName';
		SET @sAccessTableName = 'ASRSysMailMergeAccess';
		SET @sIDColumnName = 'mailMergeID';
 	END
    
	IF (@piUtilityType = 20) /* Record Profile */
	BEGIN
		SET @sTableName = 'ASRSysRecordProfileName';
		SET @sAccessTableName = 'ASRSysRecordProfileAccess';
		SET @sIDColumnName = 'recordProfileID';
 	END
    
	IF (@piUtilityType = 14) OR (@piUtilityType = 23) OR (@piUtilityType = 24) /* Match Report, Succession, Career */
	BEGIN
		SET @sTableName = 'ASRSysMatchReportName';
		SET @sAccessTableName = 'ASRSysMatchReportAccess';
		SET @sIDColumnName = 'matchReportID';
 	END

	IF (@piUtilityType = 25) /* Workflow */
	BEGIN
		SELECT @fEnabled = enabled
		FROM [dbo].[ASRSysWorkflows]
		WHERE ID = @plngID;
		
		IF @fEnabled = 1
		BEGIN
			SET @psAccess = 'RW';
		END
	END

	IF len(@sTableName) > 0
	BEGIN
		SET @sSQL = 'SELECT @sValue = 
					CASE
						WHEN (SELECT count(*)
							FROM ASRSysGroupPermissions
							INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID
								AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
								OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER''))
							INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
								AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS'')
							WHERE b.Name = ASRSysGroupPermissions.groupname
								AND ASRSysGroupPermissions.permitted = 1) > 0 THEN ''RW''
						WHEN ' + @sTableName + '.userName = system_user THEN ''RW''
						ELSE
							CASE
								WHEN ' + @sAccessTableName + '.access IS null THEN ''HD''
								ELSE ' + @sAccessTableName + '.access 
							END
						END
					FROM sysusers b
					INNER JOIN sysusers a ON b.uid = a.gid
					LEFT OUTER JOIN ' + @sAccessTableName + ' ON (b.name = ' + @sAccessTableName + '.groupName
						AND ' + @sAccessTableName + '.id = ' + convert(nvarchar(100), @plngID) + ')
					INNER JOIN ' + @sTableName + ' ON ' + @sAccessTableName + '.ID = ' + @sTableName + '.' + @sIDColumnName + '
					WHERE b.Name = ''' + @psRoleName + ''''

		SET @sParamDefinition = N'@sValue varchar(MAX) OUTPUT';
		EXEC sp_executesql @sSQL,  @sParamDefinition, @psAccess OUTPUT;
	END

	IF @psAccess IS null SET @psAccess = 'HD';
END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetExprFunctions]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetExprFunctions];
GO

CREATE PROCEDURE [dbo].[spASRIntGetExprFunctions] (
	@piTableID 		integer,
	@pbAbsenceEnabled	bit
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of tab-delimited runtime function definitions ;
	<function id><tab><function name><tab><function category> */
	DECLARE @iTemp 					integer,
		@iPersonnelTableID		integer,
		@iHierarchyTableID		integer,
		@iPostAllocationTableID	integer,
		@iIdentifyingColumnID	integer,
		@iReportsToColumnID		integer,
		@iLoginColumnID			integer,
		@iSecondLoginColumnID	integer,
		@fIsPostSubOfOK			bit = 0,
		@fIsPostSubOfUserOK		bit = 0,
		@fIsPersSubOfOK			bit = 0,
		@fIsPersSubOfUserOK		bit = 0,
		@fHasPostSubOK			bit = 0,
		@fHasPostSubUserOK		bit = 0,
		@fHasPersSubOK			bit = 0,
		@fHasPersSubUserOK		bit = 0,
		@fPostBased				bit = 0, 
		@sSQLVersion			integer,
		@fBaseTablePersonnelOK	bit = 0;
	
	SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_TablePersonnel';

	SELECT @iHierarchyTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_TableHierarchy';

	SELECT @iIdentifyingColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_FieldIdentifier';

	SELECT @iReportsToColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_FieldReportsTo';

	SELECT @iPostAllocationTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_TablePostAllocation';

	SELECT @iLoginColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_FieldsLoginName';

	SELECT @iSecondLoginColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_FieldsSecondLoginName';

	IF (@iLoginColumnID = 0) AND (@iSecondLoginColumnID > 0)
	BEGIN
		SET @iLoginColumnID = @iSecondLoginColumnID;
		SET @iSecondLoginColumnID = 0;
	END

	IF @iPersonnelTableID <> @iHierarchyTableID SET @fPostBased = 1;
	IF @iPersonnelTableID = @piTableID 
	BEGIN
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			((@fPostBased = 0) OR (@iPersonnelTableID > 0)) AND
			((@fPostBased = 0) OR (@iPostAllocationTableID > 0)) 
		BEGIN
			SET @fIsPersSubOfOK = 1;
			SET @fHasPersSubOK = 1;
		END
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@iPersonnelTableID > 0) AND
			(@iLoginColumnID > 0) AND
			((@fPostBased = 0) OR (@iPostAllocationTableID > 0)) 
		BEGIN
			SET @fIsPersSubOfUserOK = 1;
			SET @fHasPersSubUserOK = 1;
		END
	END
				
	IF @iHierarchyTableID = @piTableID 
	BEGIN
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@fPostBased = 1)
		BEGIN
			SET @fIsPostSubOfOK = 1;
			SET @fHasPostSubOK = 1;
		END
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@iPersonnelTableID > 0) AND
			(@iLoginColumnID > 0) AND
			(@fPostBased = 1) AND
			(@iPostAllocationTableID > 0)
		BEGIN
			SET @fIsPostSubOfUserOK = 1;
			SET @fHasPostSubUserOK = 1;
		END
	END
	IF @iPersonnelTableID = @piTableID 
	BEGIN
		SET @fBaseTablePersonnelOK = 1;
	END
	ELSE
	BEGIN
		SELECT @iTemp = COUNT(*)
		FROM ASRSysRelations
		WHERE parentID = @iPersonnelTableID
			AND childID = @piTableID;
		IF @iTemp > 0
		BEGIN
			SET @fBaseTablePersonnelOK = 1;
		END
	END

	SELECT 
		convert(varchar(255), functionID) + char(9) +
		functionName + 
		CASE 
			WHEN len(shortcutKeys) > 0 THEN ' ' + shortcutKeys
			ELSE ''
		END + char(9) +
		category AS [definitionString]
	FROM ASRSysFunctions
	WHERE (runtime = 1 OR UDF = 1)
		AND ((functionID <> 65) OR (@fIsPostSubOfOK = 1))
		AND ((functionID <> 66) OR (@fIsPostSubOfUserOK = 1))
		AND ((functionID <> 67) OR (@fIsPersSubOfOK = 1))
		AND ((functionID <> 68) OR (@fIsPersSubOfUserOK = 1))
		AND ((functionID <> 69) OR (@fHasPostSubOK = 1))
		AND ((functionID <> 70) OR (@fHasPostSubUserOK = 1))
		AND ((functionID <> 71) OR (@fHasPersSubOK = 1))
		AND ((functionID <> 72) OR (@fHasPersSubUserOK = 1))
		AND ((functionID <> 30) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 46) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 47) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 73) OR ((@fBaseTablePersonnelOK = 1) AND (@pbAbsenceEnabled = 1)));
END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntValidateTrainingBooking]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntValidateTrainingBooking];
GO

CREATE PROCEDURE [dbo].[sp_ASRIntValidateTrainingBooking] (
	@piResultCode		varchar(MAX) OUTPUT,
	@piEmpRecID		integer,
	@piCourseRecID		integer,
	@psBookingStatus	varchar(MAX),
	@piTBRecID		integer,
	@psCourseOverbooked integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Perform the Training Booking validation on the given insert/update SQL string.
	Return codes are :
		@piResultCode = '000' - completely valid
		If non-zero then the result code is composed as abc,
		where a is the result of the PRE-REQUISITES check
			b is the result of the AVAILABILITY check
			c is the result of the OVERLAPPED BOOKING check.
		the values of which can be :
			0 if the check PASSED
			1 if the check FAILED and CANNOT be overridden
			2 if the check FAILED but CAN be overridden

	The psCourseOverbooked parameter returns if the course is overbooked
	*/
	DECLARE	@fIncludeProvisionals	bit,
		@sIncludeProvisionals	varchar(MAX),
		@iCount					integer,
		@iResult				integer,
		@iTemp					integer,
		@piResultOverlapping   integer = 0,
		@piResultPrerequisites	integer = 0,
		@piResultUnavailability	integer = 0;

	SET @piResultCode = '';
	SET @psCourseOverbooked = 0;

	IF (@piCourseRecID > 0) AND ((@psBookingStatus = 'B') OR (@psBookingStatus = 'P'))
	BEGIN  
		SELECT @sIncludeProvisionals = parameterValue
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseIncludeProvisionals'
		IF @sIncludeProvisionals IS NULL SET @sIncludeProvisionals = 'FALSE'
		IF @sIncludeProvisionals = 'FALSE'
		BEGIN
			SET @sIncludeProvisionals = 0
		END
		ELSE
		BEGIN
			SET @sIncludeProvisionals = 1
		END

		/* Only check that the selected course is not fully booked if the new booking is included in the number booked. */
		IF (@fIncludeProvisionals = 1) OR (@psBookingStatus = 'B') 
		BEGIN
			/* Check if the overbooking stored procedure exists. */
			SELECT @iCount = COUNT(*) 
			FROM sysobjects
			WHERE id = object_id('sp_ASR_TBCheckOverbooking')
				AND sysstat & 0xf = 4

			IF @iCount > 0
			BEGIN
				exec sp_ASR_TBCheckOverbooking @piCourseRecID, @piTBRecID, 1, @iResult OUTPUT
				SET @psCourseOverbooked = @iResult -- @iResult = 1 -> Course fully booked (error). @iResult = 2 -> Course fully booked (over-rideable by the user).
			END
		END
      
		IF @piEmpRecID > 0
		BEGIN
			/* Check that the employee has satisfied the pre-requisite criteria for the selected course. */
			/* First check if the pre-requisite table is configured. If not, we do not need to do the pre-req check. */
			SELECT @iTemp = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_PreReqTable'
			IF @iTemp IS NULL SET @iTemp = 0

			IF @iTemp > 0 
			BEGIN
				/* Check if the pre-req stored procedure exists. */
				SELECT @iCount = COUNT(*) 
				FROM sysobjects
				WHERE id = object_id('sp_ASR_TBCheckPreRequisites')
					AND sysstat & 0xf = 4

				IF @iCount > 0
				BEGIN
					exec sp_ASR_TBCheckPreRequisites @piCourseRecID, @piEmpRecID, @iResult OUTPUT
					SET @piResultPrerequisites = @iResult -- @iResult = 1 -> Pre-requisites not satisfied (error). @iResult = 2 -> Pre-requisites not satisfied (over-rideable by the user). 
				END
			END

			/* Check that the employee is available for the selected course. */
			/* First check if the unavailability table is configured. If not, we do not need to do the unavailability check. */
			SELECT @iTemp = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_UnavailTable'
			IF @iTemp IS NULL SET @iTemp = 0

			IF @iTemp > 0 
			BEGIN
				/* Check if the unavailability stored procedure exists. */
				SELECT @iCount = COUNT(*) 
				FROM sysobjects
				WHERE id = object_id('sp_ASR_TBCheckUnavailability')
					AND sysstat & 0xf = 4

				IF @iCount > 0
				BEGIN
					exec sp_ASR_TBCheckUnavailability @piCourseRecID, @piEmpRecID, @iResult OUTPUT
					SET @piResultUnavailability = @iResult -- @iResult = 1 -> Employee unavailable (error). @iResult = 2 -> Employee unavailable (over-rideable by the user).
				END
			END

			/* Check if the overlapped booking stored procedure exists. */
			SELECT @iCount = COUNT(*) 
			FROM sysobjects
			WHERE id = object_id('sp_ASR_TBCheckOverlappedBooking')
				AND sysstat & 0xf = 4

			IF @iCount > 0
			BEGIN
				exec sp_ASR_TBCheckOverlappedBooking @piCourseRecID, @piEmpRecID, @piTBRecID, @iResult OUTPUT
				SET @piResultOverlapping = @iResult -- @iResult = 1 -> Overlapped booking (error). @iResult = 2 -> Overlapped booking (over-rideable by the user). 
			END
		END
		SET @piResultCode = CONVERT(VARCHAR(1), @piResultPrerequisites) + CONVERT(VARCHAR(1), @piResultUnavailability) + CONVERT(VARCHAR(1), @piResultOverlapping)
	END
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetColumnControlValues]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetColumnControlValues];
GO

CREATE PROCEDURE [dbo].[spASRIntGetColumnControlValues]
	@ColumnIDs nvarchar(100)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @sql nvarchar(MAX);

	SET @sql = 'SELECT columnID, Value, sequence FROM ASRSysColumnControlValues WHERE columnID IN (' + @ColumnIDs + ')'
	EXECUTE sp_executeSQL @sql;
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetFindRecords]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetFindRecords];
GO

CREATE PROCEDURE [dbo].[spASRIntGetFindRecords] (
	@pfError 				bit 			OUTPUT, 
	@pfSomeSelectable 		bit 			OUTPUT, 
	@pfSomeNotSelectable 	bit 			OUTPUT, 
	@psRealSource			varchar(255)	OUTPUT,
	@pfInsertGranted		bit				OUTPUT,
	@pfDeleteGranted		bit				OUTPUT,
	@piTableID 				integer, 
	@piViewID 				integer, 
	@piOrderID 				integer, 
	@piParentTableID		integer,
	@piParentRecordID		integer,
	@psFilterDef			varchar(MAX),
	@piRecordsRequired		integer,
	@pfFirstPage			bit				OUTPUT,
	@pfLastPage				bit				OUTPUT,
	@psLocateValue			varchar(MAX),
	@piColumnType			integer			OUTPUT,
	@piColumnSize			integer			OUTPUT,
	@piColumnDecimals		integer			OUTPUT,
	@psAction				varchar(255),
	@piTotalRecCount		integer			OUTPUT,
	@piFirstRecPos			integer			OUTPUT,
	@piCurrentRecCount		integer,
	@psDecimalSeparator		varchar(255),
	@psLocaleDateFormat		varchar(255),
	@RecordID				integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the find records for the current user, given the table/view and order IDs.
		@pfError = 1 if errors occured in getting the find records. Else 0.
		@pfSomeSelectable = 1 if some find columns were selectable. Else 0.
		@pfSomeNotSelectable = 1 if some find columns were NOT selectable. Else 0.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@piParentTableID = the ID of the parent table.
		@piParentRecordID = the ID of the associated record in the parent table.
	*/
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 			varchar(10),
		@fSelectGranted 	bit,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString			varchar(MAX),
		@sViewName 			varchar(255),
		@sExecString		nvarchar(MAX),
		@sTempString		varchar(MAX),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iDataType 			integer,
		@iTempAction		integer,
		@iTemp				integer,
		@sRemainingSQL		varchar(MAX),
		@iLastCharIndex		integer,
		@iCharIndex 		integer,
		@sDESCstring		varchar(5),
		@sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@fFirstColumnAsc	bit,
		@sFirstColCode		varchar(MAX),
		@sLocateCode		varchar(MAX),
		@iCount				integer,
		@iGetCount			integer,
		@iSize				integer,
		@iDecimals			integer,
		@bUse1000Separator	bit,
		@sActualLoginName	varchar(250),
		@iIndex1			integer,
		@iIndex2			integer,
		@iIndex3			integer,
		@iColumnID			integer,
		@iOperatorID		integer,
		@sValue				varchar(MAX),
		@sFilterSQL			nvarchar(MAX),
		@sTempLocateValue	varchar(MAX),
		@sSubFilterSQL		nvarchar(MAX),
		@bBlankIfZero		bit,
		@sTempFilterString	varchar(MAX),
		@sJoinSQL			nvarchar(max),
		@psOriginalAction		varchar(255),
		@sThousandColumns		varchar(255),
		@sBlankIfZeroColumns	varchar(255),
		@sTableOrViewName varchar(255);

	DECLARE @FindDefinition TABLE(tableID integer, columnID integer, columnName nvarchar(255), tableName nvarchar(255)
									, ascending bit, type varchar(1), datatype integer, controltype integer, size integer, decimals integer, Use1000Separator bit, BlankIfZero bit, Editable bit
									, LookupTableID integer, LookupColumnID integer, LookupFilterColumnID integer, LookupFilterValueID integer, SpinnerMinimum smallint, SpinnerMaximum smallint, SpinnerIncrement smallint
									, DefaultValue varchar(max)
									)

	DECLARE @OriginalColumns TABLE(columnID integer, columnName nvarchar(255))

	/* Clean the input string parameters. */
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''');
	/* Initialise variables. */
	SET @pfError = 0;
	SET @pfSomeSelectable = 0;
	SET @pfSomeNotSelectable = 0;
	SET @sDESCstring = ' DESC';
	SET @fFirstColumnAsc = 1;
	SET @sFirstColCode = '';
	SET @piColumnSize = 0;
	SET @piColumnDecimals = 0;
	SET @bUse1000Separator = 0;
	SET @bBlankIfZero = 0;
	SET @sThousandColumns = '';
	SET @sTableOrViewName = '';
	SET @sBlankIfZeroColumns = '';

	SET @sRealSource = '';
	SET @sSelectSQL = '';
	SET @sOrderSQL = '';
	SET @sReverseOrderSQL = '';
	SET @fSelectDenied = 0;
	SET @sExecString = '';
	SET @sFilterSQL = '';
	SET @sTempLocateValue = '';
	
	SET @psOriginalAction = @psAction;
	
	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000;
	SET @psAction = UPPER(@psAction);
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE') AND 
		(@psAction <> 'LOCATEID')
	BEGIN
		SET @psAction = 'MOVEFIRST';
	END
	
	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	-- Create a temporary view of the sysprotects table
	DECLARE @SysProtects TABLE([ID] int, Columns varbinary(8000)
								, [Action] tinyint
								, ProtectType tinyint);
	INSERT INTO @SysProtects
	SELECT p.[ID], p.[Columns], p.[Action], p.ProtectType FROM ASRSysProtectsCache p WHERE p.UID = @iUserGroupID;
	
	-- Create a temporary table to hold the tables/views that need to be joined.
	DECLARE @JoinParents TABLE(tableViewName sysname, tableID integer);

	-- Create a temporary table of the 'select' column permissions for all tables/views used in the order.
	DECLARE @ColumnPermissions TABLE(tableID integer,
								tableViewName	sysname,
								columnName	sysname,
								selectGranted	bit,
								updateGranted	bit);
								
	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM [dbo].[ASRSysTables]
		WHERE ASRSysTables.tableID = @piTableID;
	
	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1;
		RETURN;
	END
	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM [dbo].[ASRSysViews]
			WHERE viewID = @piViewID;
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM [dbo].[ASRSysChildViews2]
		WHERE tableID = @piTableID
			AND role = @sUserGroupName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END
	SET @psRealSource = @sRealSource;
	
	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	IF len(@psFilterDef)> 0 
	BEGIN
		WHILE charindex('	', @psFilterDef) > 0
		BEGIN
			SET @sSubFilterSQL = '';
			SET @iIndex1 = charindex('	', @psFilterDef);
			SET @iIndex2 = charindex('	', @psFilterDef, @iIndex1+1);
			SET @iIndex3 = charindex('	', @psFilterDef, @iIndex2+1);
				
			SET @iColumnID = convert(integer, LEFT(@psFilterDef, @iIndex1-1));
			SET @iOperatorID = convert(integer, SUBSTRING(@psFilterDef, @iIndex1+1, @iIndex2-@iIndex1-1));
			SET @sValue = SUBSTRING(@psFilterDef, @iIndex2+1, @iIndex3-@iIndex2-1);
			
			SET @psFilterDef = SUBSTRING(@psFilterDef, @iIndex3+1, LEN(@psFilterDef) - @iIndex3);
			SELECT @iDataType = dataType,
				@sColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @iColumnID;
							
			SET @sColumnName = @sRealSource + '.' + @sColumnName;
			IF (@iDataType = -7) 
			BEGIN
				/* Logic column (must be the equals operator).	*/
				SET @sSubFilterSQL = @sColumnName + ' = ';
			
				IF UPPER(@sValue) = 'TRUE'
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '1';
				END
				ELSE
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '0';
				END
			END
			
			IF ((@iDataType = 2) OR (@iDataType = 4)) 
			BEGIN
				/* Numeric/Integer column. */
				/* Replace the locale decimal separator with '.' for SQL's benefit. */
				SET @sValue = REPLACE(@sValue, @psDecimalSeparator, '.');
				IF (@iOperatorID = 1) 
				BEGIN
					/* Equals. */
					SET @sSubFilterSQL = @sColumnName + ' = ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <> ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' AND ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
        
				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' >= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					SET @sSubFilterSQL = @sColumnName + ' > ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 6) 
				BEGIN
					/* Less than.*/
					SET @sSubFilterSQL = @sColumnName + ' < ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
			END
			IF (@iDataType = 11) 
			BEGIN
				/* Date column. */
				IF LEN(@sValue) > 0
				BEGIN
					/* Convert the locale date into the SQL format. */
					/* Note that the locale date has already been validated and formatted to match the locale format. */
					SET @iIndex1 = CHARINDEX('mm', @psLocaleDateFormat);
					SET @iIndex2 = CHARINDEX('dd', @psLocaleDateFormat);
					SET @iIndex3 = CHARINDEX('yyyy', @psLocaleDateFormat);
						
					SET @sValue = SUBSTRING(@sValue, @iIndex1, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex2, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex3, 4);
				END

				IF (@iOperatorID = 1) 
				BEGIN
					/* Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' = ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END

				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					IF LEN(@sValue) > 0 
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <= ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' >= ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' > ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 6)
				BEGIN
					/* Less than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' < ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
					END
				END
			END
			
			IF ((@iDataType <> -7) AND (@iDataType <> 2) AND (@iDataType <> 4) AND (@iDataType <> 11)) 
			BEGIN
				/* Character/Working Pattern column. */
				IF (@iOperatorID = 1) 
				BEGIN
					/* Equal To. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' = '''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''' + @sValue + '''';
					END
				END
				
				IF (@iOperatorID = 2) 
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> '''' AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''' + @sValue + '''';
					END
				END

				IF (@iOperatorID = 7)
				BEGIN
					/* Contains */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''%' + @sValue + '%''';
					END
				END
				
				IF (@iOperatorID = 8) 
				BEGIN
					/* Does Not Contain. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''%' + @sValue + '%''';
					END
				END
			END
			
			IF LEN(@sSubFilterSQL) > 0
			BEGIN
				/* Add the filter code for this grid record into the complete filter code. */
				IF LEN(@sFilterSQL) > 0
				BEGIN
					SET @sFilterSQL = @sFilterSQL + ' AND (';
				END
				ELSE
				BEGIN
					SET @sFilterSQL = @sFilterSQL + '(';
				END
				SET @sFilterSQL = @sFilterSQL + @sSubFilterSQL + ')';
			END
		END
	END
	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT ASRSysColumns.tableID
		FROM [dbo].[ASRSysOrderItems]
		INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
		WHERE ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 193 THEN 1	ELSE 0 END) AS selectGranted,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 197 THEN 1	ELSE 0 END) AS updateGranted
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action IN (193, 197)
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			GROUP BY syscolumns.name,[protectType];
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 193 THEN 1	ELSE 0 END) AS selectGranted,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 197 THEN 1	ELSE 0 END) AS updateGranted
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action IN (193, 197)
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			GROUP BY sysobjects.name, syscolumns.name,[protectType];

		END
		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/* Create the order select strings. */
	INSERT @FindDefinition
		SELECT c.tableID,
			o.columnID, 
			c.columnName,
	    	t.tableName,
			o.ascending,
			o.type,
			c.dataType,
			c.controltype,
			c.size,
			c.decimals,
			c.Use1000Separator,
			c.BlankIfZero,
			CASE WHEN ISNULL(o.Editable, 0) = 0 OR ISNULL(c.readonly, 1) = 1 THEN 0 ELSE 1 END,
			ISNULL(c.LookupTableID, 0) AS LookupTableID,
			ISNULL(c.LookupColumnID, 0) AS LookupColumnID,
			ISNULL(c.LookupFilterColumnID, 0) AS LookupFilterColumnID,
			ISNULL(c.LookupFilterValueID, 0) AS LookupFilterValueID,
			c.SpinnerMinimum,
			c.SpinnerMaximum,
			c.SpinnerIncrement,
			c.defaultvalue AS DefaultValue
		FROM [dbo].[ASRSysOrderItems] o
		INNER JOIN ASRSysColumns c ON o.columnID = c.columnID
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID
		WHERE o.orderID = @piOrderID
		ORDER BY o.sequence;

	
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableID, columnID, columnName, tableName, ascending, type, datatype, size, decimals, Use1000Separator, BlankIfZero FROM @FindDefinition;


	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END
	WHILE (@@fetch_status = 0)
	BEGIN

		SET @fSelectGranted = 0;
		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;
				
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName;

					SET @sSelectSQL = @sSelectSQL + @sTempString;
					INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
					SET @sThousandColumns = @sThousandColumns + convert(varchar(1),@bUse1000Separator);
					SET @sBlankIfZeroColumns = @sBlankIfZeroColumns + convert(varchar(1),@bBlankIfZero);
					
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							IF @piColumnType = 11 /* Date column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = isnull(convert(varchar(MAX), ' + @sRealSource + '.' + @sColumnName + ', 101), '''')';
							END
							ELSE
							IF @piColumnType = -7 /* Logic column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = replace(isnull(' + @sRealSource + '.' + @sColumnName + ', ''''), '''''''', '''''''''''')' ;
							END
							
							SET @sTempExecString = @sTempExecString
								+ ' FROM ' + @sRealSource
								+ ' WHERE ' + @sRealSource + '.ID = ' + @psLocateValue;
							SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
							EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
						END
					END
					
					SET @sOrderSQL = @sOrderSQL + CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName +
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END;
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1;
			END	
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName;
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + @sColumnTableName + '.' + @sColumnName;
					SET @sSelectSQL = @sSelectSQL + @sTempString;
					INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							SET @sTempLocateValue = '';
						END
					END
					SET @sOrderSQL = @sOrderSQL + CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName + 
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END;
				END
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @JoinParents
				WHERE tableViewName = @sColumnTableName;
				
				IF @iTempCount = 0
				BEGIN
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = '';
				
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT tableViewName
					FROM @ColumnPermissions
					WHERE tableID = @iColumnTableId
						AND tableViewName <> @sColumnTableName
						AND columnName = @sColumnName
						AND selectGranted = 1;
						
				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = 'CASE';
					SET @sSubString = @sSubString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @JoinParents
					WHERE tableViewname = @sViewName;
					
					IF @iTempCount = 0
					BEGIN
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId);
					END
					FETCH NEXT FROM viewCursor INTO @sViewName;
				END

				CLOSE viewCursor;
				DEALLOCATE viewCursor;

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +	' ELSE NULL END';
					IF @sType = 'F'
					BEGIN
						/* Find column. */
						SET @sTempString = CASE 
								WHEN (len(@sSelectSQL) > 0) THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END;
						SET @sSelectSQL = @sSelectSQL + @sTempString;
							
						SET @sTempString = ' AS [' + @sColumnName + ']';
						SET @sSelectSQL = @sSelectSQL + @sTempString;
						INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
						
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType;
							SET @piColumnSize = @iSize;
							SET @piColumnDecimals = @iDecimals;
							SET @fFirstColumnAsc = @fAscending;
							SET @sFirstColCode = @sSubString;
							IF (@psAction = 'LOCATEID')
							BEGIN
								SET @sTempLocateValue = '';
							END
						END
						SET @sOrderSQL = @sOrderSQL + CASE 
								WHEN len(@sOrderSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END + 
							CASE 
								WHEN @fAscending = 0 THEN ' DESC' 
								ELSE '' 
							END;
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1;
				END	
			END
		END
		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;
	END

	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	IF @psAction = 'LOCATEID'
	BEGIN
		SET @psAction = 'LOCATE';
		SET @psLocateValue = @sTempLocateValue;
	END
	
	/* Set the flags that show if no order columns could be selected, or if only some of them could be selected. */
	SET @pfSomeSelectable = CASE WHEN (len(@sSelectSQL) > 0) THEN 1 ELSE 0 END;
	SET @pfSomeNotSelectable = @fSelectDenied;

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN ',' ELSE '' END + 
		@sRealSource + '.ID';

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sRemainingSQL = @sOrderSQL;
		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(',', @sOrderSQL);
		WHILE @iCharIndex > 0 
		BEGIN
 			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', ';
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', ';
			END
			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(',', @sOrderSQL, @iLastCharIndex + 1);
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex);
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring;
	END
	
	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
	IF @piParentTableID > 0 
	BEGIN
		SET @sTempExecString = @sTempExecString + 
			' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
		IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
	END
	ELSE
	BEGIN
		IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
	END
	
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;

	SET @piTotalRecCount = @iCount;
	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sTempString = ',' + @sRealSource + '.ID, CONVERT(int, ' + @sRealSource + '.Timestamp) AS [Timestamp]';
		SET @sSelectSQL = @sSelectSQL + @sTempString;
		INSERT INTO @OriginalColumns VALUES(0, 'ID') -- We don't need the ID
		INSERT INTO @OriginalColumns VALUES(0, 'Timestamp') -- We don't need the ID

		SET @sExecString = 'SELECT ' 
		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
		BEGIN
			SET @sTempString = 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
			SET @sExecString = @sExecString + @sTempString;
		END
	
		SET @sTempString = @sSelectSQL + ' FROM ' + @sRealSource;
		SET @sExecString = @sExecString + @sTempString;

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;
		
		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempString = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
			SET @sExecString = @sExecString + @sTempString;

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		
		CLOSE joinCursor;
		DEALLOCATE joinCursor;
		
		IF (@psAction = 'MOVELAST')
		BEGIN
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'MOVENEXT') 
		BEGIN
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
		END
		IF @psAction = 'MOVEPREVIOUS'
		BEGIN
			IF @piFirstRecPos <= @piRecordsRequired
			BEGIN
				SET @iGetCount = @piFirstRecPos - 1;
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;

		END
		/* Add the filter code. */
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			SET @sExecString = @sExecString + @sTempString;
				
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' AND ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' WHERE ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
		IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sReverseOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'LOCATE')
		BEGIN
			SET @sLocateCode = '';
			IF (@piParentTableID > 0) OR (len(@sFilterSQL) > 0) 
			BEGIN
				SET @sLocateCode = @sLocateCode + ' AND (' + @sFirstColCode;
			END
			ELSE
			BEGIN
				SET @sLocateCode = @sLocateCode + ' WHERE (' + @sFirstColCode;
			END;
			
			SET @sJoinSQL = '';
			DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName, 
				tableID
			FROM @JoinParents;
			OPEN joinCursor;
			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sJoinSQL = @sJoinSQL + 
					' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
				FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			END;
			
			CLOSE joinCursor;
			DEALLOCATE joinCursor;

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = replace(isnull(' + @sFirstColCode + ', ''''), '''''''', '''''''''''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL';
						
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' <= ''' + @sTempLocateValue + '''';
					END;
				END
			END
			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = isnull(convert(varchar(MAX),' + @sFirstColCode + ', 101), '''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sFirstColCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
						END
					END;
				END
				ELSE
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + @sFirstColCode + ' IS NULL';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ''' + @sTempLocateValue + '''';
						END
					END;
				END
			END
			IF @piColumnType = -7 /* Logic column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;
				
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
			END
			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue;
					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + @sTempLocateValue;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL';

					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + @sTempLocateValue ;
					END;
				END
			END
			SET @sLocateCode = @sLocateCode + ')';
			SET @sTempString = @sLocateCode;
			SET @sExecString = @sExecString + @sTempString;
		END

		/* Add the ORDER BY code to the find record selection string if required. */
		IF (@RecordID = -1) BEGIN --Return all records
			SET @sTempString = ' ORDER BY ' + @sOrderSQL;
		END ELSE BEGIN --Return only the requested record
			IF charindex(' WHERE ', @sExecString) > 0 BEGIN --A WHERE clause may have been added for the filtering of records
				SET @sTempString = ' AND '
			END ELSE BEGIN
				SET @sTempString = ' WHERE '
			END

			SET @sTempString = @sTempString + @sRealSource + '.ID = ' + CONVERT(varchar(100), @RecordID) + ' ORDER BY ' + @sOrderSQL;
		END

		SET @sExecString = @sExecString + @sTempString;
				
	END

	/* Check if the user has insert or delete permission on the table. */
	SET @pfInsertGranted = 0;
	SET @pfDeleteGranted = 0;
	IF LEN(@sRealSource) > 0
	BEGIN
		DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT p.action
		FROM @sysprotects p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		WHERE p.protectType <> 206
			AND ((p.action = 195) OR (p.action = 196))
			AND sysobjects.name = @sRealSource;
			
		OPEN tableInfo_cursor;
		FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iTempAction = 195
			BEGIN
				SET @pfInsertGranted = 1;
			END
			ELSE
			BEGIN
				SET @pfDeleteGranted = 1;
			END
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		END
		
		CLOSE tableInfo_cursor;
		DEALLOCATE tableInfo_cursor;
		
	END
	
	/* Set the IsFirstPage, IsLastPage flags, and the page number. */
	IF @psAction = 'MOVEFIRST'
	BEGIN
		SET @piFirstRecPos = 1;
		SET @pfFirstPage = 1;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVENEXT'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount;
		SET @pfFirstPage = 0;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos - @iGetCount;
		IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVELAST'
	BEGIN
		SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1;
		IF @piFirstRecPos < 1 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 1;
	END
	IF @psAction = 'LOCATE'
	BEGIN
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		
		CLOSE joinCursor;
		DEALLOCATE joinCursor;
		
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
		END
		
		SET @sTempExecString = @sTempExecString + @sLocateCode;
		SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iTemp OUTPUT;
		IF @iTemp <=0 
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount + 1;
		END
		ELSE
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount - @iTemp + 1;
		END
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	
	-- Return a recordset of the required columns in the required order from the given table/view.
	IF (@pfSomeSelectable = 1)
	BEGIN
		IF @piViewID <>0 BEGIN
			SELECT @sTableOrViewName = ViewName FROM tbsys_views WHERE ViewID = @piViewID
		END ELSE IF @piTableID <> 0 BEGIN
			SELECT @sTableOrViewName = TableName FROM tbsys_tables WHERE TableID = @piTableID
		END

		SELECT @sBlankIfZeroColumns AS BlankIfZeroColumns
			, @sThousandColumns AS ThousandColumns, @sTableOrViewName AS TableOrViewName

		EXECUTE sp_executeSQL @sExecString;

		DECLARE @IsSingleTable bit = 1;

    	SELECT @IsSingleTable = CASE WHEN COUNT(DISTINCT tableID) = 1 THEN 1 ELSE 0 END
        FROM @FindDefinition;

		SELECT f.tableID, f.columnID, f.columnName, f.ascending, f.type, f.datatype, f.controltype, f.size, f.decimals, f.Use1000Separator, f.BlankIfZero
			 , CASE WHEN f.Editable = 1 AND p.updateGranted = 1 THEN @IsSingleTable ELSE 0 END AS updateGranted
			 , LookupTableID, LookupColumnID, LookupFilterColumnID, LookupFilterValueID
			 ,SpinnerMinimum, SpinnerMaximum, SpinnerIncrement, DefaultValue
			FROM @FindDefinition f
				INNER JOIN @ColumnPermissions p ON p.columnName = f.columnName
			WHERE f.[type] = 'F';

		SELECT columnID, columnName FROM @OriginalColumns ORDER BY columnName

	END

END
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntShowOutOfOfficeHyperlink]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink];
GO

CREATE PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink]	
	(
		@piTableID		integer,
		@piViewID		integer,
		@pfDisplayHyperlink	bit 	OUTPUT
	)
	AS
	BEGIN

		SET NOCOUNT ON;

		SELECT @pfDisplayHyperlink = WFOutOfOffice
			FROM ASRSysSSIViews
			WHERE (TableID = @piTableID) 
				AND (ViewID = @piViewID);

		SELECT @pfDisplayHyperlink = ISNULL(@pfDisplayHyperlink, 0);

	END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntTrackSession]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntTrackSession];
GO

CREATE PROCEDURE dbo.[spASRIntTrackSession](
	@IISServer nvarchar(255),
	@SessionID nvarchar(255),
	@UserName nvarchar(255),
	@SecurityGroup varchar(255),
	@HostName varchar(255),
	@WebArea varchar(20),
	@TrackType tinyint)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @LoginTime	datetime = GETDATE();

	-- Current user tracking
	MERGE INTO ASRSysCurrentSessions AS Target 
	USING (VALUES 
		(@IISServer, @SessionID, @UserName, @HostName, @WebArea) 
	) AS Source (IISServer, SessionID, Username, HostName, WebArea) 
		ON Target.SessionID = Source.SessionID
	WHEN MATCHED AND @TrackType = 1 THEN 
		UPDATE SET webArea = @WebArea, Username = @UserName, HostName = @HostName, IISServer = @IISServer
	WHEN MATCHED AND @TrackType IN (2, 3, 4, 5, 6, 8) THEN
		DELETE
	WHEN NOT MATCHED BY TARGET AND @TrackType = 1 THEN 
		INSERT (IISServer, SessionID, Username, HostName, WebArea)
		VALUES (@IISServer, @SessionID, @UserName, @HostName, @WebArea);

	-- Track in audit log
	IF @TrackType <> 5
		INSERT INTO [dbo].[ASRSysAuditAccess]	(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action) 
			VALUES (@LoginTime, @SecurityGroup, @UserName, @HostName, @WebArea
				, CASE @TrackType
					WHEN 1 THEN 'Log In'
					WHEN 2 THEN 'Log Out'
					WHEN 3 THEN 'Forced Log Out'
					WHEN 8 THEN 'Insufficient Licence'
					ELSE 'Session Timeout'
				END);

END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetSetting]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetSetting];
GO

CREATE PROCEDURE [dbo].[spASRIntGetSetting] (
	@psSection		varchar(MAX),
	@psKey			varchar(MAX),
	@psDefault		varchar(MAX),
	@pfUserSetting	bit,
	@psResult		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return the required user or system setting. */
	DECLARE	@iCount	integer;

	IF @pfUserSetting = 1
	BEGIN
		SELECT @iCount = COUNT(userName)
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = settingValue 
		FROM ASRSysUserSettings
		WHERE userName = SYSTEM_USER
			AND section = @psSection		
			AND settingKey = @psKey;
	END
	ELSE
	BEGIN
		SELECT @iCount = COUNT(settingKey)
		FROM ASRSysSystemSettings
		WHERE section = @psSection		
			AND settingKey = @psKey;

		SELECT @psResult = settingValue 
		FROM ASRSysSystemSettings
		WHERE section = @psSection		
			AND settingKey = @psKey;
	END

	IF @iCount = 0
	BEGIN
		SET @psResult = @psDefault;	
	END
END
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntGetStandardReportDates]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntGetStandardReportDates];
GO

CREATE PROCEDURE [dbo].[spASRIntGetStandardReportDates] (
	@piReportType 		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the prompted values for the given utililty. */
	DECLARE	@sComponents			varchar(MAX),
			@vDateID				varchar(100),
			@iStartDateID			integer,
			@iEndDateID				integer,
			@iStartDateComponentID	integer,
			@iEndDateComponentID	integer;

	/* Create a temp table to hold the propmted value details. */
	DECLARE @promptedValues TABLE(
		componentID			integer,
		promptDescription	varchar(255),
		valueType			integer,
		promptMask			varchar(255),
		promptSize			integer,
		promptDecimals		integer,
		valueCharacter		varchar(255),
		valueNumeric		float,
		valueLogic			bit,
		valueDate			datetime,
		promptDateType		integer, 
		fieldColumnID		integer,
		StartEndType		varchar(5)
	)

	-- Absence Breakdown	
	IF @piReportType = 15
	BEGIN
		EXEC dbo.spASRIntGetSetting 'AbsenceBreakdown', 'Start Date', 0, 0, @vDateID OUTPUT;
		SET @iStartDateID = convert(integer,@vDateID);

		EXEC dbo.spASRIntGetSetting 'AbsenceBreakdown', 'End Date', 0, 0, @vDateID OUTPUT;
		SET @iEndDateID = convert(integer,@vDateID);
	END

	-- Bradford Factor
	IF @piReportType = 16
	BEGIN
		EXEC dbo.spASRIntGetSetting 'BradfordFactor', 'Start Date', 0, 0, @vDateID OUTPUT;
		SET @iStartDateID = convert(integer,@vDateID);

		EXEC dbo.spASRIntGetSetting 'BradfordFactor', 'End Date', 0, 0, @vDateID OUTPUT;
		SET @iEndDateID = convert(integer,@vDateID);
	END

	/* Start Date prompted value. */		
	IF @iStartDateID > 0
	BEGIN
		EXEC sp_ASRIntGetFilterPromptedValues @iStartDateID, @sComponents OUTPUT
		SET @iStartDateComponentID = @sComponents

		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,StartEndType)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,'start'
			FROM ASRSysExprComponents
			WHERE componentID = @iStartDateComponentID)
	END

	/* End Date prompted value. */
	IF @iEndDateID > 0
	BEGIN
		EXEC sp_ASRIntGetFilterPromptedValues @iEndDateID, @sComponents OUTPUT
		SET @iEndDateComponentID = @sComponents

		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,StartEndType)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,'end'
			FROM ASRSysExprComponents
			WHERE componentID = @iEndDateComponentID)
	END


	SELECT DISTINCT * 
	FROM @promptedValues
	ORDER BY startEndType DESC;
	
END
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRIntGetUtilityPromptedValues]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[spASRIntGetUtilityPromptedValues]
GO

CREATE PROCEDURE [dbo].[spASRIntGetUtilityPromptedValues] (
	@piUtilType 	integer,
	@piUtilID 		integer,
	@piRecordID 	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the prompted values for the given utililty. */
	DECLARE	@iBaseFilter		integer,
			@iParent1Filter		integer,
			@iParent2Filter		integer,
			@iChildFilter		integer,
			@iEventFilter		integer,
			@iStartDateCalc		integer,
			@iEndDateCalc		integer,
			@iDescCalc			integer,
			@iLoop				integer,
			@iFilterID			integer,
			@iCalcID			integer,
			@iComponentID		integer,
			@sComponents		varchar(MAX),
			@sAllComponents		varchar(MAX),
			@iIndex				integer;
	
	SET @sAllComponents = '';

	IF @piRecordID IS null SET @piRecordID = 0;

	/* Create a temp table to hold the propmted value details. */
	DECLARE @promptedValues TABLE(
		componentID			integer,
		promptDescription	varchar(255),
		valueType			integer,
		promptMask			varchar(255),
		promptSize			integer,
		promptDecimals		integer,
		valueCharacter		varchar(255),
		valueNumeric		float,
		valueLogic			bit,
		valueDate			datetime,
		promptDateType		integer, 
		fieldColumnID		integer);


	IF @piUtilType = 1 OR @piUtilType = 35
	BEGIN
		/* Cross Tabs. */
		SELECT @iFilterID = filterid
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID

		IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sAllComponents OUTPUT
		END		
	END


	IF @piUtilType = 15 OR @piUtilType = 16
	BEGIN
		/* Standard report (Absence Calendar or Bradford Factor) */
		IF (NOT @piUtilID IS NULL) AND (@piUtilID > 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @piUtilID, @sAllComponents OUTPUT
		END		
	END



	IF @piUtilType = 2 OR @piUtilType = 9
	BEGIN

		IF @piUtilType = 2
		BEGIN
			/* Custom Reports. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iBaseFilter = filter, 
				@iParent1Filter = parent1Filter, 
				@iParent2Filter = parent2Filter /*, 
				@iChildFilter = childFilter*/
			FROM [dbo].[ASRSysCustomReportsName]
			WHERE ID = @piUtilID

			IF @piRecordID <> 0
			BEGIN
				SET @iBaseFilter = 0
			END
			
			/* Get the prompted values used in the Base and Parent table filters. */
			SET @iLoop = 0
			WHILE @iLoop < 3
			BEGIN
				IF @iLoop = 0 SET @iFilterID = @iBaseFilter
				IF @iLoop = 1 SET @iFilterID = @iParent1Filter
				IF @iLoop = 2 SET @iFilterID = @iParent2Filter
				--IF @iLoop = 3 SET @iFilterID = @iChildFilter

				IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
				BEGIN
					EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sComponents OUTPUT

					IF LEN(@sComponents) > 0
					BEGIN
						SET @sAllComponents = @sAllComponents + 
							CASE
								WHEN LEN(@sAllComponents) > 0 THEN ','
								ELSE ''
							END + 
							@sComponents
					END
				END

				SET @iLoop = @iLoop + 1
			END		

			/* Get the promted values used in the Child table filters. */
			DECLARE childs_cursor CURSOR LOCAL FAST_FORWARD FOR
				
			SELECT childFilter
			FROM [dbo].[ASRSysCustomReportsChildDetails]
			WHERE CustomReportID = @piUtilID

			OPEN childs_cursor
			FETCH NEXT FROM childs_cursor INTO @iChildFilter
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iChildFilter, @sComponents OUTPUT
		
				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
				FETCH NEXT FROM childs_cursor INTO @iChildFilter
			END
			CLOSE childs_cursor
			DEALLOCATE childs_cursor


			/* Get the prompted values used in the runtime calcs in the report. */
			DECLARE calcs_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT colExprID 
				FROM [dbo].[ASRSysCustomReportsDetails]
				WHERE customReportID = @piUtilID
					AND type = 'E'

		END

		IF @piUtilType = 9
		BEGIN
			/* Mail Merge. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iFilterID = filterID
			FROM [dbo].[ASRSysMailMergeName]
			WHERE MailMergeID = @piUtilID

			IF @piRecordID <> 0
			BEGIN
				SET @iFilterID = 0
			END

			IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sAllComponents OUTPUT
			END		

			/* Get the prompted values used in the runtime calcs in the report. */
			DECLARE calcs_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ColumnID 
				FROM [dbo].[ASRSysMailMergeColumns]
				WHERE MailMergeID = @piUtilID
					AND type = 'E'
		END


		OPEN calcs_cursor
		FETCH NEXT FROM calcs_cursor INTO @iCalcID
		WHILE (@@fetch_status = 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT
		
			IF LEN(@sComponents) > 0
			BEGIN
				SET @sAllComponents = @sAllComponents + 
					CASE
						WHEN LEN(@sAllComponents) > 0 THEN ','
						ELSE ''
					END + 
					@sComponents
			END
			FETCH NEXT FROM calcs_cursor INTO @iCalcID
		END
		CLOSE calcs_cursor
		DEALLOCATE calcs_cursor
	END

	IF @piUtilType = 17
		BEGIN
			/* Calendar Reports. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iBaseFilter = filter, 
						 @iStartDateCalc = StartDateExpr, 
			 			 @iEndDateCalc = EndDateExpr,
						 @iDescCalc = DescriptionExpr
			FROM ASRSysCalendarReports
			WHERE ID = @piUtilID
				
			IF @piRecordID = 0
			BEGIN
				/* Get the prompted values used in the Base table filter. */
				SET @iFilterID = @iBaseFilter

				IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
				BEGIN
					EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sComponents OUTPUT

					IF LEN(@sComponents) > 0
					BEGIN
						SET @sAllComponents = @sAllComponents + 
							CASE
								WHEN LEN(@sAllComponents) > 0 THEN ','
								ELSE ''
							END + 
							@sComponents
					END
				END
			END
			
			/* Get the prompted values used in the Report Start Date Calculation. */
			SET @iCalcID = @iStartDateCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the prompted values used in the Report End Date Calculation. */
			SET @iCalcID = @iEndDateCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the prompted values used in the Report Description Calculation. */
			SET @iCalcID = @iDescCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC sp_ASRIntGetFilterPromptedValues @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the promted values used in the Event table filters. */
			DECLARE events_cursor CURSOR LOCAL FAST_FORWARD FOR
				
			SELECT FilterID
			FROM ASRSysCalendarReportEvents
			WHERE CalendarReportID = @piUtilID

			OPEN events_cursor
			FETCH NEXT FROM events_cursor INTO @iEventFilter
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC sp_ASRIntGetFilterPromptedValues @iEventFilter, @sComponents OUTPUT
		
				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
				FETCH NEXT FROM events_cursor INTO @iEventFilter
			END
			CLOSE events_cursor
			DEALLOCATE events_cursor
			
	END
		
		
	/* We now have a string of all of the prompted value components that are used in the filters and calculations. */
	WHILE LEN(@sAllComponents) > 0 
	BEGIN
		/* Get the individual component IDs from the comma-delimited string. */
		SET @iIndex = CHARINDEX(',', @sAllComponents)

		IF @iIndex > 0 
		BEGIN
			SET @iComponentID = convert(integer, SUBSTRING(@sAllComponents, 1, @iIndex - 1))
			SET @sAllComponents = SUBSTRING(@sAllComponents, @iIndex + 1, LEN(@sAllComponents) - @iIndex)
		END
		ELSE
		BEGIN
			/* No comma, must be dealing with the last component in the list. */
			SET @iComponentID = convert(integer, @sAllComponents)
			SET @sAllComponents = ''
		END

		/* Get the parameters of the prompted values. */
		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID
			FROM ASRSysExprComponents
			WHERE componentID = @iComponentID)
	END

	SELECT DISTINCT * 
	FROM @promptedValues
END
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRIntGetCustomReportDetails]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[spASRIntGetCustomReportDetails]
GO

CREATE PROCEDURE [dbo].[spASRIntGetCustomReportDetails] (@piCustomReportID integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT d.*, ISNULL(c.Use1000separator,0) AS Use1000separator
			, ISNULL(c.columnname,'') AS [columnname]
			, ISNULL(t.tableid,0) AS [tableid]
			, ISNULL(t.tablename,'') AS [tablename]
			, CASE c.datatype WHEN 11 THEN 1 ELSE 0 END AS [IsDateColumn]
			, CASE c.datatype WHEN -7 THEN 1 ELSE 0 END AS [IsBooleanColumn]
			, ISNULL(c.datatype,0) AS [DataType]
		FROM ASRSysCustomReportsDetails d
		LEFT JOIN ASRSysColumns c ON c.columnid = d.ColExprID And d.Type = 'C'
		LEFT JOIN ASRSysTables t ON c.tableid = t.tableid
	WHERE CustomReportID = @piCustomReportID ORDER BY [Sequence];

END
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRIntGetMailMergeDS]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[spASRIntGetMailMergeDS]
GO

CREATE PROCEDURE [dbo].[spASRIntGetMailMergeDS](@id AS integer)
AS
BEGIN

	SET NOCOUNT ON;

	-- Definition
	SELECT m.*, t.TableName, t.RecordDescExprID
		FROM ASRSysMailMergeName m
		JOIN ASRSYSTables t ON (t.TableID = m.TableID) WHERE MailMergeID = @id;

	-- Columns
	SELECT 0 AS [IsExpression],  c.ColumnID AS ColExpId
		, t.TableID AS [tableid], t.Tablename AS [TableName], c.ColumnName AS [Name]
		, c.DataType AS [Type], m.Size, m.Decimals, c.Use1000Separator
	FROM ASRSysMailMergeColumns m
		INNER JOIN ASRSysColumns c ON (c.ColumnID = m.ColumnID) 
		INNER JOIN ASRSysTables t ON (t.TableID = c.TableID) WHERE m.Type = 'C' AND m.MailMergeID = @id
	UNION    
	SELECT 1 AS [IsExpression], e.ExprID AS [ColExpId],  0 AS [TableID]
		, 'Calculation_' AS [Table], e.Name AS [Name]
		, e.ReturnType as [Type], m.Size, m.Decimals, 0 AS [Use1000Separator]
	FROM ASRSysMailMergeColumns m
		LEFT OUTER JOIN ASRSysExpressions e ON (e.ExprID = m.ColumnID)
		WHERE m.Type = 'E' AND m.MailMergeID = @id
	ORDER BY [TableName], [Name];

	-- Sort Order
	SELECT t.TableID, t.TableName, c.ColumnID AS ColExpId, c.ColumnName AS [Name], mc.SortOrder 
		FROM ASRSysMailMergeColumns mc
		INNER JOIN ASRSysColumns c ON mc.ColumnID = c.ColumnID
		INNER JOIN ASRSysTables t ON c.TableID = t.TableID
		WHERE mc.MailMergeID = @id AND SortOrderSequence > 0
	ORDER BY SortOrderSequence;

END
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spASRIntGetLookupFindRecords2]') AND type in (N'P', N'PC'))
	DROP PROCEDURE [dbo].[spASRIntGetLookupFindRecords2]
GO

CREATE PROCEDURE [dbo].[spASRIntGetLookupFindRecords2] (
	@piTableID 			integer, 
	@piViewID 			integer, 
	@piOrderID 			integer,
	@piLookupColumnID 	integer,
	@piRecordsRequired	integer,
	@pfFirstPage		bit		OUTPUT,
	@pfLastPage			bit		OUTPUT,
	@psLocateValue		varchar(MAX),
	@piColumnType		integer		OUTPUT,
	@piColumnSize		integer		OUTPUT,
	@piColumnDecimals	integer		OUTPUT,
	@psAction			varchar(MAX),
	@piTotalRecCount	integer		OUTPUT,
	@piFirstRecPos		integer		OUTPUT,
	@piCurrentRecCount	integer,
	@psFilterValue		varchar(MAX),
	@piCallingColumnID	integer,
	@piLookupColumnGridNumber	integer		OUTPUT,
	@pfOverrideFilter	bit
)
AS
BEGIN
	/* Return a recordset of the link find records for the current user, given the table/view and order IDs.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@pfError = 1 if errors occured in getting the find records. Else 0.
	*/
	
	SET NOCOUNT ON;
	
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@iColumnID 			integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 			varchar(10),
		@iDataType 			integer,
		@fSelectGranted 	bit,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sExecString		nvarchar(MAX),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString			varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iTemp				integer,
		@sRemainingSQL		varchar(MAX),
		@iLastCharIndex		integer,
		@iCharIndex 		integer,
		@sDESCstring		varchar(5),
		@sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@fFirstColumnAsc	bit,
		@sFirstColCode		varchar(MAX),
		@sLocateCode		varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@iCount				integer,
		@iGetCount			integer,
		@iColSize			integer,
		@iColDecs			integer,
		@fLookupColumnDone	bit,
		@sLookupColumnName	sysname,
		@iLookupTableID		integer,
		@iLookupColumnType	integer,
		@iLookupColumnSize	integer,
		@iLookupColumnDecimals integer,
		@iCount2			integer,
		@sFilterSQL			nvarchar(MAX),
		@sColumnTemp		sysname,
		@iLookupFilterColumnID	integer,
		@iLookupFilterOperator	integer,
		@iLookupFilterColumnDataType	integer,
		@sActualUserName	sysname;

	/* Initialise variables. */
	SET @sRealSource = ''
	SET @sSelectSQL = ''
	SET @sOrderSQL = ''
	SET @fSelectDenied = 0
	SET @sExecString = ''
	SET @sDESCstring = ' DESC'
	SET @fFirstColumnAsc = 1
	SET @sFirstColCode = ''
	SET @sReverseOrderSQL = ''
	SET @fLookupColumnDone = 0
	SET @piLookupColumnGridNumber = 0
	SET @sFilterSQL = ''

	/* Clean the input string parameters. */
	IF len(@psFilterValue) > 0 SET @psFilterValue = replace(@psFilterValue, '''', '''''')
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''')
	
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000
	SET @psAction = UPPER(@psAction)
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE')
	BEGIN
		SET @psAction = 'MOVEFIRST'
	END

	/* Get the column name. */
	SELECT @sLookupColumnName = ASRSysColumns.columnName,
		@iLookupTableID = ASRSysColumns.tableID,
		@iLookupColumnType = ASRSysColumns.dataType,
		@iLookupColumnSize = ASRSysColumns.size,
		@iLookupColumnDecimals = ASRSysColumns.decimals
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piLookupColumnID

	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @piViewID
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piTableID
			AND role = @sUserGroupName
			
		IF @iChildViewID IS null SET @iChildViewID = 0
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_')
			SET @sRealSource = left(@sRealSource, 255)
		END
	END

	IF len(@sRealSource) = 0
	BEGIN
		RETURN
	END

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(
		tableViewName	sysname,
		tableID			integer);

	/* Create a temporary table of the 'select' column permissions for all tables/views used in the order. */
	DECLARE @columnPermissions TABLE(
		tableID			integer,
		tableViewName	sysname,
		columnName		sysname,
		selectGranted	bit);

	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	WHERE ASRSysOrderItems.orderID = @piOrderID

	OPEN tablesCursor
	FETCH NEXT FROM tablesCursor INTO @iTempTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @columnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @columnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				CASE protectType
				        	WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	/* Create the lookup filter string. NB. We already know that the user has SELECT permission on this from the spASRIntGetLookupViews stored procedure.*/
	SELECT @iLookupFilterColumnID = ASRSysColumns.LookupFilterColumnID,
		@iLookupFilterOperator = ASRSysColumns.LookupFilterOperator
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piCallingColumnID

	IF (@iLookupFilterColumnID > 0) and (@pfOverrideFilter = 0)
	BEGIN
		SELECT @sColumnTemp = ASRSysColumns.columnName,
			@iLookupFilterColumnDataType = ASRSysColumns.dataType
		FROM ASRSysColumns
		WHERE ASRSysColumns.columnId = @iLookupFilterColumnID

		SELECT @iCount = COUNT(*)
		FROM @columnPermissions
		WHERE columnName = @sColumnTemp
			AND selectGranted = 1

		IF @iCount > 0 AND @psFilterValue <> ''
		BEGIN
			IF @iLookupFilterColumnDataType = -7 /* Boolean */
			BEGIN
				SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = '
					+ CASE
						WHEN UPPER(@psFilterValue) = 'TRUE' THEN '1'
						WHEN UPPER(@psFilterValue) = 'FALSE' THEN '0'
						ELSE @psFilterValue
					END
					+ ') '
			END
			ELSE
			BEGIN
				IF (@iLookupFilterColumnDataType = 2) OR (@iLookupFilterColumnDataType = 4) /* Numeric, Integer */
				BEGIN
					IF @iLookupFilterOperator = 1 /* Equals */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) = 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 2 /* NOT Equal To */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) = 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
						END
					END

					IF @iLookupFilterOperator = 3 /* Is At Most */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <= ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) >= 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 4 /* Is At Least */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' >= ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) <= 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 5 /* Is More Than */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' > ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) < 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 6 /* Is Less Than */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' < ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) > 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END
				END
				ELSE
				BEGIN
					IF (@iLookupFilterColumnDataType = 11) /* Date */
					BEGIN
						IF @iLookupFilterOperator = 7 /* On */
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
						END

						IF @iLookupFilterOperator = 8 /* NOT On */
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
							END
						END

						IF @iLookupFilterOperator = 12 /* On OR Before*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <= ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
						END

						IF @iLookupFilterOperator = 11 /* On OR After*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' >= ''' + @psFilterValue + ''') ' 
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null)' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END

						IF @iLookupFilterOperator = 9 /* After*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' > ''' + @psFilterValue + ''') ' 
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END

						IF @iLookupFilterOperator = 10 /* Before*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' < ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null)' +
									' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END
					END
					ELSE
					BEGIN
						IF (@iLookupFilterColumnDataType = 12) OR (@iLookupFilterColumnDataType = -3) OR (@iLookupFilterColumnDataType = -1) /* varchar, working patter, photo*/
						BEGIN
							IF @iLookupFilterOperator = 14 /* Is */
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = '''') ' +
										' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
								END
							END

							IF @iLookupFilterOperator = 16 /* Is NOT*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> '''') ' +
										' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') '
								END
							END

							IF @iLookupFilterOperator = 13 /* Contains*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) ' +
										' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' LIKE ''%' + @psFilterValue + '%'') '
								END
							END

							IF @iLookupFilterOperator = 15 /* Does NOT Contain*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) ' +
										' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' NOT LIKE ''%' + @psFilterValue + '%'') '
								END
							END
						END
					END
				END
			END
		END
	END

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 

	SELECT ASRSysColumns.tableID,
		ASRSysOrderItems.columnID, 
		ASRSysColumns.columnName,
	    	ASRSysTables.tableName,
		ASRSysOrderItems.ascending,
		ASRSysOrderItems.type,
		ASRSysColumns.dataType,
		ASRSysColumns.size,
		ASRSysColumns.decimals
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
	ORDER BY ASRSysOrderItems.sequence

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		RETURN
	END
	SET @iCount2 = 0

	WHILE (@@fetch_status = 0) OR (@fLookupColumnDone = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF (@@fetch_status <> 0)
		BEGIN
			SET @iColumnTableId = @iLookupTableID
			SET @iColumnId = @piLookupColumnID
			SET @sColumnName = @sLookupColumnName
			SET @sColumnTableName = @sTableName
			SET @fAscending = 1
			SET @sType = 'F'
			SET @iDataType = @iLookupColumnType
			SET @iColSize = @iLookupColumnSize
			SET @iColDecs = @iLookupColumnDecimals
		END

		IF (@iColumnId  = @piLookupColumnID ) AND (@sType = 'F')
		BEGIN
			SET @fLookupColumnDone = 1
			SET @piLookupColumnGridNumber = @iCount2
		END

		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0

			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN

					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName
					SET @iCount2 = @iCount2 + 1
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType
						SET @piColumnSize = @iColSize
						SET @piColumnDecimals = @iColDecs
						SET @fFirstColumnAsc = @fAscending
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName
					END

					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName +
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END				
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1
			END	
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @columnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName
					SET @iCount2 = @iCount2 + 1
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType
						SET @piColumnSize = @iColSize
						SET @piColumnDecimals = @iColDecs
						SET @fFirstColumnAsc = @fAscending
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName
					END

					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName + 
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END				
				END

				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)

				FROM @joinParents
				WHERE tableViewName = @sColumnTableName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID)
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = ''

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = 'CASE'

					SET @sSubString = @sSubString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName 
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewname = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +
						' ELSE NULL END'

					IF @sType = 'F'
					BEGIN
						/* Find column. */
						SET @sSubString = @sSubString +
							' AS [' + @sColumnName + ']'

						SET @sSelectSQL = @sSelectSQL + 
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END
						SET @iCount2 = @iCount2 + 1
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType
							SET @piColumnSize = @iColSize
							SET @piColumnDecimals = @iColDecs
							SET @fFirstColumnAsc = @fAscending
							SET @sFirstColCode = @sSubString
						END

						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END + 
							CASE 
								WHEN @fAscending = 0 THEN ' DESC' 
								ELSE '' 
							END				
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1
				END	
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN ',' ELSE '' END + 
		@sRealSource + '.ID'

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sRemainingSQL = @sOrderSQL

		SET @iLastCharIndex = 0
		SET @iCharIndex = CHARINDEX(',', @sOrderSQL)
		WHILE @iCharIndex > 0 
		BEGIN
 			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', '
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', '
			END

			SET @iLastCharIndex = @iCharIndex
			SET @iCharIndex = CHARINDEX(',', @sOrderSQL, @iLastCharIndex + 1)
	
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex)
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring

	END

	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource

	IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL

	SET @sTempParamDefinition = N'@recordCount integer OUTPUT'
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT
	SET @piTotalRecCount = @iCount

	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = @sSelectSQL + ',' + @sRealSource + '.ID'
		SET @sExecString = 'SELECT ' 

		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
		BEGIN
			SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' '
		END
		SET @sExecString = @sExecString + @sSelectSQL + 
			' FROM ' + @sRealSource

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sExecString = @sExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID'

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		IF (@psAction = 'MOVELAST')
		BEGIN
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END

		IF @psAction = 'MOVENEXT' 
		BEGIN
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1)
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired
			END
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource

			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired  - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END
		IF @psAction = 'MOVEPREVIOUS'
		BEGIN
			IF @piFirstRecPos <= @piRecordsRequired
			BEGIN
				SET @iGetCount = @piFirstRecPos - 1
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired
			END
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource

			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END

		IF len(@sFilterSQL) > 0 SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL

		IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL + ')'
		END

		IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL + ')'
		END

		IF (@psAction = 'LOCATE')
		BEGIN
			IF len(@sFilterSQL) > 0 
			BEGIN
				SET @sLocateCode = ' AND (' + @sFirstColCode 
			END
			ELSE
			BEGIN
				SET @sLocateCode = ' WHERE (' + @sFirstColCode 
			END

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''

					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL'
				END

			END

			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sFirstColCode + ' IS NULL'
					END
					ELSE
					BEGIN

						SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''
					END
				END
				ELSE
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NULL'
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
			END

			IF @piColumnType = -7 /* Logic column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END
				END
			END

			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue

					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL'
				END

			END

			SET @sLocateCode = @sLocateCode + ')'
			SET @sExecString = @sExecString + @sLocateCode
		END

		/* Add the ORDER BY code to the find record selection string if required. */
		SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL
	END

	/* Set the IsFirstPage, IsLastPage flags, and the page number. */
	IF @psAction = 'MOVEFIRST'
	BEGIN
		SET @piFirstRecPos = 1
		SET @pfFirstPage = 1
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVENEXT'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount
		SET @pfFirstPage = 0
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos - @iGetCount
		IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVELAST'
	BEGIN
		SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1
		IF @piFirstRecPos < 1 SET @piFirstRecPos = 1
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 1
	END

	IF @psAction = 'LOCATE'
	BEGIN
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID'

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL

		SET @sTempExecString = @sTempExecString + @sLocateCode

		SET @sTempParamDefinition = N'@recordCount integer OUTPUT'
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iTemp OUTPUT

		IF @iTemp <=0 
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount + 1
		END
		ELSE
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount - @iTemp + 1
		END

		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END
	END

	/* Return a recordset of the required columns in the required order from the given table/view. */
	IF (len(@sExecString) > 0)
	BEGIN
		EXEC sp_executeSQL @sExecString;
	END
END
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

		
DECLARE @sVersion varchar(10) = '8.1.11'

EXEC spsys_setsystemsetting 'database', 'version', '8.1';
EXEC spsys_setsystemsetting 'intranet', 'version', @sVersion;
EXEC spsys_setsystemsetting 'ssintranet', 'version', @sVersion;
