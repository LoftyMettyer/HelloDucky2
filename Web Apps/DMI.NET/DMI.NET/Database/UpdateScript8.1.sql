
---- Drop redundant functions (or renamed)
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRIntGetMailMergeDefinition]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASRIntGetCrossTabDefinition];
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
	DROP PROCEDURE [dbo].[spASRIntGetCrossTabDefinition];
GO
CREATE PROCEDURE [dbo].[spASRIntGetMailMergeDefinition] (	
			@piReportID 			integer, 	
			@psCurrentUser			varchar(255),		
			@psAction				varchar(255),
			@psErrorMsg				varchar(MAX)	OUTPUT,		
			@psReportName			varchar(255)	OUTPUT,		
			@psReportOwner			varchar(255)	OUTPUT,		
			@psReportDesc			varchar(255)	OUTPUT,		
			@piBaseTableID			integer			OUTPUT,		
			@piSelection			integer			OUTPUT,		
			@piPicklistID			integer			OUTPUT,		
			@psPicklistName			varchar(255)	OUTPUT,		
			@pfPicklistHidden		bit				OUTPUT,		
			@piFilterID				integer			OUTPUT,		
			@psFilterName			varchar(255)	OUTPUT,		
			@pfFilterHidden			bit				OUTPUT,		
			@piOutputFormat				integer			OUTPUT,		
			@pfOutputSave				bit				OUTPUT,		
			@psOutputFileName			varchar(MAX)	OUTPUT,		
			@piEmailAddrID 			integer			OUTPUT,		
			@psEmailSubject			varchar(255)	OUTPUT,		
			@psTemplateFileName		varchar(MAX)	OUTPUT,		
			@pfOutputScreen				bit				OUTPUT,		
			@pfEmailAsAttachment	bit				OUTPUT,		
			@psEmailAttachmentName	varchar(MAX)	OUTPUT,		
			@pfSuppressBlanks		bit				OUTPUT,		
			@pfPauseBeforeMerge		bit				OUTPUT,		
			@pfOutputPrinter			bit				OUTPUT,		
			@psOutputPrinterName	varchar(255)	OUTPUT,		
			@piDocumentMapID			integer		OUTPUT,		
			@pfManualDocManHeader		bit		OUTPUT,		
		 	@piTimestamp			integer			OUTPUT,		
			@psWarningMsg			varchar(MAX)	OUTPUT		
		)		
		AS		
		BEGIN		
			SET NOCOUNT ON;		
			DECLARE	@iCount		integer,		
					@sTempHidden	varchar(MAX),		
					@sAccess 		varchar(MAX),		
					@fSysSecMgr		bit;		
			SET @psErrorMsg = '';		
			SET @psPicklistName = '';		
			SET @pfPicklistHidden = 0;		
			SET @psFilterName = '';		
			SET @pfFilterHidden = 0;		
			SET @psWarningMsg = '';		
			exec [dbo].[spASRIntSysSecMgr] @fSysSecMgr OUTPUT;		
			/* Check the mail merge exists. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeName]		
			WHERE MailMergeID = @piReportID;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been deleted by another user.';		
				RETURN;		
			END		
			SELECT @psReportName = name,		
				@psReportDesc	 = description,		
				@psReportOwner = userName,		
				@piBaseTableID = tableID,		
				@piSelection = selection,		
				@piPicklistID = picklistID,		
				@piFilterID = filterID,		
				@piOutputFormat = outputformat,		
				@pfOutputSave = outputsave,		
				@psOutputFileName = outputfilename,		
				@piEmailAddrID = emailAddrID,		
				@psEmailSubject = emailSubject,		
				@psTemplateFileName = templateFileName,		
				@pfOutputScreen = outputscreen,		
				@pfEmailAsAttachment = emailasattachment,		
				@psEmailAttachmentName = isnull(emailattachmentname,''),		
				@pfSuppressBlanks = suppressblanks,		
				@pfPauseBeforeMerge = pausebeforemerge,		
				@pfOutputPrinter = outputprinter,		
				@psOutputPrinterName = outputprintername,		
				@piDocumentMapID = documentmapid,		
				@pfManualDocManHeader = manualdocmanheader,				
				@piTimestamp = convert(integer, timestamp)		
			FROM [dbo].[ASRSysMailMergeName]		
			WHERE MailMergeID = @piReportID;		
			/* Check the current user can view the report. */		
			exec [dbo].[spASRIntCurrentUserAccess]		
				9, 		
				@piReportID,		
				@sAccess OUTPUT;		
			IF (@sAccess = 'HD') AND (@psReportOwner <> @psCurrentUser) 		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been made hidden by another user.';		
				RETURN;		
			END		
			IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@sAccess = 'RO') AND (@psReportOwner <> @psCurrentUser) 		
			BEGIN		
				SET @psErrorMsg = 'mail merge has been made read only by another user.';		
				RETURN;		
			END		
			/* Check the report has details. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeColumns]		
			WHERE MailMergeID = @piReportID;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge contains no details.';		
				RETURN;		
			END		
			/* Check the report has sort order details. */		
			SELECT @iCount = COUNT(*)		
			FROM [dbo].[ASRSysMailMergeColumns]		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.sortOrderSequence > 0;		
			IF @iCount = 0		
			BEGIN		
				SET @psErrorMsg = 'mail merge contains no sort order details.';		
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
				FROM [dbo].[ASRSysPicklistName]		
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
				FROM [dbo].[ASRSysExpressions]		
				WHERE exprID = @piFilterID;		
				IF UPPER(@sTempHidden) = 'HD'		
				BEGIN		
					SET @pfFilterHidden = 1;		
				END		
			END

			-- Columns
			SELECT ASRSysMailMergeColumns.[type],
				ASRSysColumns.tableID,
				ASRSysMailMergeColumns.columnID,
				ASRSysColumns.columnName AS [name], 
				ASRSysTables.tableName + '.' + ASRSysColumns.columnName AS [heading],
				ASRSysColumns.DataType,
				ASRSysMailMergeColumns.size,
				ASRSysMailMergeColumns.decimals,
				CASE WHEN ASRSysColumns.DataType = 2 or ASRSysColumns.DataType = 4 THEN '1' ELSE '0' END AS [isnumeric],		
				ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
			FROM ASRSysMailMergeColumns		
			INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnID = ASRSysColumns.columnId		
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.type = 'C'		

			-- Expressions
			SELECT CASE WHEN ASRSysExpressions.access = 'HD' THEN 1 ELSE 0 END AS [ishidden],		
				ASRSysMailMergeColumns.[type],
				ASRSysExpressions.tableID,
				ASRSysMailMergeColumns.columnID,
				ASRSysExpressions.name AS [name],
				convert(varchar(MAX), '<Calc> ' + replace(ASRSysExpressions.name, '_', ' ')) AS [heading],
				ASRSysMailMergeColumns.size,
				ASRSysMailMergeColumns.decimals,
				ASRSysMailMergeColumns.SortOrderSequence AS [sequence]
			FROM ASRSysMailMergeColumns		
			INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnID = ASRSysExpressions.exprID		
			WHERE ASRSysMailMergeColumns.MailMergeID = @piReportID		
				AND ASRSysMailMergeColumns.type <> 'C'		
				AND ((ASRSysExpressions.username = @psReportOwner)	OR (ASRSysExpressions.access <> 'HD'))		

			-- Orders
			SELECT ASRSysMailMergeColumns.columnID,
				convert(varchar(MAX), ASRSysTables.tableName + '.' + ASRSysColumns.columnName) AS [columnname],
				ASRSysMailMergeColumns.sortOrder,
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
