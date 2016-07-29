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

	IF @piUtilityType = 38
	BEGIN
		/* Talent Report */
		SET @sAccessTable = 'ASRSysTalentReportAccess';
		SET @sKey = 'dfltaccess TalentReports';
	END

	IF @piUtilityType = 39
	BEGIN
		/* Organisation Report */
		SET @sAccessTable = 'ASRSysOrganisationReportAccess';
		SET @sKey = 'dfltaccess OrganisationReport';
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

