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

	IF @piUtilityType = 38 /* Talent Report*/
	BEGIN
		SET @sTableName = 'ASRSysTalentReports';
		SET @sAccessTableName = 'ASRSysTalentReportAccess';
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