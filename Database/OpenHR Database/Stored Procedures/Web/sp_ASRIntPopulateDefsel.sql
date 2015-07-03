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
				'lower(' + @strTableName + '.Username) as [Username], ';
				
			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					'CASE WHEN Username = SYSTEM_USER THEN ''rw'' ELSE LOWER(' + @sAccessTableName + '.Access) END AS [Access], ';

			END
			ELSE
			BEGIN
				SET @strSQL = @strSQL +
					'''rw'' as [Access], ';
			END
								
			SET @strSQL = @strSQL +
				@strTableName + '.' + @strIDName + '  AS [ID] 
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
			SET @strSQL = 'SELECT Name, replace(Description, char(9), '''') AS [description], lower(Username) AS [Username], ';
			IF (@fSysSecMgr = 0)  
			BEGIN
				SET @strSQL = @strSQL +
					'CASE WHEN Username = SYSTEM_USER THEN ''rw'' ELSE LOWER([Access]) END AS [Access], ';
			END
			ELSE
			BEGIN
				SET @strSQL = @strSQL +
					'''rw'' AS [Access], ';
			END
			SET @strSQL = @strSQL +
				@strIDName + '  AS [ID] FROM ' + @strTableName;

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
	
	-- Return the resultset.
	EXECUTE sp_executeSQL @strSQL;
	
END