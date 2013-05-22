CREATE PROCEDURE sp_ASRIntGetLinkParentValues (
	@piChildScreenID 	integer,
	@piTableID 			integer,
	@piRecordID			integer	
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of values from the given record in the given table that in the given child screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iScreenTableID 	integer,
		@iTableType			integer,
		@sTableName			varchar(255),
		@sRealSource 		varchar(1000),
		@iChildViewID 		integer,
		@sViewName 			varchar(255),
		@sSelectSQL 		varchar(MAX),
		@sFromSQL 			varchar(MAX),
		@sSelectString 		varchar(MAX),
		@sExecString		nvarchar(MAX),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@sColumnTableName 	varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sTemp				varchar(MAX),
		@sTempSPName		sysname,
		@iLoop				integer,
		@sActualUserName	sysname;

	/* Initialise variables. */
	SET @sSelectSQL = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID

	/* Get the screen's table ID. */
	SELECT @iScreenTableID = tableID
	FROM ASRSysScreens
	WHERE screenID = @piChildScreenID

	/* Check if the current user is System or Security Manager. 
	If so we don't need to do so much work figuring out what permissions they have. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1
	END
	ELSE
	BEGIN	
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
			AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'
	END

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		SET @sRealSource = @sTableName
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

	/* Create a temporary table to hold the views that need to be joined. */
	DECLARE @joinViews TABLE(viewName sysname);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @columnPermissions TABLE(
		tableViewName	sysname,
		columnName		sysname,
		granted			bit);

	IF @fSysSecMgr = 1
	BEGIN
		INSERT INTO @columnPermissions
		SELECT 
			@sRealSource,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns
		INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
		WHERE ASRSysTables.tableID = @piTableID

	END
	ELSE
	BEGIN
		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			INSERT INTO @columnPermissions
			SELECT 
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
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @piTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @piTableID)
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			INSERT INTO @columnPermissions
			SELECT 
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
	END

	/* Populate the temporary table with info for all columns used in the screen controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysTables.tableName,
		ASRSysColumns.dataType
	FROM ASRSysControls
	LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
	LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID 
		AND ASRSysColumns.columnID = ASRSysControls.columnID
	WHERE screenID = @piChildScreenID
	AND ASRSysControls.columnID > 0
	AND ASRSysControls.tableID = @piTableID

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF @fSysSecMgr = 1 
		BEGIN
			SET @fSelectGranted = 1
		END
		ELSE
		BEGIN

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
		END

		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the table. */
			IF len(@sSelectSQL) > 0 SET @sSelectSQL = @sSelectSQL + ', '
			
			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sTemp = 'convert(varchar(10), ' + @sRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']'
				SET @sSelectSQL = @sSelectSQL + @sTemp
			END
			ELSE
			BEGIN
				 /* Non-date */
				SET @sTemp = @sRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']'
				SET @sSelectSQL = @sSelectSQL + @sTemp
			END			
		END
		ELSE	
		BEGIN
			IF @iTableType = 1 /* Top-level. */
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */
				SET @sSelectString = ''

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableViewName <> @sRealSource
					AND columnName = @sColumnName
					AND granted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					SET @fSelectGranted = 1 

					IF len(@sSelectString) = 0 SET @sSelectString = 'CASE'
	
					IF @iColumnDataType = 11 /* Date */
					BEGIN
						 /* Date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)'
					END
					ELSE
					BEGIN
						 /* Non-date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName 
					END

					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(viewName)
					FROM @joinViews
					WHERE viewName = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @joinViews (viewName) VALUES(@sViewName)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor

				IF len(@sSelectString) > 0
				BEGIN
					SET @sSelectString = @sSelectString +
						' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']'
					IF LEN(@sSelectSQL) > 0 SET @sSelectSQL = @sSelectSQL + ', '
					SET @sSelectSQL = @sSelectSQL + @sSelectString		
				END
			END

		END

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor

	/* Add the id column to the select string. */
	SELECT @iColumnID = columnID
		FROM ASRSysColumns
		WHERE columnName = 'ID_' + convert(varchar(255), @piTableID)
		AND tableID = @iScreenTableID

	SET @sTemp = 	CASE
			WHEN LEN(@sSelectSQL) > 0 THEN ', '
			ELSE ''
		END + 
		@sRealSource + '.ID AS [' + convert(varchar(100), @iColumnID) + ']'

	SET @sSelectSQL = @sSelectSQL + @sTemp


	/* Create the FROM code. */
	SET @sFromSQL = @sRealSource;
	SELECT @sFromSQL = @sFromSQL
		+ ' LEFT OUTER JOIN ' + ViewName + ' ON ' + @sRealSource
		+ '.ID = ' + ViewName + '.ID'
	FROM @joinViews;


	/* Return a recordset of the required columns in the required order from the given table/view. */
	IF LEN(@sSelectSQL) > 0 
	BEGIN

		/* Create the temp stored procedure name. */
		SET @sTempSPName = ''
		SET @iLoop = 1
		WHILE len(@sTempSPName) = 0
		BEGIN
			SET @sTemp = 'tmpsp_ASRIntGetLinkParentValues' + convert(varchar(100), @iLoop)

			SELECT @iTempCount = COUNT(*)
			FROM sysobjects
			WHERE name = @sTemp

			IF @iTempCount = 0
			BEGIN
				SET @sTempSPName = @sTemp
			END
			ELSE
			BEGIN
				SET @iLoop = @iLoop + 1
			END
		END

		SET @sTemp = convert(varchar(255), @piRecordID)
	
		SET @sExecString = 'CREATE PROCEDURE ' + @sTempSPName + ' AS' +
			' BEGIN' +
			' SELECT ' + @sSelectSQL + 
			' FROM ' + @sFromSQL + 
			' WHERE ' + @sRealSource + '.ID = ' + @sTemp +
			' END'

		-- Create the temporary stored procedure
		EXECUTE sp_executeSQL @sExecString;

		-- Execute the temporary stored procedure
		EXECUTE sp_executeSQL @sTempSPName;

		SET @sExecString = 'DROP PROCEDURE ' + @sTempSPName
		exec sp_executeSQL @sExecString;
	END
END

