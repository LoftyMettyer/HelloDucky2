CREATE PROCEDURE [dbo].[spASRIntGetSummaryValues] (
	@piHistoryTableID	integer,
	@piParentTableID 	integer,
	@piParentRecordID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(255),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(255),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sRootTable 		varchar(255),
		@sSelectString 		varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	varchar(255),
		@sTemp				varchar(MAX),
		@sSelectSQL			nvarchar(MAX),
		@sActualUserName	sysname,
		@strTempSepText		varchar(500);

	SET @sSelectSQL = '';
		
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

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
		AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';

	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID;

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(tableViewName	sysname);

	/* Create a temporary table of the 'read' column permissions for all tables/views used. */
	DECLARE @columnPermissions TABLE (tableViewName	sysname,
				columnName	sysname,
				granted		bit);

	-- Cached view of SysProtects
	DECLARE @SysProtects TABLE([ID] int, [ProtectType] tinyint, [Columns] varbinary(8000));

	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO @ColumnPermissions
		SELECT 
			@sParentTableName,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns 
		WHERE ASRSysColumns.tableID = @piParentTableID;
	END
	ELSE
	BEGIN
		IF @iParentTableType <> 2 /* ie. top-level or lookup */
		BEGIN

			-- Get list of views/table columns that are summary fields
			DECLARE @SummaryColumns TABLE ([ID] int, [TableName] sysname, [ColumnName] sysname, [ColID] int)
			INSERT @SummaryColumns
				SELECT sysobjects.id, sysobjects.name,
					syscolumns.name, syscolumns.ColID
				FROM sysobjects
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.name IN (SELECT ASRSysTables.tableName
												FROM ASRSysTables
												WHERE ASRSysTables.tableID = @piParentTableID
											UNION SELECT ASRSysViews.viewName
												FROM ASRSysViews
												WHERE ASRSysViews.viewTableID = @piParentTableID)
					AND syscolumns.name IN (SELECT ac.ColumnName
												FROM ASRSysSummaryFields am
												INNER JOIN ASRSysColumns ac ON am.ParentColumnID = ac.ColumnID
												WHERE HistoryTableID = @piHistoryTableID);

			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM #SysProtects
				WHERE Action = 193;

			-- Generate security context on selected columns
			INSERT INTO @ColumnPermissions
				SELECT sm.TableName,
					sm.ColumnName,
					CASE p.protectType
						WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @SysProtects p
				INNER JOIN @SummaryColumns sm ON p.id = sm.id
				WHERE (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) = 0));

		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @piParentTableID
				AND role = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sParentRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sParentTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sParentRealSource = left(@sParentRealSource, 255);
			END

			INSERT INTO @SysProtects
				SELECT ID, ProtectType, Columns FROM #SysProtects
				WHERE Action = 193;

			INSERT INTO @ColumnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	END

	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnID, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence;

	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		/* Get the select permission on the column. */

		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM @ColumnPermissions
		WHERE tableViewName = @sParentTableName
			AND columnName = @sColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			SET @sTemp = ',';
			IF LEN(@sSelectSQL) > 0
				SET @sSelectSQL = @sSelectSQL + @sTemp;

			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sTemp = 'convert(varchar(10), ' + @sParentTableName + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END 
			ELSE
			BEGIN
				 /* Non-date */
				SET @sTemp = @sParentTableName + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sSelectSQL = @sSelectSQL + @sTemp;
			END

			/* Add the table to the array of tables/views to join if it has not already been added. */
			SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
				WHERE tableViewName = @sParentTableName;

			IF @iTempCount = 0
			BEGIN
				INSERT INTO @joinParents (tableViewName) VALUES(@sParentTableName);
			END
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = '';

			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @ColumnPermissions
			WHERE tableViewName <> @sParentTableName
				AND columnName = @sColumnName
				AND granted = 1;

			OPEN viewCursor;
			FETCH NEXT FROM viewCursor INTO @sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Column CAN be read from the view. */
				SET @fSelectGranted = 1;

				IF len(@sSelectString) > 0 SET @sSelectString = @sSelectString + ',';

				IF @iColumnDataType = 11 /* Date */
					SET @sSelectString = @sSelectString + 'convert(varchar(10),' + @sViewName + '.' + @sColumnName + ',101)';
				ELSE
					SET @sSelectString = @sSelectString + @sViewName + '.' + @sColumnName;


				/* Add the view to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewName = @sViewName;

				IF @iTempCount = 0
					INSERT INTO @joinParents (tableViewName) VALUES(@sViewName);

				FETCH NEXT FROM viewCursor INTO @sViewName;
			END
			CLOSE viewCursor;
			DEALLOCATE viewCursor;

			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = 'COALESCE(' + @sSelectString + ', NULL) AS [' + convert(varchar(100), @iColumnID) + ']';
				SET @sTemp = ',';

				IF LEN(@sSelectSQL) > 0
					SET @sSelectSQL = @sSelectSQL + @sTemp;

				SET @sTemp = @sSelectString;
				SET @sSelectSQL = @sSelectSQL + @sTemp;
				
			END
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = 'SELECT ' + @sSelectSQL ;

		SELECT @iTempCount = COUNT(tableViewName)
			FROM @joinParents;

		IF @iTempCount = 1 
		BEGIN
			SELECT TOP 1 @sRootTable = tableViewName
			FROM @joinParents;
		END
		ELSE
		BEGIN
			SET @sRootTable = @sParentTableName;
		END

		SET @sTemp = ' FROM ' + @sRootTable;
		SET @sSelectSQL = @sSelectSQL + @sTemp;

		/* Add the join code. */
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @joinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sTableViewName <> @sRootTable
			BEGIN
				SET @sTemp = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRootTable + '.ID' + ' = ' + @sTableViewName + '.ID';
				SET @sSelectSQL = @sSelectSQL + @sTemp
			END

			FETCH NEXT FROM joinCursor INTO @sTableViewName;
		END
		CLOSE joinCursor;
		DEALLOCATE joinCursor;

		SET @sTemp = ' WHERE ' + @sRootTable + '.id = ' + convert(varchar(255), @piParentRecordID);
		SET @sSelectSQL = @sSelectSQL + @sTemp;

	END

	-- Run the constructed SQL SELECT string.
	EXEC sp_executeSQL @sSelectSQL;

END