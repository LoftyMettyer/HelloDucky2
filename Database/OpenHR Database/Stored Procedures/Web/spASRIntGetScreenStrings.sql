CREATE PROCEDURE [dbo].[spASRIntGetScreenStrings] (
	@piScreenID 	integer,
	@piViewID 		integer,
	@psSelectSQL	nvarchar(MAX)	OUTPUT,
	@psFromDef		varchar(MAX)	OUTPUT,
	@psOrderSQL		varchar(MAX)	OUTPUT,
	@piOrderID		integer			OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@iScreenTableID		integer,
		@iScreenTableType	integer,
		@sScreenTableName	varchar(255),
		@iScreenOrderID 	integer,
		@sRealSource 		varchar(255),
		@iChildViewID 		integer,
		@sJoinCode 			varchar(MAX),
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@sColumnTableName 	varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@fUpdateGranted 	bit,
		@sSelectString 		varchar(MAX),
		@iTempCount 		integer,
		@sViewName 			varchar(255),
		@fAscending 		bit,
		@sOrderString 		varchar(MAX),
		@sTableViewName 	varchar(255),
		@iJoinTableID 		integer,
		@sParentRealSource	varchar(255),
		@iParentChildViewID	integer,
		@iParentTableType	integer,
		@sParentTableName	sysname,
		@iColumnType		integer,
		@iLinkTableID		integer,
		@lngPermissionCount	integer,
		@iLinkChildViewID	integer,
		@sLinkRealSource	varchar(255),
		@sLinkTableName		varchar(255),
		@iLinkTableType		integer,
		@sNewBit			varchar(MAX),
		@iID				integer,
		@iCount				integer,
		@iUserType			integer,
		@sRoleName			sysname,
		@sActualUserName	sysname;
		
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;
		
	/* Get the table type and name. */
	SELECT @iScreenTableID = ASRSysScreens.tableID,
		@iScreenTableType = ASRSysTables.tableType,
		@sScreenTableName = ASRSysTables.tableName,
		@iScreenOrderID = 
				CASE 
					WHEN ASRSysScreens.orderID > 0 THEN ASRSysScreens.orderID
					ELSE ASRSysTables.defaultOrderID 
				END
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.ScreenID = @piScreenID;
	
	IF @iScreenOrderID IS NULL SET @iScreenOrderID = 0;
	
	IF @piOrderID <= 0 SET @piOrderID = @iScreenOrderID;
	
	/* Get the real source of the given screen's table/view. */
	IF @iScreenTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @piViewID;
		END
		ELSE
		BEGIN
			/* RealSource is the table. */	
			SET @sRealSource = @sScreenTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @iScreenTableID
			AND role = @sRoleName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sScreenTableName, ' ', '_') +
				'#' + replace(@sRoleName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END

	/* Initialise the select and order parameters. */
	SET @psSelectSQL = '';
	SET @psFromDef = '';
	SET @psOrderSQL = '';
	SET @sJoinCode = '';
	
	-- Create a temporary table to hold the tables/views that need to be joined.
	DECLARE @JoinParents TABLE(tableViewName sysname, tableID int);
	
	-- Create a temporary table of the column permissions for all tables/views used in the screen.
	DECLARE @columnPermissions TABLE(tableID integer,
					tableViewName	sysname,
					columnName		sysname,
					action			int,		
					granted			bit);

	-- Temporary view of the sysprotects
	DECLARE @SysProtects TABLE ([ID] int,
					action tinyint,
					protecttype tinyint,
					columns varbinary(8000));
	INSERT INTO @SysProtects
	SELECT DISTINCT p.[ID], p.[Action], p.[ProtectType], p.[Columns]
		FROM sys.sysprotects p
		WHERE (p.[Action] = 193 OR p.[Action] = 197)
			AND [uid] = @iUserGroupID;

	-- Loop through the tables used in the screen, getting the column permissions for each one.
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID
	FROM ASRSysControls
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0
	UNION
	SELECT DISTINCT c.tableID 
	FROM ASRSysOrderItems oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	WHERE oi.type = 'O' AND oi.orderID = @piOrderID;
	
	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN

		IF @iTempTableID = @iScreenTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @columnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				p.action,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @SysProtects p
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND OBJECT_NAME(p.ID) = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			SELECT @iParentTableType = tableType,
				@sParentTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @iTempTableID;
			
			IF @iParentTableType <> 2 /* ie. top-level or lookup */
			BEGIN

				INSERT INTO @columnPermissions
				SELECT 
					@iTempTableID,
					OBJECT_NAME(p.ID),
					syscolumns.name,
					p.action,
					CASE protectType
					   	WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @SysProtects p
				INNER JOIN syscolumns ON p.id = syscolumns.id
				WHERE syscolumns.name <> 'timestamp'
					AND OBJECT_NAME(p.id) IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @iTempTableID 
						UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
			END
			ELSE
			BEGIN

				/* Get permitted child view on the parent table. */
				SELECT @iParentChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTempTableID
					AND role = @sRoleName;
					
				IF @iParentChildViewID IS null SET @iParentChildViewID = 0;
					
				IF @iParentChildViewID > 0 
				BEGIN
					SET @sParentRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iParentChildViewID) +
						'#' + replace(@sParentTableName, ' ', '_') +
						'#' + replace(@sRoleName, ' ', '_');
					SET @sParentRealSource = left(@sParentRealSource, 255);
					INSERT INTO @columnPermissions
					SELECT 
						@iTempTableID,
						@sParentRealSource,
						syscolumns.name,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @sysprotects p
					INNER JOIN syscolumns ON p.id = syscolumns.id
					WHERE syscolumns.name <> 'timestamp'
						AND OBJECT_NAME(p.ID) = @sParentRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

				END
			END
		END
		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	SET @iUserType = 1;
	
	SELECT @iID = ASRSysPermissionItems.itemID
		FROM ASRSysPermissionItems
		INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
		WHERE ASRSysPermissionItems.itemKey = 'INTRANET'
			AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS';
			
	IF @iID IS NULL SET @iID = 0;
	IF @iID > 0
	BEGIN
		/* The permission does exist in the current version so check if the user is granted this permission. */
		SELECT @iCount = count(*)
		FROM ASRSysGroupPermissions 
		WHERE ASRSysGroupPermissions.itemID = @iID
			AND ASRSysGroupPermissions.groupName = @sRoleName
			AND ASRSysGroupPermissions.permitted = 1;
			
		IF @iCount > 0 SET @iUserType = 0;

	END
	/* Create a temporary table of the column info for all columns used in the screen controls. */
	/* Populate the temporary table with info for all columns used in the screen controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID, 
		ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysTables.tableName,
		ASRSysColumns.dataType,
		ASRSysColumns.columnType,
		ASRSysColumns.linkTableID
	FROM ASRSysControls
		LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
		LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID AND ASRSysColumns.columnId = ASRSysControls.columnID
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0;
	
	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID;	
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;
		SET @fUpdateGranted = 0;
		IF @iColumnTableID = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193;
				
			/* Get the update permission on the column. */
			SELECT @fUpdateGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 197;

			/* If the column is a link column, ensure that the link table can be seen. */
			IF (@fUpdateGranted = 1) AND (@iColumnType = 4)
			BEGIN
				SELECT @sLinkTableName = tableName,
					@iLinkTableType = tableType
				FROM ASRSysTables
				WHERE tableID = @iLinkTableID;
				
				IF @iLinkTableType = 1
				BEGIN
					/* Top-level table. */
					SELECT @lngPermissionCount = COUNT(*)
					FROM @sysprotects p
					INNER JOIN syscolumns ON p.id = syscolumns.id
					WHERE p.action = 193
						AND p.protectType <> 206
						AND syscolumns.name <> 'timestamp'
						AND syscolumns.name <> 'ID'
						AND OBJECT_NAME(p.ID) = @sLinkTableName
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
						
					IF @lngPermissionCount = 0 
					BEGIN
						/* No permission on the table itself check the views. */
						SELECT @lngPermissionCount = COUNT(*)
						FROM ASRSysViews
						INNER JOIN sysobjects ON ASRSysViews.viewName = sysobjects.name
						INNER JOIN @sysprotects p ON sysobjects.id = p.id  
						WHERE ASRSysViews.viewTableID = @iLinkTableID
							AND p.action = 193
							AND p.protecttype <> 206;
						IF @lngPermissionCount = 0 SET @fUpdateGranted = 0;
					END
				END
				ELSE
				BEGIN
					/* Child/history table. */
					SELECT @iLinkChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @iLinkTableID
						AND role = @sRoleName;
						
					IF @iLinkChildViewID IS null SET @iLinkChildViewID = 0;
						
					IF @iLinkChildViewID > 0 
					BEGIN
						SET @sLinkRealSource = 'ASRSysCV' + 
							convert(varchar(1000), @iLinkChildViewID) +
							'#' + replace(@sLinkTableName, ' ', '_') +
							'#' + replace(@sRoleName, ' ', '_');
						SET @sLinkRealSource = left(@sLinkRealSource, 255);
					END
					SELECT @lngPermissionCount = COUNT(p.ID)
					FROM @sysprotects p 
					WHERE p.protectType <> 206
						AND p.action = 193
						AND OBJECT_NAME(p.ID) = @sLinkRealSource;
		
					IF @lngPermissionCount = 0 SET @fUpdateGranted = 0;
				END
			END

			IF @fSelectGranted = 1 
			BEGIN
				/* Get the select string for the column. */
				IF LEN(@psSelectSQL) > 0 
					SET @psSelectSQL = @psSelectSQL + ',';
			
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sNewBit = 'convert(varchar(10), ' + @sRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sNewBit = @sRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
			END
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sColumnTableName
				AND columnName = @sColumnName
				AND action = 193;
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN

				/* Column COULD be read directly from the parent table. */
				IF len(@psSelectSQL) > 0 
					SET @psSelectSQL = @psSelectSQL + ',';

				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sNewBit = 'convert(varchar(10), ' + @sColumnTableName + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sNewBit = @sColumnTableName + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
					SET @psSelectSQL = @psSelectSQL + @sNewBit;
				END
				
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @JoinParents
					WHERE tableViewName = @sColumnTableName;
					
				IF @iTempCount = 0
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
					
			END
			ELSE	
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */
				SET @sSelectString = '';
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT tableViewName
					FROM @columnPermissions
					WHERE tableID = @iColumnTableID
						AND tableViewName <> @sColumnTableName
						AND columnName = @sColumnName
						AND action = 193
						AND granted = 1;
						
				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					SET @fSelectGranted = 1;
					IF len(@sSelectString) = 0 SET @sSelectString = 'CASE';
	
					IF @iColumnDataType = 11 /* Date */
					BEGIN
						 /* Date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN convert(varchar(10), ' + @sViewName + '.' + @sColumnName + ', 101)';
					END
					ELSE
					BEGIN
						 /* Non-date */
						SET @sSelectString = @sSelectString +
							' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
					END

					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
						WHERE tableViewName = @sViewName;
						
					IF @iTempCount = 0
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableID);

					FETCH NEXT FROM viewCursor INTO @sViewName;
				END

				CLOSE viewCursor;
				DEALLOCATE viewCursor;
				
				IF len(@sSelectString) > 0
				BEGIN
					SET @sSelectString = @sSelectString +
						' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']';
					IF len(@psSelectSQL) > 0 SET @psSelectSQL = @psSelectSQL + ',';
					SET @psSelectSQL = @psSelectSQL + @sSelectString;
				END
			END
		END
		FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID;
	END

	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;
	
	/* Create the order string. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT c.tableID, oi.columnID, c.columnName, t.tableName, oi.ascending
		FROM ASRSysOrderItems oi
			INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
			INNER JOIN ASRSysTables t ON t.tableID = c.tableID
		WHERE oi.orderID = @piOrderID AND oi.type = 'O'
			AND c.dataType <> -4 AND c.datatype <> -3
		ORDER BY oi.sequence;
		
	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;
		IF @iColumnTableId = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193;
			IF @fSelectGranted = 1
			BEGIN
				/* Get the order string for the column. */
				IF len(@psOrderSQL) > 0 SET @psOrderSQL = @psOrderSQL + ', ';
				SET @psOrderSQL = @psOrderSQL + @sRealSource + '.' + @sColumnName + CASE WHEN @fAscending = 0 THEN ' DESC' ELSE '' END;
			END
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName
				AND action = 193;
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* Get the order string for the column. */
				IF len(@psOrderSQL) > 0 
					SET @psOrderSQL = @psOrderSQL + ', ';
				SET @psOrderSQL = @psOrderSQL + @sColumnTableName + '.' + @sColumnName + CASE WHEN @fAscending = 0 THEN ' DESC' ELSE '' END;
				
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
					WHERE tableViewName = @sColumnTableName;
					
				IF @iTempCount = 0
					INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
					
			END
			ELSE	
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */
				SET @sOrderString = ''
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND action = 193
					AND granted = 1;
					
				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sOrderString) = 0 SET @sOrderString = 'CASE';
					SET @sOrderString = @sOrderString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
						WHERE tableViewname = @sViewName;
						
					IF @iTempCount = 0
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId);
						
					FETCH NEXT FROM viewCursor INTO @sViewName;
				END

				CLOSE viewCursor;
				DEALLOCATE viewCursor;
				
				IF len(@sOrderString) > 0
				BEGIN
					SET @sOrderString = @sOrderString +	' ELSE NULL END';
					IF len(@psOrderSQL) > 0 
						SET @psOrderSQL = @psOrderSQL + ', ';

					SET @psOrderSQL = @psOrderSQL + @sOrderString + CASE WHEN @fAscending = 0 THEN ' DESC' ELSE '' END;
				END
			END
		END
		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending;
	END
	
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	
	-- Add the ID column to the order string.
	IF LEN(@psOrderSQL) > 0 SET @psOrderSQL = @psOrderSQL + ', ';
	SET @psOrderSQL = @psOrderSQL + @sRealSource + '.ID';

	-- Add columns from the screen.
	SELECT @psSelectSQL = @psSelectSQL 
		+ CASE LEN(@psSelectSQL) WHEN 0 THEN '' ELSE ', ' END
		+ @sRealSource + '.' + [columnName]	+ ' AS [' + convert(varchar(10), [ColumnID]) + ']'
	FROM ASRSysColumns
	WHERE tableID = @iScreenTableID
		AND columnType = 3;

	-- Add timestamp to the select statement.
	SET @psSelectSQL = @psSelectSQL + ', CONVERT(integer, ' + @sRealSource + '.TimeStamp) AS timestamp ';

	-- Create the FROM code.
	SET @psFromDef = @sRealSource + '	';
	SELECT @psFromDef = @psFromDef + tableViewName + '	'
		+ convert(varchar(10), tableID) + '	'
	FROM @joinParents;

END