CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenControlsString2] (
	@piScreenID 	integer,
	@piViewID 		integer,
	@psSelectSQL	varchar(MAX) OUTPUT,
	@psFromDef		varchar(MAX) OUTPUT,
	@piOrderID		integer	OUTPUT
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
		@sNewBit			varchar(max),
		@iID				integer,
		@iCount				integer,
		@iUserType			integer,
		@sRoleName			sysname,
		@iEmployeeTableID	integer,
		@sActualUserName	sysname,
		@AppName varchar(50),
		@ItemKey varchar(20);

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;


	DECLARE @SysProtects TABLE([ID] int, Columns varbinary(8000)
								, [Action] tinyint
								, ProtectType tinyint)
	INSERT INTO @SysProtects
	SELECT p.[ID], p.[Columns], p.[Action], p.ProtectType FROM ASRSysProtectsCache p
		INNER JOIN SysColumns c ON (c.id = p.id
			AND c.[Name] = 'timestamp'
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0)))
		WHERE p.UID = @iUserGroupID
			AND [ProtectType] IN (204, 205)
			AND [Action] IN (193, 197);


	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @JoinParents TABLE(tableViewName	sysname,
								tableID		integer);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @ColumnPermissions TABLE(tableID		integer,
										tableViewName	sysname,
										columnName	sysname,
										action		int,		
										granted		bit);


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
	SET @sJoinCode = '';

	/* Loop through the tables used in the screen, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID
	FROM ASRSysControls
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0
	UNION
	SELECT DISTINCT c.tableID 
	FROM ASRSysOrderItems oi
	INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	WHERE oi.type = 'O' 
		AND oi.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @iScreenTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
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
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			SELECT @iParentTableType = tableType,
				@sParentTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @iTempTableID

			IF @iParentTableType <> 2 /* ie. top-level or lookup */
			BEGIN
				INSERT INTO @ColumnPermissions
				SELECT 
					@iTempTableID,
					sysobjects.name,
					syscolumns.name,
					p.[action],
					CASE p.protectType
					        	WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM @sysprotects p
				INNER JOIN sysobjects ON p.id = sysobjects.id
				INNER JOIN syscolumns ON p.id = syscolumns.id
				WHERE syscolumns.name <> 'timestamp'
					AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @iTempTableID 
						UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
					AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			END
			ELSE
			BEGIN
				/* Get permitted child view on the parent table. */
				SELECT @iParentChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTempTableID
					AND role = @sRoleName
					
				IF @iParentChildViewID IS null SET @iParentChildViewID = 0
					
				IF @iParentChildViewID > 0 
				BEGIN
					SET @sParentRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iParentChildViewID) +
						'#' + replace(@sParentTableName, ' ', '_') +
						'#' + replace(@sRoleName, ' ', '_')
					SET @sParentRealSource = left(@sParentRealSource, 255)

					INSERT INTO @ColumnPermissions
					SELECT 
						@iTempTableID,
						@sParentRealSource,
						syscolumns.name,
						p.[action],
						CASE p.protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @sysprotects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					WHERE syscolumns.name <> 'timestamp'
						AND sysobjects.name = @sParentRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				END
			END
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	SET @iUserType = 1

	/*Ascertain application name in order to select by correct item key  */
	SELECT @AppName = APP_NAME()
	IF @AppName = 'OPENHR SELF-SERVICE INTRANET'
	BEGIN
		SET @ItemKey = 'SSINTRANET'
	END
	ELSE
	BEGIN
		SET @ItemKey = 'INTRANET'
	END

	SELECT @iID = ASRSysPermissionItems.itemID
	FROM ASRSysPermissionItems
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	WHERE ASRSysPermissionItems.itemKey = @ItemKey
		AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS'

	IF @iID IS NULL SET @iID = 0
	IF @iID > 0
	BEGIN
		/* The permission does exist in the current version so check if the user is granted this permission. */
		SELECT @iCount = count(ASRSysGroupPermissions.itemID)
		FROM ASRSysGroupPermissions 
		WHERE ASRSysGroupPermissions.itemID = @iID
			AND ASRSysGroupPermissions.groupName = @sRoleName
			AND ASRSysGroupPermissions.permitted = 1
			
		IF @iCount > 0
		BEGIN
			SET @iUserType = 0
		END
	END

	/* Get the EMPLOYEE table information. */
	SELECT @iEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel'
	IF @iEmployeeTableID IS NULL SET @iEmployeeTableID = 0

	/* Create a temporary table of the column info for all columns used in the screen controls. */
	DECLARE @columnInfo TABLE
	(
		columnID	integer,
		selectGranted	bit,
		updateGranted	bit
	)

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
	AND ASRSysControls.columnID > 0

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0
		SET @fUpdateGranted = 0

		IF @iColumnTableID = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193

			/* Get the update permission on the column. */
			SELECT @fUpdateGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 197

			/* If the column is a link column, ensure that the link table can be seen. */
			IF (@fUpdateGranted = 1) AND (@iColumnType = 4)
			BEGIN
				SELECT @sLinkTableName = tableName,
					@iLinkTableType = tableType
				FROM ASRSysTables
				WHERE tableID = @iLinkTableID

				IF @iLinkTableType = 1
				BEGIN
					/* Top-level table. */
					SELECT @lngPermissionCount = COUNT(sysprotects.uid)
					FROM sysprotects
					INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
					INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
					WHERE sysprotects.uid = @iUserGroupID
						AND sysprotects.action = 193
						AND sysprotects.protectType <> 206
						AND syscolumns.name <> 'timestamp'
						AND syscolumns.name <> 'ID'
						AND sysobjects.name = @sLinkTableName
						AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					IF @lngPermissionCount = 0 
					BEGIN
						/* No permission on the table itself check the views. */
						SELECT @lngPermissionCount = COUNT(ASRSysViews.viewTableID)
						FROM ASRSysViews
						INNER JOIN sysobjects ON ASRSysViews.viewName = sysobjects.name
						INNER JOIN sysprotects ON sysobjects.id = sysprotects.id  
						WHERE ASRSysViews.viewTableID = @iLinkTableID
							AND sysprotects.uid = @iUserGroupID
							AND sysprotects.action = 193
							AND sysprotects.protecttype <> 206

						IF @lngPermissionCount = 0 SET @fUpdateGranted = 0
					END
				END
				ELSE
				BEGIN
					/* Child/history table. */
					SELECT @iLinkChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @iLinkTableID
						AND role = @sRoleName
						
					IF @iLinkChildViewID IS null SET @iLinkChildViewID = 0
						
					IF @iLinkChildViewID > 0 
					BEGIN
						SET @sLinkRealSource = 'ASRSysCV' + 
							convert(varchar(1000), @iLinkChildViewID) +
							'#' + replace(@sLinkTableName, ' ', '_') +
							'#' + replace(@sRoleName, ' ', '_')
						SET @sLinkRealSource = left(@sLinkRealSource, 255)
					END

					SELECT @lngPermissionCount = COUNT(sysobjects.name)
					FROM sysprotects 
					INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
					WHERE sysprotects.uid = @iUserGroupID
						AND sysprotects.protectType <> 206
						AND sysprotects.action = 193
						AND sysobjects.name = @sLinkRealSource
		
					IF @lngPermissionCount = 0 SET @fUpdateGranted = 0
				END
			END

			IF @fSelectGranted = 1 
			BEGIN
				/* Get the select string for the column. */
				IF len(@psSelectSQL) > 0 
					SET @psSelectSQL = @psSelectSQL + ',';
			
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sNewBit = 'convert(varchar(10), ' + @sRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(255), @iColumnID) + ']';
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
			FROM @ColumnPermissions
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
				FROM @ColumnPermissions
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
						FROM @JoinParents
						WHERE tableViewName = @sViewName;

					IF @iTempCount = 0
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableID);

					FETCH NEXT FROM viewCursor INTO @sViewName;
				END
				CLOSE viewCursor;
				DEALLOCATE viewCursor;

				IF len(@sSelectString) > 0
				BEGIN
					SET @sSelectString = @sSelectString +
						' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']';

					IF len(@psSelectSQL) > 0 
						SET @psSelectSQL = @psSelectSQL + ',';

					SET @psSelectSQL = @psSelectSQL + @sSelectString;
				END
			END

			/* Reset the update permission on the column. */
			SET @fUpdateGranted = 0
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0
		IF @fUpdateGranted IS NULL SET @fUpdateGranted = 0

		IF (@iUserType = 1) 
			AND (@iScreenTableType = 1)
			AND (@iScreenTableID <> @iEmployeeTableID)
		BEGIN
			SET @fUpdateGranted = 0
		END

		INSERT INTO @columnInfo (columnID, selectGranted, updateGranted)
			VALUES (@iColumnId, @fSelectGranted, @fUpdateGranted)

		FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType, @iColumnType, @iLinkTableID
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor

	/* Create the order string. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT c.tableID, oi.columnID, c.columnName, t.tableName, oi.ascending
	FROM ASRSysOrderItems oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
		INNER JOIN ASRSysTables t ON t.tableID = c.tableID
	WHERE oi.orderID = @piOrderID	AND oi.type = 'O'
			AND c.dataType <> -4 AND c.datatype <> -3
	ORDER BY oi.sequence;

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF @iColumnTableId = @iScreenTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
				AND action = 193
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName
				AND action = 193

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				
				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @JoinParents
				WHERE tableViewName = @sColumnTableName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID)
				END
			END
			ELSE	
			BEGIN
				/* Column could NOT be read directly from the parent table, so try the views. */

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @ColumnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND action = 193
					AND granted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @JoinParents
					WHERE tableViewname = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Add the id and timestamp columns to the select string. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnId, 
		ASRSysColumns.columnName
	FROM ASRSysColumns
	WHERE tableID = @iScreenTableID
		AND columnType = 3

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName
	WHILE (@@fetch_status = 0)
	BEGIN
		IF len(@psSelectSQL) > 0 
			SET @psSelectSQL = @psSelectSQL + ',';

		SET @sNewBit = @sRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
		SET @psSelectSQL = @psSelectSQL + @sNewBit;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	SET @sNewBit = ', CONVERT(integer, ' + @sRealSource + '.TimeStamp) AS timestamp ';
	SET @psSelectSQL = @psSelectSQL + @sNewBit;

	/* Create the FROM code. */
	SET @psFromDef = @sRealSource + '	'
	DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableViewName, tableID
		FROM @JoinParents;

	OPEN joinCursor;
	FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @psFromDef = @psFromDef + @sTableViewName + '	' + convert(varchar(100), @iJoinTableID) + '	';
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
	END
	CLOSE joinCursor;
	DEALLOCATE joinCursor;

	SELECT
		convert(varchar(MAX), case when ASRSysControls.pageNo IS null then '' else ASRSysControls.pageNo end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.tableID IS null then '' else ASRSysControls.tableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.columnID IS null then '' else ASRSysControls.columnID end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.controlType IS null then '' else ASRSysControls.controlType end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.topCoord IS null then '' else ASRSysControls.topCoord end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.leftCoord IS null then '' else ASRSysControls.leftCoord end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.height IS null then '' else ASRSysControls.height end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.width IS null then '' else ASRSysControls.width end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.caption IS null then '' else ASRSysControls.caption end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.backColor IS null then '' else ASRSysControls.backColor end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.foreColor IS null then '' else ASRSysControls.foreColor end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontName IS null then '' else ASRSysControls.fontName end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontSize IS null then '' else ASRSysControls.fontSize end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontBold IS null then '' else ASRSysControls.fontBold end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontItalic IS null then '' else ASRSysControls.fontItalic end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontStrikethru IS null then '' else ASRSysControls.fontStrikethru end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.fontUnderline IS null then '' else ASRSysControls.fontUnderline end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.displayType IS null then '' else ASRSysControls.displayType end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.tabIndex IS null then '' else ASRSysControls.tabIndex end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.borderStyle IS null then '' else ASRSysControls.borderStyle end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.alignment IS null then '' else ASRSysControls.alignment end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnName IS null then '' else ASRSysColumns.columnName end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnType IS null then '' else ASRSysColumns.columnType end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.datatype IS null then '' else ASRSysColumns.datatype end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.defaultValue IS null then '' else ASRSysColumns.defaultValue end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.size IS null then '' else convert(nvarchar(max),ASRSysColumns.size) end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.decimals IS null then '' else ASRSysColumns.decimals end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.lookupTableID IS null then '' else ASRSysColumns.lookupTableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.lookupColumnID IS null then '' else ASRSysColumns.lookupColumnID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerMinimum IS null then '' else ASRSysColumns.spinnerMinimum end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerMaximum IS null then '' else ASRSysColumns.spinnerMaximum end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.spinnerIncrement IS null then '' else ASRSysColumns.spinnerIncrement end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.mandatory IS null then '' else ASRSysColumns.mandatory end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.uniquechecktype IS null then '' when ASRSysColumns.uniquechecktype <> 0 then 1 else 0 end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.convertcase IS null then '' else ASRSysColumns.convertcase end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.mask IS null then '' else rtrim(ASRSysColumns.mask) end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.blankIfZero IS null then '' else ASRSysColumns.blankIfZero end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.multiline IS null then '' else ASRSysColumns.multiline end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.alignment IS null then '' else ASRSysColumns.alignment end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.dfltValueExprID IS null then '' else ASRSysColumns.dfltValueExprID end) + char(9) +
		convert(varchar(MAX), case when isnull(ASRSysColumns.readOnly,0) = 1 then 1 else CASE WHEN ASRSysColumns.tableid = @iScreenTableID THEN 0 ELSE 1 END end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.statusBarMessage IS null then '' else ASRSysColumns.statusBarMessage end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkTableID IS null then '' else ASRSysColumns.linkTableID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkOrderID IS null then '' else ASRSysColumns.linkOrderID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.linkViewID IS null then '' else ASRSysColumns.linkViewID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.Afdenabled IS null then '' else ASRSysColumns.Afdenabled end) + char(9) +
		convert(varchar(MAX), case when ASRSysTables.TableName IS null then '' else ASRSysTables.TableName end) + char(9) +
		convert(varchar(MAX), case when ci.selectGranted IS null then '' else ci.selectGranted end) + char(9) +
		convert(varchar(MAX), case when ci.updateGranted IS null then '' else ci.updateGranted end) + char(9) +
		'' + char(9) +
		convert(varchar(MAX), case when ASRSysControls.pictureID IS null then '' else ASRSysControls.pictureID end)+ char(9) +
		convert(varchar(MAX), case when ASRSysColumns.trimming IS null then '' else ASRSysColumns.trimming end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.Use1000Separator IS null then '' else ASRSysColumns.Use1000Separator end) + char(9) +	
		convert(varchar(MAX), case when ASRSysColumns.lookupFilterColumnID IS null then '' else ASRSysColumns.lookupFilterColumnID end) + char(9) +	
		convert(varchar(MAX), case when ASRSysColumns.LookupFilterValueID IS null then '' else ASRSysColumns.LookupFilterValueID end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.OLEType IS null then '' else ASRSysColumns.OLEType end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.MaxOLESizeEnabled IS null then '' else ASRSysColumns.MaxOLESizeEnabled end) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.MaxOLESize IS null then '' else ASRSysColumns.MaxOLESize end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateTo IS null then '' else ASRSysControls.NavigateTo end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateIn IS null then '' else ASRSysControls.NavigateIn end) + char(9) +
		convert(varchar(MAX), case when ASRSysControls.NavigateOnSave IS null then '' else ASRSysControls.NavigateOnSave end) + char(9) +
		convert(varchar(MAX), case when isnull(ASRSysControls.readOnly,0) = 1 then 1 else 0 end)
		AS [controlDefinition],
		ASRSysControls.pageNo AS [pageNo],
		ASRSysControls.controlLevel AS [controlLevel],
		ASRSysControls.tabIndex AS [tabIndex]
	FROM ASRSysControls
	LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
	LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID AND ASRSysColumns.columnId = ASRSysControls.columnID
	LEFT OUTER JOIN @columnInfo ci ON ASRSysColumns.columnId = ci.columnID
	WHERE screenID = @piScreenID
	UNION
	SELECT 
		convert(varchar(MAX), -1) + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnId IS null then '' else ASRSysColumns.columnId end)  + char(9) +
		convert(varchar(MAX), case when ASRSysColumns.columnName IS null then '' else ASRSysColumns.columnName end) 
		AS [controlDefinition],
		0 AS [pageNo],
		0 AS [controlLevel],
		0 AS [tabIndex]
	FROM ASRSysColumns
	WHERE tableID = @iScreenTableID
		AND columnType = 3
	ORDER BY [pageNo],
		[controlLevel] DESC, 
		[tabIndex];

END