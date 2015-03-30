CREATE PROCEDURE [dbo].[sp_ASRIntGetOrderSQL] (
	@piScreenID 	integer,
	@piViewID 		integer,
	@piOrderID		integer,
	@psFromDef		varchar(MAX) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iScreenTableID		integer,
		@iScreenTableType	integer,
		@sScreenTableName	varchar(255),
		@fSysSecMgr			bit,
		@sRealSource 		varchar(255),
		@sParentSource		varchar(255),
		@iChildViewID 		integer,
		@sJoinCode 			varchar(MAX),
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@sColumnTableName 	varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sViewName 			varchar(255),
		@fAscending 		bit,
		@sTableViewName 	varchar(255),
		@iJoinTableID 		integer,
		@sParentRealSource	varchar(255),
		@iParentTableType	integer,
		@sParentTableName	sysname,
		@sActualUserName	sysname;

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Get the table type and name. */
	SELECT @iScreenTableID = ASRSysScreens.tableID,
		@iScreenTableType = ASRSysTables.tableType,
		@sScreenTableName = ASRSysTables.tableName
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.ScreenID = @piScreenID

	/* Check if the current user is a System or Security manager. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1
	END
	ELSE
	BEGIN	

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
	END

	/* Get the real source of the given screen's table/view. */
	IF @iScreenTableType <> 2 /* ie. top-level or lookup */
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
			/* RealSource is the table. */	
			SET @sRealSource = @sScreenTableName
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @iScreenTableID
			AND role = @sUserGroupName
			
		IF @iChildViewID IS null SET @iChildViewID = 0
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sScreenTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_')
			SET @sRealSource = left(@sRealSource, 255)
		END
	END

	/* Initialise the select and order parameters. */
	SET @psFromDef = ''
	SET @sJoinCode = ''

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(
		tableViewName	sysname,
		tableID			integer);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @columnPermissions TABLE(
		tableID			integer,
		tableViewName	sysname,
		columnName		sysname,
		granted			bit);

	/* Loop through the controls used in the screen, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID
	FROM ASRSysControls
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0
	UNION
	SELECT DISTINCT c.tableID 
	FROM ASRSysOrderItems oi
	INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	WHERE oi.type = 'O' AND oi.orderID = @piOrderID 

	OPEN tablesCursor
	FETCH NEXT FROM tablesCursor INTO @iTempTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @fSysSecMgr =1
		BEGIN
			IF @iTempTableID = @iScreenTableID
			BEGIN
				/* Base table - use the real source. */
				INSERT INTO @columnPermissions
				SELECT 
					@iTempTableID,
					@sRealSource,
					ASRSysColumns.columnName,
					1
				FROM ASRSysColumns
				WHERE ASRSysColumns.tableID = @iTempTableID
			END
			ELSE
			BEGIN
				/* Parent of the base table - get permissions for the table, and any associated views. */
				SELECT @iParentTableType = tableType,
					@sParentSource = tableName
				FROM ASRSysTables
				WHERE tableID = @iTempTableID

				IF @iParentTableType <> 2 
				BEGIN
					/* ie. top-level or lookup */
					INSERT INTO @columnPermissions
					SELECT 
						@iTempTableID,
						@sParentSource,
						ASRSysColumns.columnName,
						1
					FROM ASRSysColumns
					WHERE ASRSysColumns.tableID = @iTempTableID
				END	
				ELSE
				BEGIN
					/* RealSource is the child view on the table which is derived from full access on the table's parents. */	
					SELECT @iChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @iTempTableID
						AND role = @sUserGroupName
						
					IF @iChildViewID IS null SET @iChildViewID = 0
						
					IF @iChildViewID > 0 
					BEGIN
						SET @sParentSource = 'ASRSysCV' + 
							convert(varchar(1000), @iChildViewID) +
							'#' + replace(@sParentSource, ' ', '_') +
							'#' + replace(@sUserGroupName, ' ', '_')
						SET @sParentSource = left(@sParentSource, 255)
					END

					INSERT INTO @columnPermissions
					SELECT 
						@iTempTableID,
						@sParentSource,
						ASRSysColumns.columnName,
						1
					FROM ASRSysColumns
					WHERE ASRSysColumns.tableID = @iTempTableID
				END
			END
		END
		ELSE
		BEGIN
			IF @iTempTableID = @iScreenTableID
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
				SELECT @iParentTableType = tableType,
					@sParentTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @iTempTableID

				IF @iParentTableType <> 2 
				BEGIN
					/* ie. top-level or lookup */
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
						AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @iTempTableID 
							UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
						AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				END	
				ELSE
				BEGIN
					/* Get permitted child view on the parent table. */
					SELECT @iChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @iTempTableID
						AND role = @sUserGroupName
						
					IF @iChildViewID IS null SET @iChildViewID = 0
						
					IF @iChildViewID > 0 
					BEGIN
						SET @sParentRealSource = 'ASRSysCV' + 
							convert(varchar(1000), @iChildViewID) +
							'#' + replace(@sParentTableName, ' ', '_') +
							'#' + replace(@sUserGroupName, ' ', '_')
						SET @sParentRealSource = left(@sParentRealSource, 255)

						INSERT INTO @columnPermissions
						SELECT 
							@iTempTableID,
							@sParentRealSource,
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
							AND sysobjects.name = @sParentRealSource
							AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
							AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
							OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
							AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
					END
				END
			END
		END
		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	/* Create a temporary table of the column info for all columns used in the screen controls. */
	/* Populate the temporary table with info for all columns used in the screen controls. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.tableID, 
		ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysTables.tableName,
		ASRSysColumns.dataType
	FROM ASRSysControls
	LEFT OUTER JOIN ASRSysTables ON ASRSysControls.tableID = ASRSysTables.tableID 
	LEFT OUTER JOIN ASRSysColumns ON ASRSysColumns.tableID = ASRSysControls.tableID AND ASRSysColumns.columnId = ASRSysControls.columnID
	WHERE screenID = @piScreenID
		AND ASRSysControls.columnID > 0

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF @iColumnTableID <> @iScreenTableID
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = granted
			FROM @columnPermissions
			WHERE tableViewName = @sColumnTableName
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
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
				/* Column could NOT be read directly from the parent table, so try the views. */
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableID
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND granted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					SET @fSelectGranted = 1 

					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewName = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableID)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor
			END
		END

		FETCH NEXT FROM columnsCursor INTO @iColumnTableID, @iColumnID, @sColumnName, @sColumnTableName, @iColumnDataType
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
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName
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

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				
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
				/* Column could NOT be read directly from the parent table, so try the views. */
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND granted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
		
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
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Create the FROM code. */
	SET @psFromDef = @sRealSource + '	'
	DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableViewName, 
		tableID
	FROM @joinParents

	OPEN joinCursor
	FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @psFromDef = @psFromDef + @sTableViewName + '	' + convert(varchar(100), @iJoinTableID) + '	'

		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
	END
	CLOSE joinCursor
	DEALLOCATE joinCursor

END