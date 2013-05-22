CREATE PROCEDURE [dbo].[spASRIntGetParentValues] (
	@piScreenID 		integer,
	@piParentTableID 	integer,
	@piParentRecordID 	integer
)
AS
BEGIN
	
	SET NOCOUNT ON;
	
	/* Return a recordset of the parent record values required for controls in the given screen. */
	DECLARE 
		@iUserGroupID		integer,
		@sRoleName			sysname,
		@iTempCount 		integer,
		@iParentTableType	integer,
		@sParentTableName	sysname,
		@iParentChildViewID	integer,
		@sParentRealSource	varchar(255),
		@iColumnID 			integer,
		@sColumnName 		varchar(255),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@sNewBit			varchar(MAX),
		@sSelectString 		varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	varchar(255),
		@sParentSelectSQL	nvarchar(MAX),
		@sTemp				varchar(MAX),
		@fColumns			bit,
		@sSQL				nvarchar(MAX),
		@sActualUserName	sysname;

	SET @sParentSelectSQL  = 'SELECT ';
	SET @fColumns = 0;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(tableViewName sysname);

	/* Create a temporary table of the column permissions for all tables/views used in the screen. */
	DECLARE @columnPermissions TABLE(tableViewName	sysname,
		columnName	sysname,
		granted		bit);

	SELECT @iTempCount = COUNT(*)
	FROM ASRSysControls
	INNER JOIN ASRSysColumns ON ASRSysControls.columnID = ASRSysColumns.columnID
		AND ASRSysColumns.tableID = @piParentTableID
	WHERE ASRSysControls.screenID = @piScreenID
		AND ASRSysControls.columnID > 0;

	IF @iTempCount = 0 RETURN;

	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables
		WHERE tableID = @piParentTableID;

	IF @iParentTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		INSERT INTO @columnPermissions
		SELECT 
			sysobjects.name,
			syscolumns.name,
			CASE p.protectType
			        	WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END 
		FROM #sysprotects p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		INNER JOIN syscolumns ON p.id = syscolumns.id
		WHERE p.action = 193 
			AND syscolumns.name <> 'timestamp'
			AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @piParentTableID 
			UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @piParentTableID)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		SET @sParentRealSource = @sParentTableName;
	END
	ELSE
	BEGIN
		/* Get permitted child view on the parent table. */
		SELECT @iParentChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piParentTableID
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
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	END

	/* Populate the temporary table with info for all columns used in the screen controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysControls.columnID, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysControls
	INNER JOIN ASRSysColumns ON ASRSysColumns.columnID = ASRSysControls.columnID
	WHERE ASRSysControls.screenID = @piScreenID
		AND ASRSysControls.columnID > 0
		AND ASRSysColumns.tableID = @piParentTableID;
	
	OPEN columnsCursor;
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;
	
		/* Get the select permission on the column. */
		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM @columnPermissions
		WHERE tableViewName = @sParentRealSource
			AND columnName = @sColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			IF @fColumns = 1
			BEGIN
				SET @sTemp = ',';
				SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;
			END

			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @sNewBit = 'convert(varchar(10), ' + @sParentRealSource + '.' + @sColumnName + ', 101) AS [' + convert(varchar(100), @iColumnID) + ']';
			END
			ELSE
			BEGIN
				 /* Non-date */
				SET @sNewBit = @sParentRealSource + '.' + @sColumnName + ' AS [' + convert(varchar(100), @iColumnID) + ']';
			END

			SET @fColumns = 1;
			SET @sParentSelectSQL = @sParentSelectSQL + @sNewBit;
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = '';

			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM @columnPermissions
			WHERE tableViewName <> @sParentRealSource
				AND columnName = @sColumnName
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
				BEGIN
					INSERT INTO @joinParents (tableViewName) VALUES(@sViewName);
				END

				FETCH NEXT FROM viewCursor INTO @sViewName;
			END
			CLOSE viewCursor;
			DEALLOCATE viewCursor;

			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = @sSelectString +
					' ELSE NULL END AS [' + convert(varchar(100), @iColumnID) + ']';

				IF @fColumns = 1
				BEGIN
					SET @sTemp = ',';
					SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;
				END

				SET @fColumns = 1;
				SET @sParentSelectSQL = @sParentSelectSQL + @sSelectString;
			END
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType;
	END
	CLOSE columnsCursor;
	DEALLOCATE columnsCursor;

	IF @fColumns = 0 RETURN;

	SET @sTemp = ' FROM ' + @sParentRealSource;
	SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

	DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT tableViewName
	FROM @joinParents;

	OPEN joinCursor;
	FETCH NEXT FROM joinCursor INTO @sTableViewName;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sTemp = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sParentRealSource + '.ID = ' + @sTableViewName + '.ID';
		SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

		FETCH NEXT FROM joinCursor INTO @sTableViewName;
	END
	CLOSE joinCursor;
	DEALLOCATE joinCursor;

	SET @sTemp = ' WHERE ' + @sParentRealSource + '.ID = ' + convert(varchar(100), @piParentRecordID);
	SET @sParentSelectSQL = @sParentSelectSQL + @sTemp;

	EXECUTE sp_executeSQL @sParentSelectSQL;
	
END