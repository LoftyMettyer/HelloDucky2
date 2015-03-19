CREATE PROCEDURE [dbo].[sp_ASRIntGetSelectedPicklistRecords]
	(
	@psSelectionType		varchar(255),
	@piSelectionID			integer,
	@psSelectedIDs			varchar(MAX),
	@psPromptSQL			varchar(MAX),
	@piTableID				integer,
	@psErrorMessage			varchar(MAX)	OUTPUT,
	@piExpectedRecords		integer			OUTPUT
	)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@iID				integer,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource		sysname,
		@sColumnName 		sysname,
		@iDataType 			integer,
		@iTableID 			integer,
		@iChildViewID		integer,
		@sTempRealSource	sysname,
		@iTemp				integer,
		@sTemp				varchar(MAX),
		@sSubViews			varchar(MAX),
		@sSQL 				nvarchar(MAX),
		@sPositionParamDefinition 	nvarchar(500),
		@sActualUserName	sysname,
		@sSelectSQL2		varchar(MAX),
		@sSelectSQL3		varchar(MAX),
		@sSelectSQL5		varchar(MAX),
		@sSelectSQL			varchar(MAX),
		@sFromSQL			varchar(MAX),
		@iOrderID			integer,
		@sJoinTables		varchar(MAX),
		@sJoinViews			varchar(MAX),
		@sWhereSQL			varchar(MAX),
		@fBaseSelect		bit,
		@iTempTableType		integer,
		@sTempTableName		sysname,
		@iTempTableID 		integer,
		@fSelectGranted 	bit,
		@iIndex				integer,
		@iViewID			integer, 
		@sViewName			sysname,
		@iTempID			integer,
		@sSubSQL			varchar(MAX),
		@sExecuteSQL		nvarchar(MAX);
	
	/* Clean the input string parameters. */
	IF len(@psSelectedIDs) > 0 SET @psSelectedIDs = replace(@psSelectedIDs, '''', '''''');

	SET @sSelectSQL = '';
	SET @sSelectSQL2 = '';
	SET @sSelectSQL3 = '';
	SET @sSelectSQL5 = '';
		
	SET @sJoinTables = ',';
	SET @sJoinViews = ',';
	SET @sWhereSQL = '';
	SET @fBaseSelect = 0;
	SET @sFromSQL = '';
	
	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	SELECT @iOrderID = defaultOrderID, 
		@iTableType = tableType,
		@sTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @piTableID;

	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		SET @sRealSource = @sTableName;
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piTableID
			AND role = @sUserGroupName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(CONVERT(varchar(255),@sTableName), ' ', '_') +
				'#' + replace(CONVERT(varchar(255),@sUserGroupName), ' ', '_');
			SET @sRealSource = left(CONVERT(varchar(255),@sRealSource), 255);
		END
	END	

	SET @sSelectSQL = '';

	/* Create a temporary table to hold the find columns that the user can see. */
	DECLARE @columnPermissions TABLE
		(tableID		integer,
		tableViewName	sysname,
		columnName	sysname,
		selectGranted	bit);

	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID, ASRSysTables.tableType, ASRSysTables.tableName
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID= ASRSysTables.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F';

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID, @iTempTableType, @sTempTableName;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableType <> 2 /* ie. top-level or lookup */
		BEGIN
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
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName 
					FROM ASRSysViews 
					WHERE ASRSysViews.viewTableID = @iTempTableID)
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTempTableID
				AND [role] = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sTempRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTempTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sTempRealSource = left(@sTempRealSource, 255);

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
					AND sysobjects.name =@sTempRealSource
						AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
			END
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID, @iTempTableType, @sTempTableName;
	END
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
		ASRSysColumns.columnName,
		ASRSysColumns.dataType,
		ASRSysColumns.tableID,
		ASRSysTables.tableType,
		ASRSysTables.tableName
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F'
		AND ASRSysColumns.datatype <> -3
		AND ASRSysColumns.datatype <> -4
	ORDER BY ASRSysOrderItems.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		IF @iTempTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			SET @sTempRealSource = @sTempTableName;
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
				AND [role] = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sTempRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTempTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sTempRealSource = left(@sTempRealSource, 255);
			END
		END	

		SELECT @fSelectGranted = selectGranted
		FROM @columnPermissions
		WHERE tableID = @iTableID
			AND tableViewName = @sTempRealSource
			AND columnName = @sColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		IF @fSelectGranted = 1
		BEGIN
			/* Add the column code to the 'select' string. */
			SET @sSelectSQL = @sSelectSQL +
				CASE 
					WHEN len(@sSelectSQL) > 0 THEN ','
					ELSE ''
				END +
				@sTempRealSource + '.' + @sColumnName;

			IF @iTableID = @piTableID
			BEGIN
				SET @fBaseSelect = 1;
			END 
			ELSE
			BEGIN
				/* Add the table to the list of join tables if required. */
				SELECT @iIndex = CHARINDEX(',' + @sTempRealSource + ',', @sJoinTables);
				IF @iIndex = 0 SET @sJoinTables = @sJoinTables + @sTempRealSource + ',';
			END
		END
		ELSE
		BEGIN
			/* The column CANNOT be read from the Delegate table, or directly from a parent table.
			Try to read it from the views on the table. */
			SET @sSubViews = ',';

			DECLARE viewsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT viewID,
				viewName
			FROM ASRSysViews
			WHERE viewTableID = @iTableID;

			OPEN viewsCursor;
			FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @fSelectGranted = 0

				SELECT @fSelectGranted = selectGranted
				FROM @columnPermissions
				WHERE tableID = @iTableID
					AND tableViewName = @sViewName
					AND columnName = @sColumnName;

				IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

				IF @fSelectGranted = 1	
				BEGIN
					/* Add the view to the list of join views if required. */
					SELECT @iIndex = CHARINDEX(',' + @sViewName + ',', @sJoinViews);
					IF @iIndex = 0 SET @sJoinViews = @sJoinViews + @sViewName + ',';

					SET @sSubViews = @sSubViews + @sViewName + ',';
				END

				FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName;
			END
			CLOSE viewsCursor;
			DEALLOCATE viewsCursor;

			IF len(@sSubViews) > 1
			BEGIN
				SET @sSubSQL = '';

				WHILE len(@sSubViews) > 1
				BEGIN
					SELECT @iIndex = charindex(',', @sSubViews, 2);
					SET @sViewName = substring(@sSubViews, 2, @iIndex - 2);
					SET @sSubViews = substring(@sSubViews, @iIndex, len(@sSubViews) -@iIndex + 1);

					IF len(@sSubSQL) > 0 SET @sSubSQL = @sSubSQL + ',';
					SET @sSubSQL = @sSubSQL + @sViewName + '.' + @sColumnName;
				END

				SET @sSubSQL = 'COALESCE(' + @sSubSQL + ', NULL) AS [' + @sColumnName + ']';
                
				/* Add the column code to the 'select' string. */
				SET @sSelectSQL = @sSelectSQL +
					CASE 
						WHEN len(@sSelectSQL) > 0 THEN ','
						ELSE ''
					END +
					@sSubSQL;
			END
		END

		FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName;
	END
	
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	/* Add the ID column. */
	SET @sFromSQL = ' FROM ' + @sRealSource;
        
	/* Join any other tables and views that are used. */
	WHILE len(@sJoinTables) > 1
	BEGIN
		SELECT @iIndex = charindex(',', @sJoinTables, 2);
		SET @sTableName = substring(@sJoinTables, 2, @iIndex - 2);
		SET @sJoinTables = substring(@sJoinTables, @iIndex, len(@sJoinTables) -@iIndex + 1);

		SELECT @iTempID = tableID
		FROM ASRSysTables
		WHERE tableName = @sTableName;

		SET @sFromSQL = @sFromSQL + 		
			' LEFT OUTER JOIN ' + @sTableName +
			' ON ' + @sRealSource + '.ID_' + convert(varchar(255), @iTempID) + ' = ' + @sTableName + '.ID';
	END

	WHILE len(@sJoinViews) > 1
	BEGIN
		SELECT @iIndex = charindex(',', @sJoinViews, 2);
		SET @sViewName = substring(@sJoinViews, 2, @iIndex - 2);
		SET @sJoinViews = substring(@sJoinViews, @iIndex, len(@sJoinViews) -@iIndex + 1);

		SELECT @iTempID = viewTableID
		FROM ASRSysViews
		WHERE viewName = @sViewName;

		IF @iTempID = @piTableID
		BEGIN
			SET @sFromSQL = @sFromSQL + 		
				' LEFT OUTER JOIN ' + @sViewName +
				' ON ' + @sRealSource + '.ID = ' + @sViewName + '.ID';

			IF @fBaseSelect = 0
			BEGIN
				SET @sWhereSQL = @sWhereSQL + 
					CASE
						WHEN len(@sWhereSQL) > 0 THEN ' OR ('
						ELSE '('
					END +
					@sRealSource + '.ID IN (SELECT ID FROM ' + @sViewName +  '))';
			END
		END
		ELSE
		BEGIN
			SET @sFromSQL = @sFromSQL + 		
				' LEFT OUTER JOIN ' + @sViewName +
				' ON ' + @sRealSource + '.ID_' + convert(varchar(255), @iTempID) + ' = ' + @sViewName + '.ID';
		END
	END

	IF len(@sWhereSQL) > 0
	BEGIN
		SET @sFromSQL = @sFromSQL + 
			' WHERE (' + @sWhereSQL + ')'	;
	END
	
	/* Get the list of selected IDs. */
	IF len(@psSelectedIDs) = 0 SET @psSelectedIDs = '0';

	/* PICKLIST = gets the items from the original definition */
	IF UPPER(@psSelectionType) = 'PICKLIST'
	BEGIN
		SET @psSelectedIDs = ' (SELECT recordID FROM ASRSysPicklistItems WHERE picklistID = ' + CONVERT(varchar(255), @piSelectionID) + ') ';
	END
		
	/* ALLRECORDS = gets all the remaining records */
	IF UPPER(@psSelectionType) <> 'ALLRECORDS'
	BEGIN
		/* Get the required find records. */
		SET @sSelectSQL2 = @sSelectSQL2 +
			CASE
				WHEN charindex(' WHERE ', @sFromSQL) > 0 THEN ' AND ('
				ELSE ' WHERE ('
			END +
			'(' + CONVERT(varchar(255), @sRealSource) + '.id IN (' ;
						
		SET @sSelectSQL3 = @psSelectedIDs;
		SET @sSelectSQL5 = @sSelectSQL5 + ')))';
	END

	/* FILTER = gets all the filtered records that are not yet selected */
	IF (UPPER(@psSelectionType) = 'FILTER') AND (len(@psPromptSQL) > 0)
	BEGIN
		SET @sSelectSQL5 = @sSelectSQL5 + ' OR (' + 
			CONVERT(varchar(255), @sRealSource) + '.id IN (' + @psPromptSQL + '))';
	END
	

	/* Add the 'order by part. */
	SET @sSelectSQL5 = @sSelectSQL5 +
		' ORDER BY 1';
		
	/* Count the number of commas before the ' FROM ' to see how many columns are in the select statement. */
	SELECT @iTemp = CHARINDEX(' FROM ', @sFromSQL);
	SET @sTemp = SUBSTRING(@sFromSQL, 1, @iTemp);
	SET @iTemp = 2;
	WHILE charindex(',', @sTemp) > 0
	BEGIN
		SET @sSelectSQL5 = @sSelectSQL5 +
			',' + convert(varchar(MAX), @iTemp);
		SET @sTemp = substring(@sTemp, charindex(',', @sTemp)+1, len(@sTemp) - charindex(',', @sTemp));
		SET @iTemp = @iTemp + 1;
	END
	
	SET @piExpectedRecords = 0;
	IF UPPER(@psSelectionType) = 'ALL'
	BEGIN
		SET @sSQL = 'SELECT @recordPosition = COUNT(ID)' +
			' FROM ' + CONVERT(varchar(255), @sRealSource) +
			' WHERE ID IN(' + @psSelectedIDs + ')';
		SET @sPositionParamDefinition = N'@recordPosition integer OUTPUT';
		EXEC sp_executesql @sSQL, @sPositionParamDefinition, @piExpectedRecords OUTPUT;
	END

	SET @sExecuteSQL = 'SELECT ' + @sSelectSQL + ',' +
		@sRealSource + '.id ' + @sFromSQL +
		@sSelectSQL2 + @sSelectSQL3 + @sSelectSQL5;

	-- Execute the generated string
	EXECUTE sp_executeSQL @sExecuteSQL;

END