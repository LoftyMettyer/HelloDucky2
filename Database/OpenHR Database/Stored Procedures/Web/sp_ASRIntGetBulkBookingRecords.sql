CREATE PROCEDURE [dbo].[sp_ASRIntGetBulkBookingRecords] (
	@psSelectionType	varchar(MAX),
	@piSelectionID		integer,
	@psSelectedIDs		varchar(MAX),
	@psPromptSQL		varchar(MAX),
	@psErrorMessage		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the required 'employee' table records. */
	DECLARE
		@iUserGroupID			integer,
		@sUserGroupName			sysname,
		@iID					integer,
		@iEmployeeTableID		integer,
		@iTableType				integer,
		@sTableName				sysname,
		@sEmpRealSource			sysname,
		@iChildViewID			integer,
		@iTemp					integer,
		@sTemp					varchar(MAX),
		@sActualUserName		sysname,
		@sExecString			nvarchar(MAX),
		@sTempString			varchar(MAX),
		@iColumnCount			integer,
		@iOrderID				integer,
		@sSubSQL				varchar(MAX),
		@sJoinViews				varchar(MAX),
		@sSubViews				varchar(MAX),
		@sJoinTables			varchar(MAX),
		@fDelegateSelect		bit,
		@sWhereSQL				varchar(MAX),
		@iTempTableType			integer,
		@sTempTableName			sysname,
		@iTempTableID 			integer,
		@sTempRealSource		sysname,
		@sColumnName 			sysname,
		@iDataType 				integer,
		@iTableID 				integer,
		@fSelectGranted 		bit,
		@iIndex					integer,
		@iViewID				integer, 
		@sViewName				sysname,
		@iTempID				integer;
		
	/* Clean the input string parameters. */
	IF len(@psSelectedIDs) > 0 SET @psSelectedIDs = replace(@psSelectedIDs, '''', '''''');

	SET @sJoinViews = ','
	SET @sJoinTables = ',';
	SET @fDelegateSelect = 0;
	SET @sWhereSQL = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Get the EMPLOYEE table information. */
	SELECT @iEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable';
	IF @iEmployeeTableID IS NULL SET @iEmployeeTableID = 0;

	SELECT @iOrderID = defaultOrderID, 
		@iTableType = tableType,
		@sTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iEmployeeTableID;

	/* Get the real source of the employee table. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		SET @sEmpRealSource = @sTableName;
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @iEmployeeTableID
			AND [role] = @sUserGroupName;
			
		IF @iChildViewID IS null SET @iChildViewID = 0;
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sEmpRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_');
			SET @sEmpRealSource = left(@sEmpRealSource, 255);
		END
	END	
	
	SET @sExecString = 'SELECT ';
	SET @iColumnCount = 0;

	/* Create a temporary table to hold the find columns that the user can see. */
	DECLARE @columnPermissions TABLE(
		tableID			integer,
		tableViewName	sysname,
		columnName		sysname,
		selectGranted	bit);

	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT c.tableID, t.tableType, t.tableName
	FROM ASRSysOrderItems oi
	INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID
	WHERE oi.orderID = @iOrderID AND oi.type = 'F';

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
	SELECT c.columnName, c.dataType, c.tableID, t.tableType, t.tableName
	FROM ASRSysOrderItems oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
		INNER JOIN ASRSysTables t ON t.tableID = c.tableID
	WHERE oi.orderID = @iOrderID AND oi.type = 'F'
			AND c.dataType <> -4 AND c.datatype <> -3
	ORDER BY oi.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		/* Get the real source of the employee table. */
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
			SET @sTempString = CASE 
					WHEN @iColumnCount > 0 THEN ','
					ELSE ''
				END +
				@sTempRealSource + '.' + @sColumnName;
			SET @sExecString = @sExecString + @sTempString;
			SET @iColumnCount = @iColumnCount + 1;
				
			IF @iTableID = @iEmployeeTableID
			BEGIN
				SET @fDelegateSelect = 1;
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
				SET @fSelectGranted = 0;

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

				SET @sSubSQL = 'COALESCE(' + @sSubSQL + ') AS [' + @sColumnName + ']';
                
				/* Add the column code to the 'select' string. */
				SET @sTempString = CASE 
						WHEN @iColumnCount > 0 THEN ','
						ELSE ''
					END +
					@sSubSQL;

				SET @sExecString = @sExecString + @sTempString;
				SET @iColumnCount = @iColumnCount + 1;
			END
		END

		FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName;
	END
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	/* Add the ID column. */
	SET @iColumnCount = @iColumnCount + 1;
	SET @sTempString = ',' +
		@sEmpRealSource + '.id' +
		' FROM ' + @sEmpRealSource;
	SET @sExecString = @sExecString + @sTempString;
        
	/* Join any other tables and views that are used. */
	WHILE len(@sJoinTables) > 1
	BEGIN
		SELECT @iIndex = charindex(',', @sJoinTables, 2);
		SET @sTableName = substring(@sJoinTables, 2, @iIndex - 2);
		SET @sJoinTables = substring(@sJoinTables, @iIndex, len(@sJoinTables) -@iIndex + 1);

		SELECT @iTempID = tableID
		FROM ASRSysTables
		WHERE tableName = @sTableName;

		SET @sTempString = ' LEFT OUTER JOIN ' + @sTableName +
			' ON ' + @sEmpRealSource + '.ID_' + convert(varchar(8000), @iTempID) + ' = ' + @sTableName + '.ID';
		SET @sExecString = @sExecString + @sTempString;

	END

	WHILE len(@sJoinViews) > 1
	BEGIN
		SELECT @iIndex = charindex(',', @sJoinViews, 2);
		SET @sViewName = substring(@sJoinViews, 2, @iIndex - 2);
		SET @sJoinViews = substring(@sJoinViews, @iIndex, len(@sJoinViews) -@iIndex + 1);

		SELECT @iTempID = viewTableID
		FROM ASRSysViews
		WHERE viewName = @sViewName;

		IF @iTempID = @iEmployeeTableID
		BEGIN
			SET @sTempString = ' LEFT OUTER JOIN ' + @sViewName +
				' ON ' + @sEmpRealSource + '.ID = ' + @sViewName + '.ID';
			SET @sExecString = @sExecString + @sTempString;

			IF @fDelegateSelect = 0
			BEGIN
				SET @sWhereSQL = @sWhereSQL + 
					CASE
						WHEN len(@sWhereSQL) > 0 THEN ' OR ('
						ELSE '('
					END +
					@sEmpRealSource + '.ID IN (SELECT ID FROM ' + @sViewName +  '))';
			END
		END
		ELSE
		BEGIN
			SET @sTempString = ' LEFT OUTER JOIN ' + @sViewName +
				' ON ' + @sEmpRealSource + '.ID_' + convert(varchar(8000), @iTempID) + ' = ' + @sViewName + '.ID';
			SET @sExecString = @sExecString + @sTempString;
		END
	END

	IF len(@sWhereSQL) > 0
	BEGIN
		SET @sTempString = ' WHERE (' + @sWhereSQL + ')';
		SET @sExecString = @sExecString + @sTempString;
	END
	
	/* Get the list of selected IDs. */
	
	IF len(@psSelectedIDs) = 0 SET @psSelectedIDs = '0';
	
	IF UPPER(@psSelectionType) = 'PICKLIST'
	BEGIN
		DECLARE picklistCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT recordID
		FROM ASRSysPicklistItems
		WHERE picklistID = @piSelectionID;
	
		OPEN picklistCursor;
		FETCH NEXT FROM picklistCursor INTO @iID;

		WHILE (@@fetch_status = 0)
		BEGIN
			SET @psSelectedIDs = @psSelectedIDs + ',' + convert(varchar(255), @iID);
			FETCH NEXT FROM picklistCursor INTO @iID;
		END
		CLOSE picklistCursor;
		DEALLOCATE picklistCursor;
	END
	
	/* Get the required find records. */
	SET @sTempString = CASE
			WHEN (charindex(' WHERE ', @sExecString) > 0) THEN ' AND ('
			ELSE ' WHERE ('
		END +
		'(' + @sEmpRealSource + '.id IN (' + @psSelectedIDs + '))';
		
	SET @sExecString = @sExecString + @sTempString;

	IF (UPPER(@psSelectionType) = 'FILTER') AND (len(@psPromptSQL) > 0)
	BEGIN
		SET @sTempString = ' OR (' + 
			convert(varchar(255), @sEmpRealSource) + '.id IN (' + @psPromptSQL + '))';
		SET @sExecString = @sExecString + @sTempString;

	END

	SET @sTempString = ') ORDER BY 1';
	SET @sExecString = @sExecString + @sTempString;
		
	/* Count the number of commas before the ' FROM ' to see how many columns are in the select statement. */
	SET @iTemp = 2;
	WHILE @iTemp <= @iColumnCount
	BEGIN
		SET @sTempString = ',' + convert(varchar(8000), @iTemp);
		SET @sExecString = @sExecString + @sTempString;
		SET @iTemp = @iTemp + 1;
	END

	-- Return generated SQL	
	EXEC sp_executeSQL @sExecString;
	
END