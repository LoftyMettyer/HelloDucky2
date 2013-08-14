CREATE PROCEDURE [dbo].[sp_ASRIntGetTBEmployeeColumns] (
	@psErrorMsg 			varchar(MAX)	OUTPUT,
	@ps1000SeparatorCols 	varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the 'employee' table find columns that the user has 'read' permission on. */
	DECLARE 
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@iEmployeeTableID	integer,
		@iOrderID			integer,
		@sColumnName 		sysname,
		@iDataType 			integer,
		@iTableID 			integer,
		@iCount				integer,
		@sTemp				sysname,
		@iIndex				integer,
		@sEmpRealSource		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@iTempTableType		integer,
		@sTempTableName		sysname,
		@iChildViewID		integer,
		@sTempRealSource	sysname,
		@iViewID			integer, 
		@sViewName			sysname,
		@iTempID			integer,
		@sActualUserName	sysname,
		@iTempTableID 		integer,
		@fSelectGranted 	bit,
		@bUse1000Separator	bit,
		@fSomeReadable		bit,
		@fViewReadable		bit;

	SET @psErrorMsg = ''
	SET @ps1000SeparatorCols = ''
	SET @fSomeReadable = 0
	
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Get the EMPLOYEE table information. */
	SELECT @iEmployeeTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable'
	IF @iEmployeeTableID IS NULL SET @iEmployeeTableID = 0

	SELECT @iOrderID = defaultOrderID, 
		@iTableType = tableType,
		@sTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iEmployeeTableID

	/* Get the real source of the employee table. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		SET @sEmpRealSource = @sTableName
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @iEmployeeTableID
			AND role = @sUserGroupName
			
		IF @iChildViewID IS null SET @iChildViewID = 0
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sEmpRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_')
			SET @sEmpRealSource = left(@sEmpRealSource, 255)
		END
	END	
	
	/* Create a temporary table to hold the find columns that the user can see. */
	DECLARE @findColumns TABLE(
		columnName		sysname,
		dataType		integer);

	DECLARE @columnPermissions TABLE(
		tableID			integer,
		tableViewName	sysname,
		columnName		sysname,
		selectGranted	bit);

	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID, ASRSysTables.tableType, ASRSysTables.tableName
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID= ASRSysTables.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F'

	OPEN tablesCursor
	FETCH NEXT FROM tablesCursor INTO @iTempTableID, @iTempTableType, @sTempTableName
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
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTempTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sTempRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTempTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sTempRealSource = left(@sTempRealSource, 255)

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
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			END
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID, @iTempTableType, @sTempTableName
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
		ASRSysColumns.columnName,
		ASRSysColumns.dataType,
		ASRSysColumns.tableID,
		ASRSysTables.tableType,
		ASRSysTables.tableName,
		ASRSysColumns.Use1000Separator
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F'
	ORDER BY ASRSysOrderItems.sequence

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName, @bUse1000Separator

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @psErrorMsg = 'Unable to read the Employee table default order.'
		RETURN
	END

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		/* Get the real source of the employee table. */
		IF @iTempTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			SET @sTempRealSource = @sTempTableName
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @iTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sTempRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTempTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sTempRealSource = left(@sTempRealSource, 255)
			END
		END	

		SELECT @fSelectGranted = selectGranted
		FROM @columnPermissions
		WHERE tableID = @iTableID
			AND tableViewName = @sTempRealSource
			AND columnName = @sColumnName

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0

		IF @fSelectGranted = 1
		BEGIN
			SET @fSomeReadable = 1

			SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
				CASE
					WHEN @bUse1000Separator = 1 THEN '1'
					ELSE '0'
				END

			INSERT INTO @findColumns (columnName, dataType) VALUES (@sColumnName, @iDataType)
		END
		ELSE
		BEGIN
			/* The column CANNOT be read from the Delegate table, or directly from a parent table.
			Try to read it from the views on the table. */
			SET @fViewReadable = 0

			DECLARE viewsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT viewID,
				viewName
			FROM ASRSysViews
			WHERE viewTableID = @iTableID

			OPEN viewsCursor
			FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName
			WHILE (@@fetch_status = 0)
			BEGIN
				SELECT @fSelectGranted = selectGranted
				FROM @columnPermissions
				WHERE tableID = @iTableID
					AND tableViewName = @sViewName
					AND columnName = @sColumnName

				IF @fSelectGranted IS NULL SET @fSelectGranted = 0

				IF @fSelectGranted = 1	
				BEGIN
					SET @fViewReadable = 1
				END

				FETCH NEXT FROM viewsCursor INTO @iViewID, @sViewName
			END
			CLOSE viewsCursor
			DEALLOCATE viewsCursor

			IF @fViewReadable = 1
			BEGIN
				INSERT INTO @findColumns (columnName, dataType) VALUES (@sColumnName, @iDataType)

				/* Add the column code to the 'select' string. */
				SET @fSomeReadable = 1

				SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
					CASE
						WHEN @bUse1000Separator = 1 THEN '1'
						ELSE '0'
					END
			END
		END

		FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTableID, @iTempTableType, @sTempTableName, @bUse1000Separator
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	IF @fSomeReadable = 0
	BEGIN
		/* Flag to the user that they cannot see any of the find columns. */
		SET @psErrorMsg = 'You do not have permission to read the Employee table find columns.'
	END
	ELSE
	BEGIN
		/* Add the ID column. */
		INSERT INTO @findColumns (columnName, dataType) VALUES ('ID', 4)
	END

	SELECT * FROM @findColumns;
	
END