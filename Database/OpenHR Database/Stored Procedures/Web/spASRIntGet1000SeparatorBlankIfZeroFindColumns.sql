CREATE PROCEDURE [dbo].[spASRIntGet1000SeparatorBlankIfZeroFindColumns] (
	@pfError 				bit 			OUTPUT, 
	@piTableID 				integer, 
	@piViewID 				integer, 
	@piOrderID 				integer, 
	@ps1000SeparatorCols	varchar(MAX)	OUTPUT,
	@psBlankIfZeroCols		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@sType	 			varchar(10),
		@fSelectGranted 	bit,
		@iCount				integer,
		@bUse1000Separator	bit,
		@bBlankIfZero		bit,
		@sActualLoginName	varchar(250);

	/* Initialise variables. */
	SET @pfError = 0;
	SET @ps1000SeparatorCols = '';
	SET @psBlankIfZeroCols = '';
	SET @sRealSource = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID;

	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1;
		RETURN;
	END

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
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
			SET @sRealSource = @sTableName;
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piTableID
			AND [role] = @sUserGroupName;
		
		IF @iChildViewID IS null SET @iChildViewID = 0;
		
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_');
			SET @sRealSource = left(@sRealSource, 255);
		END
	END

	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	-- Cached view of sysprotects
	DECLARE @SysProtects TABLE([ID] int, [ProtectType] tinyint, [Columns] varbinary(8000))
	INSERT INTO @SysProtects
		SELECT ID, ProtectType, [Columns] FROM ASRSysProtectsCache
		WHERE [UID] = @iUserGroupID AND Action = 193;

	/* Create a temporary table of the 'select' column permissions for all tables/views used in the order. */
	DECLARE @ColumnPermissions TABLE(
				tableID			integer,
				tableViewName	sysname,
				columnName		sysname,
				selectGranted	bit);

	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	WHERE ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
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
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				CASE p.protectType
				    WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM @Sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.tableID,
		ASRSysColumns.columnName,
		ASRSysTables.tableName,
		ASRSysOrderItems.type,
		ASRSysColumns.Use1000Separator,
		ASRSysColumns.BlankIfZero
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
		AND ASRSysOrderItems.type = 'F'
	ORDER BY ASRSysOrderItems.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @sColumnName, @sColumnTableName, @sType, @bUse1000Separator, @bBlankIfZero;

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
					CASE
						WHEN @bUse1000Separator = 1 THEN '1'
						ELSE '0'
					END;
				SET @psBlankIfZeroCols = @psBlankIfZeroCols +
					CASE
						WHEN @bBlankIfZero = 1 THEN '1'
						ELSE '0'
					END;
			END
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */
	
			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @ColumnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
					CASE
						WHEN @bUse1000Separator = 1 THEN '1'
						ELSE '0'
					END;
				SET @psBlankIfZeroCols = @psBlankIfZeroCols +
					CASE
						WHEN @bBlankIfZero = 1 THEN '1'
						ELSE '0'
					END;
			END
			ELSE	
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM @ColumnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1;

				IF @iCount > 0 
				BEGIN
					SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
						CASE
							WHEN @bUse1000Separator = 1 THEN '1'
							ELSE '0'
						END;
					SET @psBlankIfZeroCols = @psBlankIfZeroCols +
						CASE
							WHEN @bBlankIfZero = 1 THEN '1'
							ELSE '0'
						END;
				END
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @sColumnName, @sColumnTableName, @sType, @bUse1000Separator, @bBlankIfZero;
	END

	CLOSE orderCursor;
	DEALLOCATE orderCursor;

END