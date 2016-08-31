CREATE PROCEDURE [dbo].[sp_ASRIntGetViewColumns]
(
	@plngTableID 	integer, 
	@plngViewID 	integer,
	@psRealSource	varchar(8000) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the IDs and names of the columns available for the given table/view. */
	DECLARE @lngTableID		integer,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sRealSource 		varchar(255),
		@sTableName 		varchar(255),
		@iTableType			integer,
		@iChildViewID		integer,
		@sActualUserName	sysname;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Get the table ID from the view ID (if required). */
	IF @plngTableID > 0 
	BEGIN
		SET @lngTableID = @plngTableID;
	END
	ELSE
	BEGIN
		SELECT @lngTableID = ASRSysViews.viewTableID
		FROM [dbo].[ASRSysViews]
		WHERE ASRSysViews.viewID = @plngViewID;
	END

	/* Get the table-type. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM [dbo].[ASRSysTables]
	WHERE ASRSysTables.tableID = @lngTableID;

	/* Check if the current user is a System or Security manager. */
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1;
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
			AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS';
	END

	/* Create a temporary table to hold our resultset. */
	DECLARE @ColumnInfo TABLE(
		columnID	integer,
		columnName	sysname,
		dataType	integer,
		readGranted	bit,
		size		integer,
		decimals	integer);

	/* Get the real source of the given screen's table/view. */
	IF (@fSysSecMgr = 1 )
	BEGIN
		/* Populate the temporary table with information on the order for the given table. */
		IF @plngViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM [dbo].[ASRSysViews]
			WHERE viewID = @plngViewID;

	   		INSERT INTO @ColumnInfo (
				columnID, 
				columnName,
				dataType,
				readGranted,
				size,
				decimals)
			(SELECT ASRSysColumns.columnId, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				1,
				ASRSysColumns.size,
				ASRSysColumns.decimals				
			FROM ASRSysColumns
			INNER JOIN ASRSysViewColumns ON ASRSysColumns.columnId = ASRSysViewColumns.columnID
			WHERE ASRSysColumns.tableID = @lngTableID
            AND ASRSysColumns.columnType <> 3   -- Remove ID Columns
            AND ASRSysColumns.dataType <> -4    -- Remove OLE columns
				AND ASRSysViewColumns.viewID = @plngViewID
				AND ASRSysViewColumns.inView=1);
		END
		ELSE
		BEGIN
			IF @iTableType <> 2 /* ie. top-level or lookup */
			BEGIN
				/* RealSource is the table. */	
				SET @sRealSource = @sTableName;
			END 
			ELSE
			BEGIN
				SELECT @iChildViewID = childViewID
				FROM [dbo].[ASRSysChildViews2]
				WHERE tableID = @lngTableID
					AND role = @sUserGroupName;
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_');
					SET @sRealSource = left(@sRealSource, 255);
				END
			END

	   		INSERT INTO @ColumnInfo (
				columnID, 
				columnName,
				dataType,
				readGranted,
				size,
				decimals)
			(SELECT ASRSysColumns.columnId, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				1,
				ASRSysColumns.size,
				ASRSysColumns.decimals				
			FROM ASRSysColumns
			WHERE ASRSysColumns.tableID = @lngTableID
               AND columnType <> 3
               AND dataType <> -4);
		END
	END
	ELSE
	BEGIN
		IF @iTableType <> 2 /* ie. top-level or lookup */
		BEGIN
			IF @plngViewID > 0 
			BEGIN	
				/* RealSource is the view. */	
				SELECT @sRealSource = viewName
				FROM [dbo].[ASRSysViews]
				WHERE viewID = @plngViewID;
			END
			ELSE
			BEGIN
				SET @sRealSource = @sTableName;
			END 
		END
		ELSE
		BEGIN
			SELECT @iChildViewID = childViewID
			FROM [dbo].[ASRSysChildViews2]
			WHERE tableID = @lngTableID
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

   		 INSERT INTO @ColumnInfo (
			columnID, 
			columnName,
			dataType,
			readGranted,
			size,
			decimals)
		(SELECT 
			ASRSysColumns.columnId,
			syscolumns.name,
			ASRSysColumns.dataType,
			CASE protectType
				WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END,
			ASRSysColumns.size,
			ASRSysColumns.decimals				
		FROM ASRSysProtectsCache p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		INNER JOIN syscolumns ON p.id = syscolumns.id
		INNER JOIN ASRSysColumns ON syscolumns.name = ASRSysColumns.columnName
		WHERE p.action = 193 
			AND p.uid = @iUserGroupID
			AND ASRSysColumns.tableID = @lngTableID
         AND ASRSysColumns.columnType <> 3
         AND ASRSysColumns.dataType <> -4
			AND syscolumns.name <> 'timestamp'
			AND sysobjects.name = @sRealSource
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)));
	END

	SET @psRealSource = @sRealSource;

	/* Return the resultset. */
	SELECT columnID, columnName, dataType, size, decimals
	FROM @ColumnInfo 
	WHERE readGranted = 1
	ORDER BY columnName;
END