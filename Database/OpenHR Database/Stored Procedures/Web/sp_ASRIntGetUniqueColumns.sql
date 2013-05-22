CREATE PROCEDURE [dbo].[sp_ASRIntGetUniqueColumns]
(
	@plngTableID 	integer, 
	@plngViewID 	integer,
	@psRealSource	varchar(MAX) output
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the IDs and names of the unique columns available for the given table/view. */
	DECLARE @lngTableID		integer,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sRealSource 		varchar(MAX),
		@sTableName 		varchar(255),
		@iTableType			integer,
		@iChildViewID		integer,
		@sActualUserName	sysname;

	/* Get the current user's group ID. */
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
		SELECT @lngTableID = viewTableID
		FROM [dbo].[ASRSysViews]
		WHERE viewID = @plngViewID;
	END

	/* Get the table-type. */
	SELECT @iTableType = tableType,
		@sTableName = tableName
	FROM [dbo].[ASRSysTables]
	WHERE tableID = @lngTableID;

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
			AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';
	END

	/* Create a temporary table to hold our resultset. */
	DECLARE @columnInfo TABLE(
		columnID	integer,
		columnName	sysname,
		dataType	integer,
		columnSize	integer,
		readGranted	bit,
		columnDecimals	integer);

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

	   		INSERT INTO @columnInfo (
				columnID, 
				columnName,
				dataType,
				columnSize,
				readGranted,
				columnDecimals)
			(SELECT ASRSysColumns.columnID, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				ASRSysColumns.size,
				1, 
				ASRSysColumns.decimals
			FROM ASRSysColumns
			INNER JOIN ASRSysViewColumns ON ASRSysColumns.columnID = ASRSysViewColumns.columnID
			WHERE ASRSysColumns.tableID = @lngTableID
				AND ASRSysColumns.columnType <> 4
				AND ASRSysColumns.columnType <> 3
				AND (ASRSysColumns.dataType = 2
					OR ASRSysColumns.dataType = 4
					OR ASRSysColumns.dataType = 11
					OR ASRSysColumns.dataType =12)
				AND ASRSysColumns.uniqueCheckType <> 0
				AND ASRSysViewColumns.viewID = @plngViewID
				AND ASRSysViewColumns.inView = 1);
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
				/* RealSource is the child view on the table which is derived from full access on the table's parents. */	
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

	   		 INSERT INTO @columnInfo (
				columnID, 
				columnName,
				dataType,
				columnSize,
				readGranted,
				columnDecimals)
			(SELECT ASRSysColumns.columnID, 
				ASRSysColumns.columnName,
				ASRSysColumns.dataType,
				ASRSysColumns.size,
				1,
				ASRSysColumns.decimals
			FROM ASRSysColumns
			WHERE ASRSysColumns.tableID = @lngTableID
				AND columnType <> 4
				AND columnType <> 3
				AND (ASRSysColumns.dataType = 2
					OR ASRSysColumns.dataType = 4
					OR ASRSysColumns.dataType = 11
					OR ASRSysColumns.dataType =12)
				AND ASRSysColumns.uniqueCheckType <> 0);
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
			/* Get appropriate child view if required. */
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

   		 INSERT INTO @columnInfo (
			columnID, 
			columnName,
			dataType,
			columnSize,
			readGranted,
			columnDecimals)
		(SELECT 
			ASRSysColumns.columnID,
			syscolumns.name,
			ASRSysColumns.dataType,
			ASRSysColumns.size,
			CASE protectType
				WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END,
			ASRSysColumns.decimals
		FROM sysprotects
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
		INNER JOIN ASRSysColumns ON syscolumns.name = ASRSysColumns.columnName
		WHERE sysprotects.uid = @iUserGroupID
			AND sysprotects.action = 193 
			AND ASRSysColumns.tableID = @lngTableID
			AND ASRSysColumns.columnType <> 4
			AND ASRSysColumns.columnType <> 3
			AND (ASRSysColumns.dataType = 2
				OR ASRSysColumns.dataType = 4
				OR ASRSysColumns.dataType = 11
				OR ASRSysColumns.dataType =12)
			AND ASRSysColumns.uniqueCheckType <> 0
			AND syscolumns.name <> 'timestamp'
			AND sysobjects.name = @sRealSource
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)));
	END

	SET @psRealSource = @sRealSource;

	/* Return the resultset. */
	SELECT columnID, columnName, dataType, columnSize, columnDecimals
	FROM @columnInfo 
	WHERE readGranted = 1
	ORDER BY columnName;
	
END



















GO

