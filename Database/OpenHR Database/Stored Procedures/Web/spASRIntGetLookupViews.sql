CREATE PROCEDURE [dbo].[spASRIntGetLookupViews]
(
	@plngTableID 		integer,
	@plngDfltOrderID 	integer	OUTPUT,
	@plngColumnID		integer
)
AS
BEGIN

	SET NOCOUNT ON;
	
	/* Return a recordset of the IDs and names of the views of the given table for use in the link find page. */
	DECLARE @sTableName 	varchar(255),
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@iChildViewID 		integer,
		@iTableType			integer,
		@lngPermissionCount	integer,
		@sLookupColumnName	sysname,
		@sFilterColumnName	sysname,
		@sRealSource 		varchar(MAX),
		@sActualUserName	sysname;

	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	/* Get the table-type. */
	SELECT @sTableName = ASRSysTables.tableName,
		@iTableType = ASRSysTables.tableType,
		@plngDfltOrderID = ASRSysTables.defaultOrderID
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @plngTableID

	SELECT @sLookupColumnName = a.columnName,
		@sFilterColumnName = b.columnName
	FROM ASRSysColumns
	LEFT OUTER JOIN ASRSysColumns a ON ASRSysColumns.lookupColumnID = a.columnID
	LEFT OUTER JOIN ASRSysColumns b ON ASRSysColumns.lookupFilterColumnID = b.columnID
	WHERE ASRSysColumns.columnID = @plngColumnID

	IF @sLookupColumnName IS null SET @sLookupColumnName = ''
	IF @sFilterColumnName IS null SET @sFilterColumnName = ''

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
			AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'
	END

	/* Create a temporary table to hold our resultset. */
	DECLARE @viewInfo TABLE(
		viewID			integer,
		viewName		sysname,
		orderTag		integer);

	/* Get the real source of the given screen's table/view. */
	IF (@fSysSecMgr = 1)
	BEGIN
		/* Populate the temporary table. */
   		INSERT INTO @viewInfo (
			viewID, 
			viewName,
			orderTag)
		VALUES (
			0,
			@sTableName,
			0	
		)
		
   		INSERT INTO @viewInfo (
			viewID, 
			viewName,
			orderTag)
		(SELECT viewID, 
			viewName,
			1
		FROM ASRSysViews
		WHERE viewTableID = @plngTableID
		)
	END
	ELSE
	BEGIN
		IF @iTableType <> 2
		BEGIN
			/* Table is a top-level or lookup table. */
			SELECT @lngPermissionCount = COUNT(*)
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193
				AND sysprotects.protectType <> 206
				AND syscolumns.name = @sLookupColumnName
				AND sysobjects.name = @sTableName
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

			IF @lngPermissionCount > 0 
			BEGIN
				IF len(@sFilterColumnName) > 0
				BEGIN
					SELECT @lngPermissionCount = COUNT(*)
					FROM sysprotects
					INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
					INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
					WHERE sysprotects.uid = @iUserGroupID
						AND sysprotects.action = 193
						AND sysprotects.protectType <> 206
						AND syscolumns.name = @sFilterColumnName
						AND sysobjects.name = @sTableName
						AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
						AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				END

				IF @lngPermissionCount > 0 
				BEGIN
					INSERT INTO @viewInfo (
						viewID, 
						viewName,
						orderTag)
					VALUES (
						0,
						@sTableName,
						0	
					)
				END
			END

			/* Now check on the views on this table. */
			INSERT INTO @viewInfo (
				viewID, 
				viewName,
				orderTag)
			(SELECT ASRSysViews.viewID, 
				ASRSysViews.viewName, 
				1
			FROM ASRSysViews
			INNER JOIN sysobjects ON ASRSysViews.viewName = sysobjects.name
			INNER JOIN sysprotects ON sysobjects.id = sysprotects.id  
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE ASRSysViews.viewTableID = @plngTableID
				AND sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193
				AND sysprotects.protecttype <> 206
				AND syscolumns.name = @sLookupColumnName
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			)

			IF len(@sFilterColumnName) > 0
			BEGIN
				DELETE FROM @viewInfo 
				WHERE viewID <> 0
					AND viewID NOT IN
				(SELECT ASRSysViews.viewID
				FROM ASRSysViews
				INNER JOIN sysobjects ON ASRSysViews.viewName = sysobjects.name
				INNER JOIN sysprotects ON sysobjects.id = sysprotects.id  
				INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
				WHERE ASRSysViews.viewTableID = @plngTableID
					AND sysprotects.uid = @iUserGroupID
					AND sysprotects.action = 193
					AND sysprotects.protecttype <> 206
					AND syscolumns.name = @sFilterColumnName
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				)
			END
		END
		ELSE
		BEGIN
			/* Get appropriate child view if required. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @plngTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sRealSource = left(@sRealSource, 255)

				SELECT @lngPermissionCount = COUNT(*)
				FROM sysprotects
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
				WHERE sysprotects.uid = @iUserGroupID
					AND sysprotects.action = 193
					AND sysprotects.protectType <> 206
					AND syscolumns.name = @sLookupColumnName
					AND sysobjects.name = @sRealSource
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
	
				IF @lngPermissionCount > 0 
				BEGIN
					IF len(@sFilterColumnName) > 0
					BEGIN
						SELECT @lngPermissionCount = COUNT(*)
						FROM sysprotects
						INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
						INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
						WHERE sysprotects.uid = @iUserGroupID
							AND sysprotects.action = 193
							AND sysprotects.protectType <> 206
							AND syscolumns.name = @sFilterColumnName
							AND sysobjects.name = @sRealSource
							AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
							AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
							OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
							AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
					END
	
					IF @lngPermissionCount > 0 
					BEGIN
						INSERT INTO @viewInfo (
							viewID, 
							viewName,
							orderTag)
						VALUES (
							0,
							@sTableName,
							0	
						)
					END
				END
			END
		END
	END

	/* Return the resultset. */
	SELECT viewID, viewName
	FROM @viewInfo
	ORDER BY orderTag, viewName
END
GO

