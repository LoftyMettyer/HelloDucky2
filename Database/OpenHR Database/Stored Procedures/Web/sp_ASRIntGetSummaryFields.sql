CREATE PROCEDURE [dbo].[sp_ASRIntGetSummaryFields] (
	@piHistoryTableID	integer,
	@piParentTableID 	integer,
	@piParentRecordID	integer,
	@pfCanSelect		bit OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@fSysSecMgr			bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(255),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(255),
		@sColumnName 		varchar(255),
		@fSelectGranted 	bit,
		@iCount				integer,
		@sActualUserName	sysname;

	SET @pfCanSelect = 0;

	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Check if the current user is a System or Security manager. */
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

	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID

	/* Create a temporary table of the 'read' column permissions for all tables/views used. */
	DECLARE @ColumnPermissions TABLE(
				tableViewName	sysname,
				columnName	sysname,
				granted		bit);

	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO @ColumnPermissions
		SELECT 
			@sParentTableName,
			ASRSysColumns.columnName,
			1
		FROM ASRSysColumns 
		WHERE ASRSysColumns.tableID = @piParentTableID
	END
	ELSE
	BEGIN
		IF @iParentTableType <> 2 /* ie. top-level or lookup */
		BEGIN

			-- Get list of views/table columns that are summary fields
			DECLARE @SummaryColumns TABLE ([ID] int, [TableName] sysname, [ColumnName] sysname, [ColID] int)
			INSERT @SummaryColumns
				SELECT sysobjects.id, sysobjects.name,
					syscolumns.name, syscolumns.ColID
				FROM sysobjects
				INNER JOIN syscolumns ON sysobjects.id = syscolumns.id
				WHERE sysobjects.name IN (SELECT ASRSysTables.tableName
												FROM ASRSysTables
												WHERE ASRSysTables.tableID = @piParentTableID
											UNION SELECT ASRSysViews.viewName
												FROM ASRSysViews
												WHERE ASRSysViews.viewTableID = @piParentTableID)
					AND syscolumns.name IN (SELECT ac.ColumnName
												FROM ASRSysSummaryFields am
												INNER JOIN ASRSysColumns ac ON am.ParentColumnID = ac.ColumnID
												WHERE HistoryTableID = @piHistoryTableID)

			-- Generate security context on selected columns
			INSERT INTO @ColumnPermissions
				SELECT sm.TableName,
					sm.ColumnName,
					CASE p.protectType
						WHEN 205 THEN 1
						WHEN 204 THEN 1
						ELSE 0
					END 
				FROM #SysProtects p
				INNER JOIN @SummaryColumns sm ON p.id = sm.id
				WHERE (((convert(tinyint,substring(p.columns,1,1))&1) = 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) != 0)
					OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
					AND (convert(int,substring(p.columns,sm.colid/8+1,1))&power(2,sm.colid&7)) = 0))
					AND p.Action = 193

		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @iChildViewID = childViewID
			FROM ASRSysChildViews2
			WHERE tableID = @piParentTableID
				AND role = @sUserGroupName
				
			IF @iChildViewID IS null SET @iChildViewID = 0
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sParentRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sParentTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_')
				SET @sParentRealSource = left(@sParentRealSource, 255)
			END

			INSERT INTO @ColumnPermissions
			SELECT 
				@sParentRealSource,
				syscolumns.name,
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM #SysProtects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
				AND p.Action = 193
		END
	END

	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnName
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @sColumnName
	WHILE (@@fetch_status = 0) AND (@pfCanSelect = 0)
	BEGIN
		SET @fSelectGranted = 0

		/* Get the select permission on the column. */

		/* Check if the column is selectable directly from the table. */
		SELECT @iCount = COUNT(*)
		FROM @ColumnPermissions
		WHERE columnName = @sColumnName
			AND granted = 1

		IF @iCount > 0 
		BEGIN
			SET @pfCanSelect = 1
		END

		FETCH NEXT FROM columnsCursor INTO @sColumnName
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor


	SELECT DISTINCT ASRSysSummaryFields.sequence, 
    	ASRSysSummaryFields.startOfGroup, 
		ASRSysColumns.columnName, 
		ASRSysColumns.columnID, 
		ASRSysColumns.dataType, 
		ASRSysColumns.size, 
		ASRSysColumns.decimals, 
		ASRSysColumns.controlType, 
		ASRSysColumns.alignment,
		ASRSysColumns.Use1000Separator
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence;
	
END