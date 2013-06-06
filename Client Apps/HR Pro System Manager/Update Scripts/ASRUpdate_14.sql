/* -------------------------------------------------- */
/* Update the database from version 13 to version 14. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16)


/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */

/* Check if the database version column exists. */

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'databaseVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0
BEGIN
	/* The database version column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [databaseVersion] [int] NULL 
END

/* Check if the refreshStoredProcedures column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'refreshStoredProcedures'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The refreshStoredProcedures column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [refreshStoredProcedures] [bit] NULL 
END

/* Check if the systemManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'systemManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The systemManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SystemManagerVersion] [varchar] (50)NULL 
END

/* Check if the securityManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'securityManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The securityManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SecurityManagerVersion] [varchar] (50)NULL 
END

/* Check if the DataManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DataManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The DataManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [DataManagerVersion] [varchar] (50)NULL 
END

/* Check if the IntranetVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'IntranetVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The IntranetVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [IntranetVersion] [varchar] (50)NULL 
END


SET @sCommand = N'SELECT @iDBVersion = databaseVersion
	FROM ASRSysConfig'
SET @sParam = N'@iDBVersion integer OUTPUT'
execute sp_executesql @sCommand, @sParam, @iDBVersion OUTPUT

IF @iDBVersion IS null SET @iDBVersion = 0

/* Exit if the database is not version 13 or 14. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 13) or (@iDBVersion > 14)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------- */
/* Drop and recreate sp_ASRIntGetSummaryFields */
/* ------------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetSummaryFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetSummaryFields]

exec('CREATE PROCEDURE sp_ASRIntGetSummaryFields (
	@piHistoryTableID	integer,
	@piParentTableID 	integer,
	@piParentRecordID	integer,
	@psSelectSQL		varchar(8000) OUTPUT
)
AS
BEGIN
	/* Return a recordset of strings describing of the controls in the given screen. */
	DECLARE @iUserGroupID	integer,
		@fSysSecMgr		bit,
		@iParentTableType	integer,
		@sParentTableName	varchar(8000),
		@iChildViewID 		integer,
		@sParentRealSource 	varchar(8000),
		@iColumnID 		integer,
		@sColumnName 	varchar(8000),
		@iColumnDataType	integer,
		@fSelectGranted 	bit,
		@iTempCount 		integer,
		@sRootTable 		varchar(8000),
		@sSelectString 		varchar(8000),
		@sViewName 		varchar(8000),
		@sTableViewName 	varchar(8000)

	SET @psSelectSQL = ''''

	/* Get the current user''s group ID. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = CURRENT_USER

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
		OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

	/* Get the parent table type and name. */
	SELECT @iParentTableType = tableType,
		@sParentTableName = tableName
	FROM ASRSysTables 
	WHERE ASRSysTables.tableID = @piParentTableID

	/* Create a temporary table to hold the tables/views that need to be joined. */
	CREATE TABLE #joinParents
	(
		tableViewName	sysname
	)	

	/* Create a temporary table of the ''read'' column permissions for all tables/views used. */
	CREATE TABLE #columnPermissions
	(
		tableViewName	sysname,
		columnName	sysname,
		granted		bit		
	)

	/* Get the column permissions for the parent table, and any associated views. */
	IF @fSysSecMgr = 1 
	BEGIN
		INSERT INTO #columnPermissions
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
			INSERT INTO #columnPermissions
			SELECT 
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
				AND syscolumns.name <> ''timestamp''
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @piParentTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @piParentTableID)
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Get permitted child view on the parent table. */
			SELECT @sParentRealSource = sysobjects.name
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND sysprotects.action = 193
				AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = @piParentTableID)

			INSERT INTO #columnPermissions
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
				AND syscolumns.name <> ''timestamp''
				AND sysobjects.name = @sParentRealSource
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
	END

	/* Create a temporary table of the column info for all columns used in the summary controls. */
	CREATE TABLE #columnInfo
	(
		columnID	integer,
		selectGranted	bit
	)

	/* Populate the temporary table with info for all columns used in the summary controls. */
	/* Create the select string for getting the column values. */
	DECLARE columnsCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnID, 
		ASRSysColumns.columnName, 
		ASRSysColumns.dataType
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence

	OPEN columnsCursor
	FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0

		/* Get the select permission on the column. */

		/* Check if the column is selectable directly from the table. */
		SELECT @fSelectGranted = granted
		FROM #columnPermissions
		WHERE tableViewName = @sParentTableName
			AND columnName = @sColumnName

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0
		IF @fSelectGranted = 1 
		BEGIN
			/* Column COULD be read directly from the parent table. */
			IF len(@psSelectSQL) > 0 
			BEGIN
				SET @psSelectSQL = @psSelectSQL + '', ''
			END

			IF @iColumnDataType = 11 /* Date */
			BEGIN
				 /* Date */
				SET @psSelectSQL = @psSelectSQL + ''convert(varchar(10), '' + @sParentTableName + ''.'' + @sColumnName + '', 101) AS ['' + convert(varchar(100), @iColumnID) + '']''
			END
			ELSE
			BEGIN
				 /* Non-date */
				SET @psSelectSQL = @psSelectSQL + @sParentTableName + ''.'' + @sColumnName + '' AS ['' + convert(varchar(100), @iColumnID) + '']''
			END

			/* Add the table to the array of tables/views to join if it has not already been added. */
			SELECT @iTempCount = COUNT(tableViewName)
			FROM #joinParents
			WHERE tableViewName = @sParentTableName

			IF @iTempCount = 0
			BEGIN
				INSERT INTO #joinParents (tableViewName) VALUES(@sParentTableName)
			END
		END
		ELSE	
		BEGIN
			/* Column could NOT be read directly from the parent table, so try the views. */
			SET @sSelectString = ''''

			DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName
			FROM #columnPermissions
			WHERE tableViewName <> @sParentTableName
				AND columnName = @sColumnName
				AND granted = 1

			OPEN viewCursor
			FETCH NEXT FROM viewCursor INTO @sViewName
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Column CAN be read from the view. */
				SET @fSelectGranted = 1 

				IF len(@sSelectString) = 0 SET @sSelectString = ''CASE''
	
				IF @iColumnDataType = 11 /* Date */
				BEGIN
					 /* Date */
					SET @sSelectString = @sSelectString +
						'' WHEN NOT '' + @sViewName + ''.'' + @sColumnName + '' IS NULL THEN convert(varchar(10), '' + @sViewName + ''.'' + @sColumnName + '', 101)''
				END
				ELSE
				BEGIN
					 /* Non-date */
					SET @sSelectString = @sSelectString +
						'' WHEN NOT '' + @sViewName + ''.'' + @sColumnName + '' IS NULL THEN '' + @sViewName + ''.'' + @sColumnName 
				END

				/* Add the view to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM #joinParents
				WHERE tableViewName = @sViewName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO #joinParents (tableViewName) VALUES(@sViewName)
				END

				FETCH NEXT FROM viewCursor INTO @sViewName
			END
			CLOSE viewCursor
			DEALLOCATE viewCursor

			IF len(@sSelectString) > 0
			BEGIN
				SET @sSelectString = @sSelectString +
					'' ELSE NULL END AS ['' + convert(varchar(100), @iColumnID) + '']''

				IF len(@psSelectSQL) > 0 
				BEGIN
					SET @psSelectSQL = @psSelectSQL + '', ''
				END
				SET @psSelectSQL = @psSelectSQL + @sSelectString		
			END
		END

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0

		INSERT INTO #columnInfo (columnID, selectGranted)
			VALUES (@iColumnID, @fSelectGranted)

		FETCH NEXT FROM columnsCursor INTO @iColumnID, @sColumnName, @iColumnDataType
	END
	CLOSE columnsCursor
	DEALLOCATE columnsCursor

	IF len(@psSelectSQL) > 0 
	BEGIN
		SET @psSelectSQL = ''SELECT '' + @psSelectSQL 

		SELECT @iTempCount = COUNT(tableViewName)
		FROM #joinParents

		IF @iTempCount = 1 
		BEGIN
			SELECT TOP 1 @sRootTable = tableViewName
			FROM #joinParents
		END
		ELSE
		BEGIN
			SET @sRootTable = @sParentTableName
		END

		SET @psSelectSQL = @psSelectSQL + '' FROM '' + @sRootTable

		/* Add the join code. */
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName
		FROM #joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @sTableViewName <> @sRootTable
			BEGIN
				SET @psSelectSQL = @psSelectSQL + 
					'' LEFT OUTER JOIN '' + @sTableViewName + '' ON '' + @sRootTable + ''.ID'' + '' = '' + @sTableViewName + ''.ID''
			END

			FETCH NEXT FROM joinCursor INTO @sTableViewName
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		SET @psSelectSQL = @psSelectSQL + '' WHERE '' + @sRootTable + ''.id = '' + convert(varchar(8000), @piParentRecordID)
		
	END

	/* Drop temporary tables no longer required. */
	DROP TABLE #joinParents
	DROP TABLE #columnPermissions

	SELECT DISTINCT ASRSysSummaryFields.sequence, 
	    	ASRSysSummaryFields.startOfGroup, 
		ASRSysColumns.columnName, 
		ASRSysColumns.columnID, 
		ASRSysColumns.dataType, 
		ASRSysColumns.size, 
		ASRSysColumns.decimals, 
		ASRSysColumns.controlType, 
		ASRSysColumns.alignment
	FROM ASRSysSummaryFields 
	INNER JOIN ASRSysColumns 
		ON ASRSysSummaryFields.parentColumnID = ASRSysColumns.columnID
	WHERE ASRSysSummaryFields.historyTableID = @piHistoryTableID
		AND ASRSysColumns.tableID = @piParentTableID 
	ORDER BY ASRSysSummaryFields.sequence

END')

/* --------------------------------------------- */
/* Drop and recreate sp_ASRIntGetHistoryMainMenu */
/* --------------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetHistoryMainMenu]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetHistoryMainMenu]

EXEC('CREATE PROCEDURE sp_ASRIntGetHistoryMainMenu
AS
BEGIN
	/* Return a recordset of information that can be used to format the History menu for the current user. 
	The recordset contains a row for each parent screen and history table in the HR Pro database. 
	The following information is given :
		parentScreenID		ID of the parent screen
		childTableID		ID of the child table
		childTableName		Name of the child table
		childTableScreenCount	Number of screens associated with the child table
		childTableScreenID	ID of the screen associated with the child table 
		childTableScreenName	Name of the screen associated with the child table

	If childTableScreenCount = 1 then the child table just requires a tool on the Database menu that calls up the given screen.
	Else, the child table requires a tool on the Database menu that calls up a sub-band of the collection of screens available for the child table. */

	DECLARE @iCurrentRoleID		integer,
		@fSysSecMgr			bit,
		@iChildTableID			integer,
		@fTableReadable		bit,
		@iTempCount			integer
	/* Get the current user''s role name and ID. */
	SELECT @iCurrentRoleID = a.uid
	FROM sysusers a
	INNER JOIN sysusers b 
		ON a.uid = b.gid
	WHERE b.name = CURRENT_USER

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iCurrentRoleID
		AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
		OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

	/* Create a temporary table to hold our resultset. */
	CREATE TABLE #historyScreens
	(
		parentScreenID		integer,
		childTableID		integer,
		childTableName		sysname,
		childTableScreenCount	integer,
		childTableScreenID	integer,
		childTableScreenName	sysname
	)

	/* Create a temporary table to hold the permitted child views. This speeds things up later. */
	IF @fSysSecMgr = 0
	BEGIN
		CREATE TABLE #permittedChildViews
		(
			childViewName	sysname
		)

		INSERT INTO #permittedChildViews (childViewName)
			(SELECT sysobjects.name
				FROM syscolumns
				INNER JOIN sysprotects ON (syscolumns.id = sysprotects.id
					AND sysprotects.action = 193 
					AND sysprotects.uid = @iCurrentRoleID
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				WHERE sysobjects.name like ''ASRSysChildView_%''
					AND syscolumns.name = ''timestamp''
					AND ((sysprotects.protectType = 205) OR (sysprotects.protectType = 204)))
	END

	/* Loop through the child screen tables getting more information for each one. */
	DECLARE screensCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT childScreens.tableID
		FROM ASRSysScreens parentScreens
		INNER JOIN ASRSysHistoryScreens ON parentScreens.screenID = ASRSysHistoryScreens.parentScreenID
		INNER JOIN ASRSysScreens childScreens ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID
		INNER JOIN ASRSysTables ON childScreens.tableID = ASRSysTables.tableID
		WHERE childScreens.quickEntry = 0
	OPEN screensCursor
	FETCH NEXT FROM screensCursor INTO @iChildTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		/* Determine if the current user has SELECT permission on the child table. */
		SET @fTableReadable = @fSysSecMgr

		IF @fTableReadable = 0
		BEGIN
			/* The user is not a system or security manager so read the table permissions. */
			SELECT @iTempCount = COUNT(*) 
			FROM ASRSysChildViews 
			WHERE ASRSysChildViews.tableID = @iChildTableID
				AND ''ASRSysChildView_'' + convert(sysname, ASRSysChildViews.childViewID) IN (SELECT childViewName FROM #permittedChildViews)
			
			IF @iTempCount > 0 SET @fTableReadable = 1
		END

		IF @fTableReadable = 1
		BEGIN
			/* The table is readable so add it to the array of history screens. */
			SELECT @iTempCount = COUNT(*)
			FROM ASRSysScreens parentScreens
			INNER JOIN ASRSysHistoryScreens ON parentScreens.screenID = ASRSysHistoryScreens.parentScreenID
			INNER JOIN ASRSysScreens childScreens ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID
			INNER JOIN ASRSysTables ON childScreens.tableID = ASRSysTables.tableID
			WHERE childScreens.tableID = @iChildTableID
				AND childScreens.quickEntry = 0

		   	 INSERT INTO #historyScreens (
				parentScreenID,
				childTableID,
				childTableName,
				childTableScreenCount,
				childTableScreenID,
				childTableScreenName)
			(SELECT parentScreens.screenID, 
				childScreens.tableID,
				ASRSysTables.tableName,
				(
					SELECT COUNT(*) 
					FROM ASRSysScreens 
					WHERE ASRSysScreens.TABLEID = childScreens.tableID 
						AND ASRSysScreens.quickEntry = 0
				),
				childScreens.screenID,
				childScreens.name
				FROM ASRSysScreens parentScreens
				INNER JOIN ASRSysHistoryScreens ON parentScreens.screenID = ASRSysHistoryScreens.parentScreenID
				INNER JOIN ASRSysScreens childScreens ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID
				INNER JOIN ASRSysTables ON childScreens.tableID = ASRSysTables.tableID
				WHERE childScreens.tableID = @iChildTableID
					AND childScreens.quickEntry = 0)
		END

		FETCH NEXT FROM screensCursor INTO @IChildTableID
	END
	CLOSE screensCursor
	DEALLOCATE screensCursor

	/* Return the resultset. */
	SELECT *
	FROM #historyScreens 
	ORDER BY parentScreenID, 
		childTableName DESC,
		childTableScreenName DESC
END')



/* --------------------------------- */
/* Create IntranetPicturePath column */
/* --------------------------------- */

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'IntranetPicturePath'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [IntranetPicturePath] [varchar] (2000) NULL 
END


/* ------------------------------------------- */
/* Drop and recreate sp_ASRIntGetConfiguration */
/* ------------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetConfiguration]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetConfiguration]

EXEC('CREATE PROCEDURE sp_ASRIntGetConfiguration (
	@psPicturePath varchar(8000) OUTPUT
)
AS
BEGIN
	/* Return the required configuation parameter. */
	SELECT @psPicturePath = IntranetPicturePath
	FROM ASRSysConfig

	IF @psPicturePath IS NULL SET @psPicturePath = ''''
END')


/* ------------------------------------ */
/* Drop and recreate sp_ASRIntGetRecord */
/* ------------------------------------ */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetRecord]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetRecord]

EXEC('CREATE PROCEDURE sp_ASRIntGetRecord (
	@piRecordID		integer OUTPUT,
	@piRecordCount	integer OUTPUT,
	@piRecordPosition	integer OUTPUT,
	@psSelectSQL		varchar(8000),
	@psFromSQL 		varchar(8000),
	@psFilterSQL 		varchar(8000),
	@psOrderSQL 		varchar(8000),
	@psRealSource 	varchar(8000),
	@psAction	 	varchar(100),
	@piParentTableID	integer,
	@piParentRecordID	integer
)
AS
BEGIN
	DECLARE @iRecordID 			integer, 
		@iRecordCount 			integer,
		@iRecordPosition 		integer,
		@sGetCommand		varchar(8000),
		@sCommand			nvarchar(4000),
		@sParamDefinition		nvarchar(500),
		@sSubCommand		nvarchar(4000),
		@sSubParamDefinition		nvarchar(500),
		@sPositionCommand		nvarchar(4000),
		@sPositionParamDefinition	nvarchar(500),
		@sMoveCommand		varchar(8000),
		@sReverseOrderSQL		varchar(8000),
		@sRelevantOrderSQL		varchar(8000),
		@sRemainingSQL		varchar(8000),
		@iCharIndex			integer,
		@iLastCharIndex		integer,
		@sDESCstring			varchar(5),
		@fPositionKnown		bit,
		@sPreviousWhere		varchar(8000),
		@sOrderItem			varchar(8000),
		@sOrderColumn			varchar(8000),
		@sOrderTable			varchar(8000),
		@iDotIndex 			integer,
		@iDataType			integer,
		@fBitValue			bit,
		@sVarCharValue		varchar(8000),
		@iIntValue			integer,
		@dblNumValue			float,
		@dtDateValue			datetime,
		@sTempTableName		sysname,
		@sTempTablePrefix		sysname,
		@iLoop				integer,
		@iSpaceIndex 			integer,
		@fDescending			integer,
		@iOriginalLength		integer

	SET @fPositionKnown = 0
	SET @sDESCstring = '' DESC''
	SET @iRecordID = @piRecordID

	IF (@psAction = ''LOAD'') AND (@piRecordID = 0) SET @psAction = ''MOVEFIRST''

	/* Create the reverse order SQL if required. */
	SET @sReverseOrderSQL = ''''
	IF (@psAction = ''MOVELAST'') OR (@psAction = ''MOVEPREVIOUS'')
	BEGIN
		SET @sRemainingSQL = @psOrderSQL

		SET @iLastCharIndex = 0
		SET @iCharIndex = CHARINDEX('', '', @psOrderSQL)
		WHILE @iCharIndex > 0 
		BEGIN
			IF UPPER(SUBSTRING(@psOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@psOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + '', ''
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@psOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + '', ''
			END

			SET @iLastCharIndex = @iCharIndex
			SET @iCharIndex = CHARINDEX('', '', @psOrderSQL, @iLastCharIndex + 1)
	
			SET @sRemainingSQL = SUBSTRING(@psOrderSQL, @iLastCharIndex + 1, LEN(@psOrderSQL) - @iLastCharIndex)
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring
	END

	/* Get the record count of the required recordset. */	
	SET @sCommand = ''SELECT @recordCount = COUNT(id)'' +
		'' FROM '' + @psRealSource

	IF @piParentTableID > 0
	BEGIN
		SET @sCommand = @sCommand +
			'' WHERE '' + @psRealSource + ''.id_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID)
	END
	SET @sParamDefinition = N''@recordCount integer OUTPUT''
	EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT

	SET @piRecordCount = @iRecordCount

	/* Get the required record ID and record position values if we''re moving to the first or last records. */
	IF (@psAction = ''MOVEFIRST'') OR (@psAction = ''MOVELAST'')
	BEGIN
		SET @fPositionKnown = 1

		SET @sCommand = ''SELECT TOP 1 @recordID = '' + @psRealSource + ''.id'' +
			'' FROM '' + @psFromSQL

		IF @piParentTableID > 0
		BEGIN
			SET @sCommand = @sCommand +
				'' WHERE '' + @psRealSource + ''.id_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID)
		END

		SET @sCommand = @sCommand +
			'' ORDER BY '' + 
			CASE 
				WHEN @psAction = ''MOVEFIRST'' THEN @psOrderSQL
				ELSE @sReverseOrderSQL
			END
		SET @sParamDefinition = N''@recordID integer OUTPUT''
		EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordID OUTPUT

		IF @iRecordID IS NULL 
		BEGIN
			SET @iRecordID = 0
		END

		SET @iRecordPosition = 
			CASE
				WHEN (@psAction = ''MOVEFIRST'') AND (@iRecordCount > 0) THEN 1
				ELSE @iRecordCount
			END
	END
	
	/* Get the required record ID and record position values if we''re moving to the next or previous records. */
	IF (@psAction = ''MOVENEXT'') OR (@psAction = ''MOVEPREVIOUS'')
	BEGIN
		/* Create a temporary table to hold the required record ID. 
		We do this using a temporary table as the @sMoveCommand string may get too long to use sp_executeSQL (the parameters of which must be nvarchar type and hence a maximum of 4000 characters). */
		SET @iLoop = 1
		SET @sTempTablePrefix = ''##ASRSysTempIntMove_''
		SET @sTempTableName = @sTempTablePrefix + CONVERT(varchar(100), @iLoop)
		WHILE EXISTS (SELECT * FROM tempdb..sysobjects WHERE name = @sTempTableName AND xType = ''U'')
		BEGIN
			SET @iLoop = @iLoop + 1
			SET @sTempTableName = @sTempTablePrefix + CONVERT(varchar(100), @iLoop)
		END
		EXECUTE (''CREATE TABLE '' + @sTempTableName + '' (recordID INT)'')

		SET @sMoveCommand = ''INSERT INTO '' + @sTempTableName + '' SELECT TOP 1 '' + @psRealSource + ''.id'' +
			'' FROM '' + @psFromSQL + 
			'' WHERE ''

		IF @piParentTableID > 0
		BEGIN
			SET @sMoveCommand = @sMoveCommand +
				@psRealSource + ''.id_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID) + '' AND ''
		END

		SET @sRelevantOrderSQL = CASE WHEN @psAction = ''MOVENEXT'' THEN @psOrderSQL ELSE @sReverseOrderSQL END
		SET @sPreviousWhere = ''''

		/* Get the order column values for the current record. */
		SET @iLastCharIndex = 0
		SET @iCharIndex = CHARINDEX('', '', @sRelevantOrderSQL)
		WHILE @iCharIndex > 0 
		BEGIN
			SET @fDescending = 
				CASE
					WHEN UPPER(SUBSTRING(@sRelevantOrderSQL, @iCharIndex - LEN(@sDESCstring), len(@sDESCstring))) = @sDESCstring THEN 1
					ELSE 0
				END

			SET @sOrderItem = SUBSTRING(@sRelevantOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1- (@fDescending * LEN(@sDESCstring)) - @iLastCharIndex)
			SET @iDotIndex = CHARINDEX(''.'', @sOrderItem)

			SET @sOrderTable = LTRIM(LEFT(@sOrderItem, @iDotIndex - 1))
			SET @iSpaceIndex = CHARINDEX('' '', REVERSE(@sOrderTable))
			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderTable = SUBSTRING(@sOrderTable, LEN(@sOrderTable) - @iSpaceIndex + 2, @iSpaceIndex - 1)
			END

			SET @sOrderColumn = RTRIM(SUBSTRING(@sOrderItem, @iDotIndex + 1, LEN(@sOrderItem) - @iDotIndex))
			SET @iSpaceIndex = CHARINDEX('' '', @sOrderColumn)
			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderColumn = SUBSTRING(@sOrderColumn, 1, @iSpaceIndex - 1)
			END

			/* Get the data type of the order. */
			SELECT @iDataType = xtype
			FROM syscolumns
			WHERE name = @sOrderColumn
				AND id = (SELECT id FROM sysobjects WHERE name = @sOrderTable)

			IF @iDataType = 104	/* bit */
			BEGIN
				SET @sSubCommand = ''SELECT @fValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@fValue bit OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @fBitValue OUTPUT

				IF @fBitValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + '' OR (''  + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @fBitValue) + '')''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @fBitValue) + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @fBitValue) + '')''		
					END
					ELSE
					BEGIN
						IF @fDescending = 0
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @fBitValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @fBitValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @fBitValue) + '')''		
					END
				END
			END

			IF @iDataType = 167	/* varchar */
			BEGIN
				SET @sSubCommand = ''SELECT @sValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@sValue varchar(8000) OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @sVarCharValue OUTPUT

				IF @sVarCharValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''('' + @sOrderItem + '' > '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''')  + '''''')''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''(('' + @sOrderItem + '' < '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''')  + '''''') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '''''' +REPLACE(@sVarCharValue, '''''''', '''''''''''') + ''''''))''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '''''' +REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''')''			
					END
				END
			END

			IF @iDataType = 56	/* integer */
			BEGIN
				SET @sSubCommand = ''SELECT @iValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@iValue integer OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @iIntValue OUTPUT

				IF @iIntValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @iIntValue)  + '')''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @iIntValue)  + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @iIntValue) + '')''		
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @iIntValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @iIntValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @iIntValue) + '')''			
					END
				END
			END

			IF @iDataType = 108	/* numeric */
			BEGIN
				SET @sSubCommand = ''SELECT @dblValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@dblValue float OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @dblNumValue OUTPUT

				IF @dblNumValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @dblNumValue)  + '')''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @dblNumValue)  + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @dblNumValue) + '')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @dblNumValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @dblNumValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @dblNumValue) + '')''			
					END
				END
			END

			IF @iDataType = 61	/* datetime */
			BEGIN
				SET @sSubCommand = ''SELECT @dtValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@dtValue datetime OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @dtDateValue OUTPUT

				IF @dtDateValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 0 SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN				
							SET @sMoveCommand = @sMoveCommand + ''('' + @sOrderItem + '' > '''''' + convert(varchar(8000), @dtDateValue, 121)  + '''''')''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + ''(('' + @sOrderItem + '' < '''''' + convert(varchar(8000), @dtDateValue, 121)  + '''''') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN				
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '''''' + convert(varchar(8000), @dtDateValue, 121) + ''''''))''
						END
						ELSE
						BEGIN
							SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''')''			
					END
				END
			END
	
			SET @iLastCharIndex = @iCharIndex
			SET @iCharIndex = CHARINDEX('', '', @sRelevantOrderSQL, @iLastCharIndex + 1)
			SET @sRemainingSQL = SUBSTRING(@sRelevantOrderSQL, @iLastCharIndex + 2, len(@sRelevantOrderSQL) - @iLastCharIndex)
		END

		/* Add on the ID condition. */
		IF (@psAction = ''MOVENEXT'')
		BEGIN
			IF LEN(@sPreviousWhere) = 0
			BEGIN
				SET @sMoveCommand = @sMoveCommand + ''('' + @psRealSource + ''.id > '' + convert(varchar(8000), @iRecordID)  + '')''
			END
			ELSE
			BEGIN
				SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @psRealSource + ''.id > '' + convert(varchar(8000), @iRecordID) + ''))''
			END
		END
		ELSE
		BEGIN
			IF LEN(@sPreviousWhere) = 0
			BEGIN
				SET @sMoveCommand = @sMoveCommand + ''('' + @psRealSource + ''.id < '' + convert(varchar(8000), @iRecordID)  + '')''
			END
			ELSE
			BEGIN
				SET @sMoveCommand = @sMoveCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @psRealSource + ''.id < '' + convert(varchar(8000), @iRecordID) + ''))''
			END
		END

		SET @sMoveCommand = @sMoveCommand +
			'' ORDER BY '' + @sRelevantOrderSQL
		EXECUTE (@sMoveCommand)

		/* Get the result from the temporary table. */
		SET @sCommand = ''SELECT @recordID = recordID FROM '' + @sTempTableName 
		SET @sParamDefinition = N''@recordID integer OUTPUT''
		EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordID OUTPUT

		/* Drop the temporary table. */
		EXEC (''DROP TABLE '' + @sTempTableName)

		IF @iRecordID IS NULL 
		BEGIN
			SET @iRecordID = 0
		END
	END

	IF @fPositionKnown = 0
	BEGIN
		/* Calculate the current record''s position. */
		SET @sPositionCommand = ''SELECT @recordPosition = COUNT('' + @psRealSource + ''.id)'' +
			'' FROM '' + @psFromSQL + 
			'' WHERE ''

		IF @piParentTableID > 0
		BEGIN
			SET @sPositionCommand = @sPositionCommand +
				''('' + @psRealSource + ''.id_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID) + '') AND ''
		END
		SET @sPositionCommand = @sPositionCommand + ''('' 

		SET @sPreviousWhere = ''''
		SET @iOriginalLength = LEN(@sPositionCommand)

		/* Get the order column values for the current record. */
		SET @iLastCharIndex = 0
		SET @iCharIndex = CHARINDEX('', '', @psOrderSQL)
		WHILE @iCharIndex > 0 
		BEGIN
			SET @fDescending = CASE
					WHEN UPPER(SUBSTRING(@psOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring THEN 1
					ELSE 0
				END

			SET @sOrderItem = SUBSTRING(@psOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - (@fDescending * LEN(@sDESCstring)) - @iLastCharIndex)
			SET @iDotIndex = CHARINDEX(''.'', @sOrderItem)

			SET @sOrderTable = LTRIM(LEFT(@sOrderItem, @iDotIndex - 1))
			SET @iSpaceIndex = CHARINDEX('' '', REVERSE(@sOrderTable))
			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderTable = SUBSTRING(@sOrderTable, LEN(@sOrderTable) - @iSpaceIndex + 2, @iSpaceIndex - 1)
			END
				SET @sOrderColumn = RTRIM(SUBSTRING(@sOrderItem, @iDotIndex + 1, LEN(@sOrderItem) - @iDotIndex))

			SET @iSpaceIndex = CHARINDEX('' '', @sOrderColumn)
			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderColumn = SUBSTRING(@sOrderColumn, 1, @iSpaceIndex - 1)
			END

			/* Get the data type of the order. */
			SELECT @iDataType = xtype
			FROM syscolumns
			WHERE name = @sOrderColumn
				AND id = (SELECT id FROM sysobjects WHERE name = @sOrderTable)


			IF @iDataType = 104	/* bit */
			BEGIN
				SET @sSubCommand = ''SELECT @fValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@fValue bit OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @fBitValue OUTPUT

				IF @fBitValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
							IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
							SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
							IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + '' OR (''  + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
							SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @fBitValue) + '')''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @fBitValue) + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @fBitValue) + '')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @fBitValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @fBitValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @fBitValue) + '')''			
					END
				END
			END

			IF @iDataType = 167	/* varchar */
			BEGIN
				SET @sSubCommand = ''SELECT @sValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@sValue varchar(8000) OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @sVarCharValue OUTPUT

				IF @sVarCharValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''('' + @sOrderItem + '' > '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''')  + '''''')''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''(('' + @sOrderItem + '' < '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''')  + '''''') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '''''' +REPLACE(@sVarCharValue, '''''''', '''''''''''') + ''''''))''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '''''' +REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '''''' + REPLACE(@sVarCharValue, '''''''', '''''''''''') + '''''')''			
					END
				END
			END

			IF @iDataType = 56	/* integer */
			BEGIN
				SET @sSubCommand = ''SELECT @iValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@iValue integer OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @iIntValue OUTPUT

				IF @iIntValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @iIntValue)  + '')''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @iIntValue)  + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @iIntValue) + '')''		
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @iIntValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @iIntValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @iIntValue) + '')''			
					END
				END
			END

			IF @iDataType = 108	/* numeric */
			BEGIN
				SET @sSubCommand = ''SELECT @dblValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@dblValue float OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @dblNumValue OUTPUT

				IF @dblNumValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''('' + @sOrderItem + '' > '' + convert(varchar(8000), @dblNumValue)  + '')''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''(('' + @sOrderItem + '' < '' + convert(varchar(8000), @dblNumValue)  + '') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '' + convert(varchar(8000), @dblNumValue) + '')''		
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '' + convert(varchar(8000), @dblNumValue) + ''))''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '' + convert(varchar(8000), @dblNumValue) + '') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '' + convert(varchar(8000), @dblNumValue) + '')''			
					END
				END
			END

			IF @iDataType = 61	/* datetime */
			BEGIN
				SET @sSubCommand = ''SELECT @dtValue = '' + @sOrderItem +
					'' FROM '' + @psFromSQL +
					'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N''@dtValue datetime OUTPUT''
				EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @dtDateValue OUTPUT

				IF @dtDateValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + ''(NOT '' + @sOrderItem + '' IS NULL)''
						SET @sPreviousWhere = ''('' + @sOrderItem + '' IS NULL)''
					END
					ELSE
					BEGIN
						IF @fDescending = 1 SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (NOT'' + @sOrderItem + '' IS NULL))''
						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' IS NULL)''
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''('' + @sOrderItem + '' > '''''' + convert(varchar(8000), @dtDateValue, 121)  + '''''')''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + ''(('' + @sOrderItem + '' < '''''' + convert(varchar(8000), @dtDateValue, 121)  + '''''') OR (''  + @sOrderItem + '' IS NULL))''
						END

						SET @sPreviousWhere = ''('' + @sOrderItem + '' = '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''')''			
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND ('' + @sOrderItem + '' > '''''' + convert(varchar(8000), @dtDateValue, 121) + ''''''))''
						END
						ELSE
						BEGIN
							SET @sPositionCommand = @sPositionCommand + '' OR ('' + @sPreviousWhere + '' AND (('' + @sOrderItem + '' < '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''') OR (''  + @sOrderItem + '' IS NULL)))''
						END

						SET @sPreviousWhere = @sPreviousWhere + '' AND ('' + @sOrderItem + '' = '''''' + convert(varchar(8000), @dtDateValue, 121) + '''''')''			
					END
				END
			END

			SET @iLastCharIndex = @iCharIndex
			SET @iCharIndex = CHARINDEX('', '', @psOrderSQL, @iLastCharIndex + 1)
			SET @sRemainingSQL = SUBSTRING(@psOrderSQL, @iLastCharIndex + 2, len(@psOrderSQL) - @iLastCharIndex)
		END

		/* Add on the ID condition. */
		IF LEN(@sPreviousWhere) = 0
		BEGIN
			SET @sPositionCommand = @sPositionCommand + ''(('' + @psRealSource + ''.id < '' + convert(varchar(8000), @iRecordID)  + '') OR (''  + @psRealSource + ''.id IS NULL))''
		END
		ELSE
		BEGIN
			SET @sPositionCommand = @sPositionCommand + 
				CASE
					WHEN @iOriginalLength = LEN(@sPositionCommand) THEN ''''
					ELSE '' OR ''
				END + 
				''('' + @sPreviousWhere + '' AND (('' + @psRealSource + ''.id < '' + convert(varchar(8000), @iRecordID) + '') OR (''  + @psRealSource + ''.id IS NULL)))''

		END
		SET @sPositionCommand = @sPositionCommand + '')'' 

		SET @sPositionParamDefinition = N''@recordPosition integer OUTPUT''
		EXEC sp_executesql @sPositionCommand,  @sPositionParamDefinition, @iRecordPosition OUTPUT

		SET @iRecordPosition = @iRecordPosition + 1
	END

	/* Set the output parameter values. */
	SET @piRecordID = @iRecordID
	SET @piRecordPosition = @iRecordPosition

	/* Return the required record. */
	SET @sGetCommand = ''SELECT '' + @psSelectSQL +
		'' FROM '' + @psFromSQL +
		'' WHERE '' + @psRealSource + ''.id = '' + convert(varchar(100), @iRecordID)
	execute(@sGetCommand)
END')


/* ------------------------------------------ */
/* Drop and recreate sp_ASRIntGetFindRecords2 */
/* ------------------------------------------ */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetFindRecords2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetFindRecords2]


EXEC('CREATE PROCEDURE sp_ASRIntGetFindRecords2 (
	@pfError 		bit 		OUTPUT, 
	@pfSomeSelectable 	bit 		OUTPUT, 
	@pfSomeNotSelectable 	bit 		OUTPUT, 
	@psRealSource		varchar(8000)	OUTPUT,
	@pfInsertGranted	bit		OUTPUT,
	@pfDeleteGranted	bit		OUTPUT,
	@piTableID 		integer, 
	@piViewID 		integer, 
	@piOrderID 		integer, 
	@piParentTableID	integer,
	@piParentRecordID	integer)
AS
BEGIN
	/* Return a recordset of the find records for the current user, given the table/view and order IDs.
		@pfError = 1 if errors occured in getting the find records. Else 0.
		@pfSomeSelectable = 1 if some find columns were selectable. Else 0.
		@pfSomeNotSelectable = 1 if some find columns were NOT selectable. Else 0.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@piParentTableID = the ID of the parent table.
		@piParentRecordID = the ID of the associated record in the parent table.
	*/
	DECLARE @iUserGroupID	integer,
		@iTableType		integer,
		@sTableName		sysname,
		@fSysSecMgr		bit,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 	integer,
		@iColumnTableID 	integer,
		@iColumnID 		integer,
		@sColumnName 	sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 	varchar(10),
		@fSelectGranted 	bit,
		@sSelectSQL		varchar(8000),
		@sOrderSQL 		varchar(8000),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString		varchar(8000),
		@sViewName 		varchar(8000),
		@sExecString		nvarchar(4000),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iDataType 		integer,
		@iTempAction		integer

	/* Initialise variables. */
	SET @pfError = 0
	SET @pfSomeSelectable = 0
	SET @pfSomeNotSelectable = 0

	SET @sRealSource = ''''
	SET @sSelectSQL = ''''
	SET @sOrderSQL = ''''
	SET @fSelectDenied = 0
	SET @sExecString = ''''

	/* Get the current user''s group ID. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = CURRENT_USER

	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID

	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1
		RETURN
	END

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER'' OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
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
			SET @sRealSource = @sTableName
		END 
	END
	ELSE
	BEGIN
		IF @fSysSecMgr = 1 
		BEGIN
			/* RealSource is the child view on the table which is derived from full access on the table''s parents. */	
			exec sp_ASRIntGetFullAccessChildView @piTableID, @iChildViewID OUTPUT
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = ''ASRSysChildView_'' + convert(varchar(1000), @iChildViewID)
			END
		END
		ELSE
		BEGIN
			/* Get appropriate child view if required. */
			SELECT @sRealSource = sysobjects.name
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND sysprotects.action = 193
				AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = @piTableID)
		END
	END
	SET @psRealSource = @sRealSource
	
	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1
		RETURN
	END

	/* Create a temporary table to hold the tables/views that need to be joined. */
	CREATE TABLE #joinParents
	(
		tableViewName	sysname,
		tableID		integer
	)	

	/* Create a temporary table of the ''select'' column permissions for all tables/views used in the order. */
	CREATE TABLE #columnPermissions
	(
		tableID		integer,
		tableViewName	sysname,
		columnName	sysname,
		selectGranted	bit		
	)

	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
	WHERE ASRSysOrderItems.orderID = @piOrderID

	OPEN tablesCursor
	FETCH NEXT FROM tablesCursor INTO @iTempTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO #columnPermissions
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
				AND syscolumns.name <> ''timestamp''
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO #columnPermissions
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
				AND syscolumns.name <> ''timestamp''
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.tableID,
		ASRSysOrderItems.columnID, 
		ASRSysColumns.columnName,
	    	ASRSysTables.tableName,
		ASRSysOrderItems.ascending,
		ASRSysOrderItems.type,
		ASRSysColumns.dataType
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
	ORDER BY ASRSysOrderItems.sequence

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1
		RETURN
	END

	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM #columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0

			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = ''F''
				BEGIN
					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						CASE 
							WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sRealSource + ''.'' + @sColumnName + '',103) AS['' + @sColumnName + '']''
							ELSE @sRealSource + ''.'' + @sColumnName
						END
				END
				ELSE
				BEGIN
					/* Order column. */
					SET @sOrderSQL = @sOrderSQL + 
						CASE WHEN len(@sOrderSQL) > 0 THEN '','' ELSE '''' END + 
						@sRealSource + ''.'' + @sColumnName + 
						CASE WHEN @fAscending = 0 THEN '' DESC'' ELSE '''' END				
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1
			END	
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM #columnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = ''F''
				BEGIN
					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						CASE 
							WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sColumnTableName + ''.'' + @sColumnName + '',103) AS['' + @sColumnName + '']''
							ELSE @sColumnTableName + ''.'' + @sColumnName
						END
				END
				ELSE
				BEGIN
					/* Order column. */
					SET @sOrderSQL = @sOrderSQL + 
						CASE WHEN len(@sOrderSQL) > 0 THEN '','' ELSE '''' END + 
						@sColumnTableName + ''.'' + @sColumnName + 
						CASE WHEN @fAscending = 0 THEN '' DESC'' ELSE '''' END				
				END

				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM #joinParents
				WHERE tableViewName = @sColumnTableName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO #joinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID)
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = ''''

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM #columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = ''CASE''

					SET @sSubString = @sSubString +
						'' WHEN NOT '' + @sViewName + ''.'' + @sColumnName + '' IS NULL THEN '' + @sViewName + ''.'' + @sColumnName 
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM #joinParents
					WHERE tableViewname = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO #joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +
						'' ELSE NULL END''

					IF @sType = ''F''
					BEGIN
						/* Find column. */
						SET @sSelectSQL = @sSelectSQL + 
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN '','' 
								ELSE '''' 
							END + 
							CASE 
								WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sSubString + '',103) AS['' + @sColumnName + '']''
								ELSE @sSubString
							END							
					END
					ELSE
					BEGIN
						/* Order column. */
						SET @sOrderSQL = @sOrderSQL + 
							CASE WHEN len(@sOrderSQL) > 0 THEN '','' ELSE '''' END + 
							@sSubString + 
							CASE WHEN @fAscending = 0 THEN '' DESC'' ELSE '''' END				
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1
				END	
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Set the flags that show if no order columns could be selected, or if only some of them could be selected. */
	SET @pfSomeSelectable = CASE WHEN len(@sSelectSQL) > 0 THEN 1 ELSE 0 END
	SET @pfSomeNotSelectable = @fSelectDenied

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN '','' ELSE '''' END + 
		@sRealSource + ''.ID''

	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = @sSelectSQL + '','' + @sRealSource + ''.ID''

		SET @sExecString = ''SELECT '' + @sSelectSQL + 
			'' FROM '' + @sRealSource

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM #joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sExecString = @sExecString + 
				'' LEFT OUTER JOIN '' + @sTableViewName + '' ON '' + @sRealSource + ''.ID_'' + convert(varchar(100), @iJoinTableID) + '' = '' + @sTableViewName + ''.ID''

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		/* Add the filter code. */
		IF @piParentTableID > 0 
		BEGIN
			SET @sExecString = @sExecString + 
				'' WHERE '' + @sRealSource + ''.ID_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID)
		END

		/* Add the ORDER BY code to the find record selection string if required. */
		SET @sExecString = @sExecString + '' ORDER BY '' + @sOrderSQL
	END

	/* Drop temporary tables no longer required. */
	DROP TABLE #joinParents
	DROP TABLE #columnPermissions

	/* Check if the user has insert or delete permission on the table. */
	SET @pfInsertGranted = 0
	SET @pfDeleteGranted = 0

	IF LEN(@sRealSource) > 0
	BEGIN
		DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT sysprotects.action
			FROM sysprotects 
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.protectType <> 206
				AND ((sysprotects.action = 195) OR (sysprotects.action = 196))
				AND sysobjects.name = @sRealSource

		OPEN tableInfo_cursor
		FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iTempAction = 195
			BEGIN
				SET @pfInsertGranted = 1
			END
			ELSE
			BEGIN
				SET @pfDeleteGranted = 1	
			END
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
		END
		CLOSE tableInfo_cursor
		DEALLOCATE tableInfo_cursor
	END

	/* Return a recordset of the required columns in the required order from the given table/view. */
	IF (@pfSomeSelectable = 1)
	BEGIN
		EXEC (@sExecString)
	END
END')


/* ------------------------------------------------------ */
/* ------------------------------------------------------ */
/* ------------------------------------------------------ */
/* Note : AsrLicense.dll must be registered on the server */
/* ------------------------------------------------------ */
/* ------------------------------------------------------ */
/* ------------------------------------------------------ */


/* --------------------------------------- */
/* Drop and recreate sp_ASRIntCalcDefaults */
/* --------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntCalcDefaults]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntCalcDefaults]

EXEC('CREATE PROCEDURE sp_ASRIntCalcDefaults (
	@piRecordCount	integer OUTPUT,
	@psFromSQL 		varchar(8000),
	@psFilterSQL 		varchar(8000),
	@psRealSource 	varchar(8000),
	@piTableID		integer,
	@piParentTableID	integer,
	@piParentRecordID	integer,
	@psDefaultCalcColumns	varchar(8000)
)
AS
BEGIN
	DECLARE @iRecordCount 	integer,
		@sCommand		nvarchar(4000),
		@sParamDefinition	nvarchar(500),
		@sColumns		varchar(8000),
		@iID			integer,
		@iDataType		integer,
		@iSize			integer,
		@iDecimals		integer,
		@iDfltExprID		integer,
		@fOneColumnDone	bit,
		@iCount		integer,
		@fOK			bit,
		@iTableID		integer,
		@sCharResult 		varchar(8000),
		@dblNumericResult 	float,
		@iIntegerResult 		integer,
		@dtDateResult 		datetime,
		@fLogicResult 		bit,
		@sTempTableName	sysname,
		@sTemp 		sysname,
		@iLoop 		integer
		
	SET @fOneColumnDone = 0
	SET @fOK = 1

	/* Get the record count of the current recordset. */
	SET @sCommand = ''SELECT @recordCount = COUNT('' + @psRealSource + ''.id)'' +
		'' FROM '' + @psFromSQL

	IF @piParentTableID > 0
	BEGIN
		SET @sCommand = @sCommand +
			'' WHERE '' + @psRealSource + ''.id_'' + convert(varchar(100), @piParentTableID) + '' = '' + convert(varchar(100), @piParentRecordID)
	END
	SET @sParamDefinition = N''@recordCount integer OUTPUT''
	EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT
	SET @piRecordCount = @iRecordCount

	/* Get the default values for the given columns. */
	SET @sColumns = @psDefaultCalcColumns
	WHILE len(@sColumns) > 0
	BEGIN
		IF CHARINDEX('','', @sColumns) > 0
		BEGIN
			SET @iID = convert(integer, left(@sColumns, CHARINDEX('','', @sColumns) - 1))
			SET @sColumns = substring(@sColumns, CHARINDEX('','', @sColumns) + 1, len(@sColumns))
		END
		ELSE
		BEGIN
			SET @iID = convert(integer, @sColumns)
			SET @sColumns = ''''
		END

		IF @iID > 0 			
		BEGIN
			/* Get the data type and size of the column. */
			SELECT @iDataType = dataType, 
				@iSize = size, 
				@iDecimals = decimals,
				@iDfltExprID = dfltValueExprID
			FROM ASRSysColumns
			WHERE columnID = @iID

			/* Check the default expression stored procedure exists. */
			SET @sCommand = ''SELECT @count = COUNT(*)'' +
				'' FROM sysobjects'' +
				'' WHERE id = object_id(N''''sp_ASRDfltExpr_'' + convert(varchar(100), @iDfltExprID) + '''''')'' +
				'' AND OBJECTPROPERTY(id, N''''IsProcedure'''') = 1''
			SET @sParamDefinition = N''@count integer OUTPUT''
			EXEC sp_executesql @sCommand,  @sParamDefinition, @iCount OUTPUT

			IF @iCount > 0 
			BEGIN
				SET @sCommand = ''exec sp_ASRDfltExpr_'' + convert(varchar(100), @iDfltExprID) + '' @result output''
	
				SET @fOK = 0

				IF @iDataType = -7 /* Logic columns. */
				BEGIN
					SET @sParamDefinition = N''@result bit OUTPUT''
					SET @fOK = 1
				END
          
				IF @iDataType = 2 /* Numeric columns. */
				BEGIN
					SET @sParamDefinition = N''@result float OUTPUT''
					SET @fOK = 1
				END
          
				IF @iDataType = 4 /* Integer columns. */
				BEGIN
					SET @sParamDefinition = N''@result integer OUTPUT''
					SET @fOK = 1
				END
          
				IF @iDataType = 11 /* Date columns. */
				BEGIN
					SET @sParamDefinition = N''@result datetime OUTPUT''
					SET @fOK = 1
				END
          
				IF @iDataType = 12 /* Character columns. */
				BEGIN
					SET @sParamDefinition = N''@result varchar(8000) OUTPUT''
					SET @fOK = 1
				END
          
				IF @iDataType = -1 /* Working Pattern columns. */
				BEGIN
					SET @sParamDefinition = N''@result varchar(8000) OUTPUT''
					SET @fOK = 1
				END

				IF @fOK = 1
				BEGIN
 					/* Append the parent table ID parameters. */
					DECLARE parentsCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT parentID
						FROM ASRSysRelations
						WHERE childID = @piTableID
						ORDER BY parentID
					OPEN parentsCursor
					FETCH NEXT FROM parentsCursor INTO @iTableID
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iTableID = @piParentTableID
						BEGIN
							SET @sCommand = @sCommand + '', '' + convert(varchar(100), @piParentRecordID)
						END
						ELSE
						BEGIN
							SET @sCommand = @sCommand + '', 0''
						END

						FETCH NEXT FROM parentsCursor INTO @iTableID
					END
					CLOSE parentsCursor
					DEALLOCATE parentsCursor

					IF @iDataType = -7 /* Logic columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @fLogicResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] bit NULL)''
							EXEC sp_executesql @sCommand

							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue bit''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @fLogicResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] bit NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '' + convert(nvarchar(4000), @fLogicResult)
							EXEC sp_executesql @sCommand
						END
					END
          
					IF @iDataType = 2 /* Numeric columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @dblNumericResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] float NULL)''
							EXEC sp_executesql @sCommand

							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue float''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @dblNumericResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] float NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '' + convert(nvarchar(4000), @dblNumericResult)
							EXEC sp_executesql @sCommand
						END
					END
          
					IF @iDataType = 4 /* Integer columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @iIntegerResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END

								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] integer NULL)''
							EXEC sp_executesql @sCommand

							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue integer''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @iIntegerResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] integer NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '' + convert(nvarchar(4000), @iIntegerResult)
							EXEC sp_executesql @sCommand
						END
					END
          
					IF @iDataType = 11 /* Date columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @dtDateResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] datetime NULL)''
							EXEC sp_executesql @sCommand

							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue datetime''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @dtDateResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] datetime NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '''''' + convert(nvarchar(4000), @dtDateResult, 101) + ''''''''
							EXEC sp_executesql @sCommand
						END
					END
          	
					IF @iDataType = 12 /* Character columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] varchar(8000) NULL)''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue varchar(8000)''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] varchar(8000) NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '''''' + REPLACE(convert(nvarchar(4000), @sCharResult), '''''''', '''''''''''') + ''''''''
							EXEC sp_executesql @sCommand
						END
					END
          
					IF @iDataType = -1 /* Working Pattern columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult OUTPUT
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = ''''
							SET @iLoop = 1
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = ''tmpDefaultValues_'' + convert(varchar(100), @iLoop)

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1
								END
							END

							SET @sCommand = ''CREATE TABLE '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + ''] varchar(8000) NULL)''
							EXEC sp_executesql @sCommand

							SET @sCommand = ''INSERT INTO '' + @sTempTableName +'' (['' + convert(varchar(100), @iID) + '']) VALUES (@newValue)''
							SET @sParamDefinition = N''@newValue varchar(8000)''
							EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = ''ALTER TABLE '' + @sTempTableName +'' ADD ['' + convert(varchar(100), @iID) + ''] varchar(8000) NULL''
							EXEC sp_executesql @sCommand
							SET @sCommand = ''UPDATE '' + @sTempTableName +'' SET ['' + convert(varchar(100), @iID) + ''] = '''''' + REPLACE(convert(nvarchar(4000), @sCharResult), '''''''', '''''''''''') + ''''''''
							EXEC sp_executesql @sCommand
						END
					END

					SET @fOneColumnDone = 1
				END
			END
		END
	END

	IF @fOneColumnDone > 0
	BEGIN
		SET @sCommand = ''SELECT * FROM '' + @sTempTableName
		EXEC sp_executesql @sCommand

		SET @sCommand = ''DROP TABLE '' + @sTempTableName
		EXEC sp_executesql @sCommand
	END
END')


/* ---------------------------------------------- */
/* Drop and recreate sp_ASRIntGetScreenDefinition */
/* ---------------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetScreenDefinition]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetScreenDefinition]

EXEC('CREATE PROCEDURE sp_ASRIntGetScreenDefinition (
	@piScreenID 	integer,
	@piViewID	integer)
AS
BEGIN
	/* Return a recordset of the given screen''s definition and table permission info. */
	DECLARE @iTabCount 		integer,
		@sTabCaptions		varchar(8000),
		@sTabCaption		varchar(8000),
		@fSysSecMgr		bit,
		@fInsertGranted		bit,
		@fDeleteGranted	bit,
		@sRealSource		sysname,
		@iUserGroupID		integer,
		@iTableID		integer,
		@iTableType		integer,
		@sTableName		sysname,
		@sTempName		sysname,
		@iTempAction		integer,
		@iChildViewID 		integer
	
	/* Get the current user''s group id. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = CURRENT_USER

	/* Get the table type and name. */
	SELECT @iTableID = ASRSysScreens.tableID,
		@iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysScreens
	INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.screenID = @piScreenID

	/* Check if the current user is a System or Security manager. */
	SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
	OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

	/* Get the real source and insert/delete permissions for the table. */
	IF @fSysSecMgr = 1 
	BEGIN
		/* Permission must be granted for System or Security mangers. */
		SET @fInsertGranted = 1
		SET @fDeleteGranted = 1	

		/* Get the realSource of the table. */
		IF @iTableType <> 2 /* ie. top-level or lookup */
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
				SET @sRealSource = @sTableName
			END 
		END
		ELSE
		BEGIN
			/* RealSource is the child view on the table which is derived from full access on the table''s parents. */	
			exec sp_ASRIntGetFullAccessChildView @iTableID, @iChildViewID OUTPUT
			IF @iChildViewID > 0 
			BEGIN
				SET @sRealSource = ''ASRSysChildView_'' + convert(varchar(1000), @iChildViewID)
			END
		END
	END
	ELSE
	BEGIN
		/* Permission must be read from the database  for Non-System and Non-Security mangers. */
		SET @fInsertGranted = 0
		SET @fDeleteGranted = 0	

		IF @iTableType <> 2 /* ie. top-level or lookup */
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
				SET @sRealSource = @sTableName
			END 

			/* Get the insert/delete permissions for the realSource. */
			DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT sysprotects.action
				FROM sysprotects 
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				WHERE sysprotects.uid = @iUserGroupID
					AND sysprotects.protectType <> 206
					AND ((sysprotects.action = 195) OR (sysprotects.action = 196))
					AND sysobjects.name = @sRealSource

			OPEN tableInfo_cursor
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			WHILE (@@fetch_status = 0)
			BEGIN
				IF @iTempAction = 195
				BEGIN
					SET @fInsertGranted = 1
				END
				ELSE
				BEGIN
					SET @fDeleteGranted = 1	
				END
				FETCH NEXT FROM tableInfo_cursor INTO @iTempAction
			END
			CLOSE tableInfo_cursor
			DEALLOCATE tableInfo_cursor
		END
		ELSE
		BEGIN
			/* Get appropriate child view if required. */
			DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT sysobjects.name, sysprotects.action
				FROM sysprotects 
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				WHERE sysprotects.uid = @iUserGroupID
					AND sysprotects.protectType <> 206
					AND ((sysprotects.action = 195) OR (sysprotects.action = 196) OR (sysprotects.action = 193))
					AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = @iTableID)

			OPEN tableInfo_cursor
			FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sRealSource = @sTempName
			
				IF @iTempAction = 195
				BEGIN
					SET @fInsertGranted = 1
				END
				ELSE
				BEGIN
					IF @iTempAction = 196
					BEGIN
						SET @fDeleteGranted = 1	
					END
				END
				FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
			END
			CLOSE tableInfo_cursor
			DEALLOCATE tableInfo_cursor
		END
	END

	/* Get the tab page captions info. */
	SET @iTabCount = 0
	SET @sTabCaptions = ''''
	
	DECLARE captions_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT caption 
		FROM ASRSysPageCaptions
		WHERE screenID = @piScreenID
		ORDER BY pageIndexID

	OPEN captions_cursor
	FETCH NEXT FROM captions_cursor INTO @sTabCaption
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTabCount > 0 SET @sTabCaptions = @sTabCaptions + char(9) 

		SET @iTabCount = @iTabCount + 1
		SET @sTabCaptions = @sTabCaptions + @sTabCaption
			
		FETCH NEXT FROM captions_cursor INTO @sTabCaption
	END
	CLOSE captions_cursor
	DEALLOCATE captions_cursor

	SELECT @sTableName AS tableName,
		@sRealSource AS realSource,
		@fInsertGranted AS insertGranted,
		@fDeleteGranted AS deleteGranted,
		height,
		width,
		fontName,
		fontSize,
		fontBold,
		fontItalic,
		fontStrikethru,
		fontUnderline,
		@iTabCount AS tabCount,
		@sTabCaptions AS tabCaptions
	FROM ASRSysScreens
	WHERE screenID = @piScreenID
END')


/* ----------------------------- */
/* Absence Between 2 Dates Stuff */
/* ----------------------------- */

DELETE FROM ASRSysFunctions WHERE FunctionID = 47
INSERT INTO ASRSysFunctions VALUES (47,'Absence between two dates',2,0,'Absence','sp_ASRFn_Absencebetweentwodates',0,0)

DELETE FROM ASRSysFunctionParameters WHERE FunctionID = 47
INSERT INTO ASRSysFunctionParameters VALUES (47,1,4,'<Start Date>')
INSERT INTO ASRSysFunctionParameters VALUES (47,2,4,'<End Date>')
INSERT INTO ASRSysFunctionParameters VALUES (47,3,1,'<Absence Type>')


/* ---------------- */
/* Absence Duration */
/* ---------------- */

DELETE FROM ASRSysFunctionParameters WHERE FunctionID = 30 AND (parameterindex = 5 OR parameterindex = 6)
UPDATE ASRSysFunctions SET category = 'Absence' WHERE FunctionID = 30


/* --------------------- */
/* General Absence Stuff */
/* --------------------- */

DELETE FROM ASRSysModuleSetup WHERE (ModuleKey = 'MODULE_ABSENCE' AND ParameterKey = 'Param_FieldAbsenceRegion')
DELETE FROM ASRSysModuleSetup WHERE (ModuleKey = 'MODULE_ABSENCE' AND ParameterKey = 'Param_FieldAbsenceWorkingPattern')
DELETE FROM ASRSysModuleSetup WHERE (ModuleKey = 'MODULE_ABSENCE' AND ParameterKey = 'Param_FieldWorkingPattern')

/* ------------------------------ */
/* Working Days Between Two Dates */
/* ------------------------------ */

DELETE FROM ASRSysFunctionParameters WHERE FunctionID = 46 AND (parameterindex = 3 OR parameterindex = 4)

/* ------------------------------------------ */
/* Drop and recreate sp_ASRFn_AbsenceDuration */
/* ------------------------------------------ */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_AbsenceDuration]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_AbsenceDuration]

EXEC('CREATE PROCEDURE sp_ASRFn_AbsenceDuration (
	@pdblResult		float OUTPUT,						
	@pdtStartDate		datetime,						
	@psStartSession	varchar(255),						
	@pdtEndDate		datetime,					
	@psEndSession		varchar(255),					
	@iPersonnelID           	int)                   			

AS
BEGIN

/* Used to work out if we can hit child tables directly, or via childviews */
DECLARE @iUserGroupID			int
DECLARE @fSysSecMgr				bit

/* Personnel Table ID and name...used for static region/wp and for ID_xx purposes */
DECLARE @sPersonnelTable 			varchar(255)
DECLARE @iPersonnelTableID        		int

/* The Bank Holiday Region (Primary) Table which contains England, Scotland, Wales etc. */
DECLARE @iBHolRegionTableID			int		
DECLARE @sBHolRegionTableName		sysname	
DECLARE @sBHolRegionColumnName		sysname	

/* The Bank Holiday Instance (Child) Table which contains 25/12/00, 26/12/00 etc. */
DECLARE @iBHolTableID				int		
DECLARE @sBHolTableName			sysname	
DECLARE @sBHolDateColumnName		sysname	

/* Flag storing if the Bank Hols are setup ok and therefore if we should use them or not */
DECLARE @fBHolSetupOK           			bit

/* ID of the persons region...used to work out which dates from the BHol Instance table apply to the employee */
DECLARE @iBHolRegionID			int

/* Date variables used when working out the next change date for historic WP/Regions - If applicable */
DECLARE @dTempDate				datetime
DECLARE @dNextChange_Region		datetime
DECLARE @dNextChange_WP			datetime

/* Date variable used to cycle through dates between start date and end date */
DECLARE @dtCurrentDate			datetime

/* Flag stating if we are using historic region setup (True) or static (False) */
DECLARE @fHistoricRegion			bit

/* Variables to hold the relevant region table/column names */
DECLARE @sStaticRegionColumnName 		varchar(255)
DECLARE @sHistoricRegionTableName 		varchar(255)
DECLARE @sHistoricRegionColumnName 		varchar(255)
DECLARE @sHistoricRegionDateColumnName 	varchar(255)

/* Flag stating if we are using historic wp setup (True) or static (False) */
DECLARE @fHistoricWP				bit

/* Variables to hold the relevant wp table/column names */
DECLARE @sStaticWPColumnName 		varchar(255)
DECLARE @sHistoricWPTableName 		varchar(255)
DECLARE @sHistoricWPColumnName 		varchar(255)
DECLARE @sHistoricWPDateColumnName 	varchar(255)

/* The current wp/region being used in the calculation */
DECLARE @psWorkPattern	        		varchar(255)   
DECLARE @psPersonnelRegion			varchar(255)   

/* Flags derived from @psWorkPattern */
DECLARE @fWorkAM				bit
DECLARE @fWorkPM				bit
DECLARE @fWorkOnSundayAM			bit
DECLARE @fWorkOnSundayPM			bit
DECLARE @fWorkOnMondayAM			bit
DECLARE @fWorkOnMondayPM			bit
DECLARE @fWorkOnTuesdayAM			bit
DECLARE @fWorkOnTuesdayPM			bit
DECLARE @fWorkOnWednesdayAM		bit
DECLARE @fWorkOnWednesdayPM		bit
DECLARE @fWorkOnThursdayAM		bit
DECLARE @fWorkOnThursdayPM		bit
DECLARE @fWorkOnFridayAM			bit
DECLARE @fWorkOnFridayPM			bit
DECLARE @fWorkOnSaturdayAM		bit
DECLARE @fWorkOnSaturdayPM			bit
DECLARE @iDayOfWeek				int
DECLARE @sCommandString			nvarchar(4000)
DECLARE @iCount				int
DECLARE @sParamDefinition 			nvarchar(500)

/* Initialise the result to be 0 */
SET @pdblResult = 0

/* Get the current users group ID */
SELECT @iUserGroupID = sysusers.gid
FROM sysusers
WHERE sysusers.name = CURRENT_USER

/* Check if the current user is a System or Security manager. */
SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
FROM ASRSysGroupPermissions
INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
WHERE sysusers.uid = @iUserGroupID
	AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
	OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
	AND ASRSysGroupPermissions.permitted = 1
	AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

/* Get the ID of the BHol Region Table (which contains England, Scotland etc */
SELECT @iBHolRegionTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
AND ParameterKey = ''Param_TableBHolRegion''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Region Table (which contains England, Scotland etc */
SELECT @sBHolRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolRegionTableID

/* Get the name of the BHol Region column in the BHol Region Table */
SELECT @sBHolRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
	AND ParameterKey = ''Param_FieldBHolRegion''
	AND ParameterType = ''PType_ColumnID''

/* Get the ID of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @iBHolTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE''
AND ParameterKey = ''Param_TableBHol''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @sBHolTableName = AsrSysTables.TableName 
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolTableID

/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sBHolTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' AND ParameterKey = ''Param_TableBHol'' AND ParameterType = ''PType_TableID''))
END

/* Get the name of the BHol Date column. */
SELECT @sBHolDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' 
	AND ParameterKey = ''Param_FieldBHolDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to state whether BHols have been setup correctly or Not */
IF ((NOT @iBHolRegionTableID IS NULL)
AND (NOT @sBHolRegionTableName IS NULL) 
AND (NOT @sBHolRegionColumnName IS NULL) 
AND (NOT @iBHolTableID IS NULL) 
AND (NOT @sBHolTableName IS NULL) 
AND (NOT @sBHolDateColumnName IS NULL))
BEGIN
	SET @fBHolSetupOK = 1
END
ELSE
BEGIN
	SET @fBHolSetupOK = 0
END

/* Get the ID of the Personnel Table */
SELECT @iPersonnelTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL''
AND ParameterKey = ''Param_TablePersonnel''
AND ParameterType = ''PType_TableID''

/* Get the name of the Personnel Table */
SELECT @sPersonnelTable = AsrSysTables.TableName 
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_TablePersonnel''
	AND ParameterType = ''PType_TableID''

/* Get the Region Setup - Static Region*/
SELECT @sStaticRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsRegion''
	AND ParameterType = ''PType_ColumnID''
	
/* Get the Region Setup - Historic Region*/
SELECT @sHistoricRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegionTable''
	AND ParameterType = ''PType_TableID''

/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sHistoricRegionTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHRegionTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegion''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricRegionDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegionDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of regions we are to use */
IF @sStaticRegionColumnName is null
BEGIN
	IF (@sHistoricRegionTableName is null) OR (@sHistoricRegionColumnName is null) OR (@sHistoricRegionDateColumnName is null)
	BEGIN
		SET @pdblResult = 0
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricRegion = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricRegion = 0
END

/* Get the WP Setup - Static WP*/
SELECT @sStaticWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

/* Get the Region Setup - Historic WP */
SELECT @sHistoricWPTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPatternTable''
	AND ParameterType = ''PType_TableID''

/* If user does not have sys/sec permission then replace child table name with correct asrsyschildview */
IF @fsyssecmgr = 0
BEGIN
	SELECT @sHistoricWPTableName = sysobjects.name
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) FROM ASRSysChildViews where tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHWorkingPatternTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricWPDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPatternDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of wp we are to use */
IF @sStaticWPColumnName is null
BEGIN
	IF (@sHistoricWPTableName is null) OR (@sHistoricWPColumnName is null) OR (@sHistoricWPDateColumnName is null)
	BEGIN
		SET @pdblResult = 0
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricWP = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricWP = 0
END

/* Calculate the Absence Duration if all parameters have been provided. */
IF (NOT @pdtStartDate IS NULL) AND (NOT @psStartSession IS NULL) AND (NOT @pdtEndDate IS NULL) AND (NOT @psEndSession IS NULL)
BEGIN

	SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
	SET @pdtEndDate = convert(datetime, convert(varchar(20), @pdtEndDate, 101))
	SET @dtCurrentDate  = @pdtStartDate

	/* If we are using static wp and static region, do it the simple way */
	IF (@fHistoricRegion = 0) AND (@fHistoricWP = 0)
	BEGIN

		/* Get The Employees Working Pattern */
		SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
			
		/* Get The Employees Region */
		SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT

		/* Get the Region ID for the persons Region */
		SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +  '' FROM '' + @sBHolRegionTableName + '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
		SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT

		/* Determine which days are work days from the given work pattern. */
		SET @fWorkOnSundayAM = 0
		SET @fWorkOnSundayPM = 0
		SET @fWorkOnMondayAM = 0
		SET @fWorkOnMondayPM = 0
		SET @fWorkOnTuesdayAM = 0
		SET @fWorkOnTuesdayPM = 0
		SET @fWorkOnWednesdayAM = 0
		SET @fWorkOnWednesdayPM = 0
		SET @fWorkOnThursdayAM = 0
		SET @fWorkOnThursdayPM = 0
		SET @fWorkOnFridayAM = 0
		SET @fWorkOnFridayPM = 0
		SET @fWorkOnSaturdayAM = 0
		SET @fWorkOnSaturdayPM = 0

		IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
		IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
		IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
		IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
		IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
		IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
		IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
		IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
		IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
		IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
		IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
		IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
		IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
		IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

		WHILE @dtCurrentDate <= @pdtEndDate
		BEGIN

			/* Check if the current date is a work day. */
			SET @fWorkAM = 0
			SET @fWorkPM = 0
			SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)

			IF @iDayOfWeek = 1 
			BEGIN
				SET @fWorkAM = @fWorkOnSundayAM
				SET @fWorkPM = @fWorkOnSundayPM
			END
			IF @iDayOfWeek = 2
			BEGIN
				SET @fWorkAM = @fWorkOnMondayAM
				SET @fWorkPM = @fWorkOnMondayPM
			END
			IF @iDayOfWeek = 3
			BEGIN
				SET @fWorkAM = @fWorkOnTuesdayAM
				SET @fWorkPM = @fWorkOnTuesdayPM
			END
			IF @iDayOfWeek = 4
			BEGIN
				SET @fWorkAM = @fWorkOnWednesdayAM
				SET @fWorkPM = @fWorkOnWednesdayPM
			END
			IF @iDayOfWeek = 5
			BEGIN
				SET @fWorkAM = @fWorkOnThursdayAM
				SET @fWorkPM = @fWorkOnThursdayPM
			END
			IF @iDayOfWeek = 6
			BEGIN
				SET @fWorkAM = @fWorkOnFridayAM
				SET @fWorkPM = @fWorkOnFridayPM
			END
			IF @iDayOfWeek = 7
			BEGIN
				SET @fWorkAM = @fWorkOnSaturdayAM
				SET @fWorkPM = @fWorkOnSaturdayPM
			END

			IF (@fWorkAM = 1) OR (@fWorkPM = 1)
			BEGIN
				IF @fBHolSetupOK = 1
				BEGIN

					/* Check that the current date is not a company holiday. */
					SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' + '' FROM '' + @sBHolTableName + 
 			 				               '' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
   								  '''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
					SET @sParamDefinition = N''@count int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT
	
					IF @iCount = 0
					BEGIN
						IF @dtCurrentDate = @pdtStartDate
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @dtCurrentDate = @pdtEndDate
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
						END
					END
				END
				ELSE
				BEGIN

					IF @dtCurrentDate = @pdtStartDate
					BEGIN	
						IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
						IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
					END
					ELSE
					BEGIN
						IF @dtCurrentDate = @pdtEndDate
						BEGIN
							IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
							IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
					END
				END
			END

		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1

		END

	END
	ELSE  /* else for if we are using all static or not */
	BEGIN 

		WHILE @dtCurrentDate <= @pdtEndDate
		BEGIN

			/* We are using a historic region, so ensure we have the right region for the @dCurrentDate */
			IF @fHistoricRegion = 1
			BEGIN
				/* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */
				IF (@dnextchange_region IS NULL) OR ((@dtCurrentDate >= @dNextChange_Region) And (@dtCurrentDate <> ''12/31/9999''))
				BEGIN

					/* Get The Employees Region For @dCurrentDate */
					SET @sCommandString = ''SELECT TOP 1 @psPersonnelRegion = '' + @sHistoricRegionColumnName +
								  '' FROM '' + @sHistoricRegionTableName +
								  '' WHERE '' + @sHistoricRegionDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' DESC'' 
					SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT


					/* Get the Region ID for the persons Region */
					SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
	               						 '' FROM '' + @sBHolRegionTableName + 
							              '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
					SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
						
					/* Get the date of next change for the Region */
					SET @sCommandString = ''SELECT TOP 1 @dTempDate = '' + @sHistoricRegionDateColumnName +
								  '' FROM '' + @sHistoricRegionTableName +
 								  '' WHERE '' + @sHistoricRegionDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' ASC''
					SET @sParamDefinition = N''@dTempDate datetime OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dTempDate OUTPUT
						
					IF @dTempDate IS NULL
					BEGIN
						SET @dNextChange_Region = ''12/31/9999''
					END
					ELSE
					BEGIN
						SET @dNextChange_Region = @dTempDate
					END
				END

			END
			ELSE
			BEGIN
				/* We are using a static region, so get it */
				SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
				SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT

				/* Get the Region ID for the persons Region */
				SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
                  					               '' FROM '' + @sBHolRegionTableName + 
						               '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
				SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
			END

			/* We are using a historic wp so ensure we are getting the right wp for @dCurrentDate */
			IF @fHistoricWP = 1  
			BEGIN
				IF (@dnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dNextChange_WP) And (@dtCurrentDate <> ''12/31/9999''))
				BEGIN
					/* Get The Employees WP For @dCurrentDate */
					SET @sCommandString = ''SELECT TOP 1 @psWorkPattern = '' + @sHistoricWPColumnName +
								  '' FROM '' + @sHistoricWPTableName +
								  '' WHERE '' + @sHistoricWPDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricWPDateColumnName + '' DESC''
					SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
					/* Get The next change date for WP */
					SET @sCommandString = ''SELECT TOP 1 @dTempDate = '' + @sHistoricWPDateColumnName +
								  '' FROM '' + @sHistoricWPTableName +
								  '' WHERE '' + @sHistoricWPDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
								  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
								  '' ORDER BY '' + @sHistoricWPDateColumnName + '' ASC''
					SET @sParamDefinition = N''@dTempDate datetime OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dTempDate OUTPUT
					IF @dTempDate IS NULL
					BEGIN
						SET @dNextChange_WP = ''12/31/9999''
					END
					ELSE
					BEGIN
						SET @dNextChange_WP = @dTempDate
					END
				END
			END
			ELSE
			BEGIN
				/* We are using a static wp, so get it */
				SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
				SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
			END

			/* Determine which days are work days from the given work pattern. */
			SET @fWorkOnSundayAM = 0
			SET @fWorkOnSundayPM = 0
			SET @fWorkOnMondayAM = 0
			SET @fWorkOnMondayPM = 0
			SET @fWorkOnTuesdayAM = 0
			SET @fWorkOnTuesdayPM = 0
			SET @fWorkOnWednesdayAM = 0
			SET @fWorkOnWednesdayPM = 0
			SET @fWorkOnThursdayAM = 0
			SET @fWorkOnThursdayPM = 0
			SET @fWorkOnFridayAM = 0
			SET @fWorkOnFridayPM = 0
			SET @fWorkOnSaturdayAM = 0
			SET @fWorkOnSaturdayPM = 0
		
			IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
			IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
			IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
			IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
			IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
			IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
			IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
			IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
			IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
			IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
			IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
			IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
			IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
			IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

			/* Check if the current date is a work day. */
			SET @fWorkAM = 0
			SET @fWorkPM = 0
			SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)
			IF @iDayOfWeek = 1 
			BEGIN
				SET @fWorkAM = @fWorkOnSundayAM
				SET @fWorkPM = @fWorkOnSundayPM
			END
			IF @iDayOfWeek = 2
			BEGIN
				SET @fWorkAM = @fWorkOnMondayAM
				SET @fWorkPM = @fWorkOnMondayPM
			END
			IF @iDayOfWeek = 3
			BEGIN
				SET @fWorkAM = @fWorkOnTuesdayAM
				SET @fWorkPM = @fWorkOnTuesdayPM
			END
			IF @iDayOfWeek = 4
			BEGIN
				SET @fWorkAM = @fWorkOnWednesdayAM
				SET @fWorkPM = @fWorkOnWednesdayPM
			END
			IF @iDayOfWeek = 5
			BEGIN
				SET @fWorkAM = @fWorkOnThursdayAM
				SET @fWorkPM = @fWorkOnThursdayPM
			END
			IF @iDayOfWeek = 6
			BEGIN
				SET @fWorkAM = @fWorkOnFridayAM
				SET @fWorkPM = @fWorkOnFridayPM
			END
			IF @iDayOfWeek = 7
			BEGIN
				SET @fWorkAM = @fWorkOnSaturdayAM
				SET @fWorkPM = @fWorkOnSaturdayPM
			END

			IF (@fWorkAM = 1) OR (@fWorkPM = 1)
			BEGIN
				IF @fBHolSetupOK = 1
				BEGIN

					/* Check that the current date is not a company holiday. */
					SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' +
              							  '' FROM '' + @sBHolTableName + 
							               '' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
  								  '''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
					SET @sParamDefinition = N''@count int OUTPUT''
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT

					IF @iCount = 0
					BEGIN
						IF @dtCurrentDate = @pdtStartDate
						BEGIN	
							IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @dtCurrentDate = @pdtEndDate
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
							END
							ELSE
							BEGIN
								IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
								IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
							END
						END
					END
				END
				ELSE
				BEGIN
					/* We arent using Bholidays, so just add to the result */
					IF @dtCurrentDate = @pdtStartDate
					BEGIN	
						IF ((@fWorkAM = 1) AND (UPPER(@psStartSession) = ''AM'')) SET @pdblResult = @pdblResult + 0.5
						IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
					END
					ELSE
					BEGIN
						IF @dtCurrentDate = @pdtEndDate
						BEGIN
							IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
							IF ((@fWorkPM = 1) AND (UPPER(@psEndSession) = ''PM''))  SET @pdblResult = @pdblResult + 0.5
						END
						ELSE
						BEGIN
							IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
							IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
						END
					END
				END
			END
		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1

		END

	END /* end for if we are using all static or not */

END /* end for if all the parameters have been provided */

END')


/* ----------------------------------------------------- */
/* Drop and recreate sp_ASRFn_WorkingDaysBetweenTwoDates */
/* ----------------------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_WorkingDaysBetweenTwoDates]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_WorkingDaysBetweenTwoDates]

EXEC('CREATE PROCEDURE sp_ASRFn_WorkingDaysBetweenTwoDates(

@pdblResult			float OUTPUT,
@pdtStartDate			datetime,
@pdtEndDate			datetime,
@iPersonnelID			int)

AS

BEGIN

/* Used to work out if we can hit child tables directly, or via childviews */
DECLARE @iUserGroupID		int
DECLARE @fSysSecMgr		bit

/* Personnel Table ID and name...used for static region/wp and for ID_xx childviews */
DECLARE @sPersonnelTable			varchar(255)
DECLARE @iPersonnelTableID			int

/* The Bank Holiday Region (Primary) Table which contains England, Scotland, Wales etc. */
DECLARE @iBHolRegionTableID			int
DECLARE @sBHolRegionTableName		varchar(255)
DECLARE @sBHolRegionColumnName		varchar(255)

/* The Bank Holiday Instance (Child) Table which contains 25/12/00, 26/12/00 etc. */
DECLARE @iBHolTableID				int
DECLARE @sBHolTableName			sysname
DECLARE @sBHolDateColumnName		sysname

/* Flag storing if the Bank Hols are setup OK and therefore if we should use them or not */
DECLARE @fBHolSetupOK			bit

/* ID of the persons region...used to work out which dates from the BHol Instance table apply to the employee */
DECLARE @iBHolRegionID			int

/* Date counter to loop thru from StartDate to EndDate */
DECLARE @dtCurrentDate           		datetime

/* Date variables used when working out the next change date for historic WP/Regions - If applicable */
DECLARE @dtTempDate				datetime
DECLARE @dtNextChange_Region			datetime
DECLARE @dtNextChange_WP			datetime

/* Flag stating if we are using historic region setup (True) or static (False) */
DECLARE @fHistoricRegion			bit

/* Variables to hold the relevant region table/column names */
DECLARE @sStaticRegionColumnName		varchar(255)
DECLARE @sHistoricRegionTableName		varchar(255)
DECLARE @sHistoricRegionColumnName		varchar(255)
DECLARE @sHistoricRegionDateColumnName		varchar(255)

/* Flag stating if we are using historic wp setup (True) or static (False) */
DECLARE @fHistoricWP				bit

/* Variables to hold the relevant wp table/column names */
DECLARE @sStaticWPColumnName		varchar(255)
DECLARE @sHistoricWPTableName		varchar(255)
DECLARE @sHistoricWPColumnName		varchar(255)
DECLARE @sHistoricWPDateColumnName		varchar(255)

/* The current wp/region being used in the calculation */
DECLARE @psWorkPattern			varchar(255)
DECLARE @psPersonnelRegion		varchar(255)

/* Working Pattern Stuff */
DECLARE @fWorkAM                 	bit
DECLARE @fWorkPM                 	bit
DECLARE @fWorkOnSundayAM         		bit
DECLARE @fWorkOnSundayPM         		bit
DECLARE @fWorkOnMondayAM         		bit
DECLARE @fWorkOnMondayPM         		bit
DECLARE @fWorkOnTuesdayAM			bit
DECLARE @fWorkOnTuesdayPM       		bit
DECLARE @fWorkOnWednesdayAM 			bit
DECLARE @fWorkOnWednesdayPM      		bit
DECLARE @fWorkOnThursdayAM       		bit
DECLARE @fWorkOnThursdayPM       		bit
DECLARE @fWorkOnFridayAM         		bit
DECLARE @fWorkOnFridayPM         		bit
DECLARE @fWorkOnSaturdayAM       		bit
DECLARE @fWorkOnSaturdayPM       		bit
DECLARE @iDayOfWeek              		int

/* Command Strings for Dynamic Calls To SQL */
DECLARE @sCommandString          		nvarchar(4000)
DECLARE @iCount                  			int
DECLARE @sParamDefinition        		nvarchar(500)

/* Initialise the result to be 0 */
SET @pdblResult = 0

/* If Calculate the Absence Duration if all parameters are valid. */
IF (@pdtStartDate IS NULL) OR (@pdtEndDate IS NULL) OR (@iPersonnelID IS NULL) 
BEGIN
	RETURN
END

/* Get the current users group ID */
SELECT @iUserGroupID = sysusers.gid
FROM sysusers
WHERE sysusers.name = CURRENT_USER

/* Check if the current user is a System or Security manager. */
SELECT @fSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
FROM ASRSysGroupPermissions
INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
WHERE sysusers.uid = @iUserGroupID
AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER''
OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
AND ASRSysGroupPermissions.permitted = 1
AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''

/* Get the ID of the BHol Region Table (which contains England, Scotland etc */
SELECT @iBHolRegionTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' 
AND ParameterKey = ''Param_TableBHolRegion''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Region Table (which contains England, Scotland etc */
SELECT @sBHolRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolRegionTableID

/* Get the name of the BHol Region column in the BHol Region Table */
SELECT @sBHolRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' 
	AND ParameterKey = ''Param_FieldBHolRegion''
	AND ParameterType = ''PType_ColumnID''

/* Get the ID of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @iBHolTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' 
AND ParameterKey = ''Param_TableBHol''
AND ParameterType = ''PType_TableID''

/* Get the Name of the BHol Table (which contains instances of BHols eg 25/12/00, 01/01/01 etc */
SELECT @sBHolTableName = AsrSysTables.TableName 
FROM AsrSysTables 
WHERE AsrSysTables.TableID = @iBHolTableID

/* If user cant hit BHol Table directly, get childview Name */
IF @fSysSecMgr = 0
BEGIN
	SELECT @sBHolTableName = sysobjects.name
	FROM sysprotects
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + CONVERT(varchar(8000), childViewID) FROM ASRSysChildViews WHERE tableID = @iBHolTableID)
END

/* Get the name of the BHol Date column. */
SELECT @sBHolDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_ABSENCE'' 
	AND ParameterKey = ''Param_FieldBHolDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to state whether BHols have been setup correctly or Not */
IF ((NOT @iBHolRegionTableID IS NULL)
AND (NOT @sBHolRegionTableName IS NULL) 
AND (NOT @sBHolRegionColumnName IS NULL) 
AND (NOT @iBHolTableID IS NULL) 
AND (NOT @sBHolTableName IS NULL) 
AND (NOT @sBHolDateColumnName IS NULL))
BEGIN
	SET @fBHolSetupOK = 1
END
ELSE
BEGIN
	SET @fBHolSetupOK = 0
END

/* Get the ID of the Personnel Table */
SELECT @iPersonnelTableID = AsrSysModuleSetup.ParameterValue
FROM AsrSysModuleSetup
WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
AND ParameterKey = ''Param_TablePersonnel''
AND ParameterType = ''PType_TableID''

/* Get the name of the Personnel Table */
SELECT @sPersonnelTable = AsrSysTables.TableName 
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_TablePersonnel''
	AND ParameterType = ''PType_TableID''

/* Get the Region Setup - Static Region*/
SELECT @sStaticRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsRegion''
	AND ParameterType = ''PType_ColumnID''
	
/* Get the Region Setup - Historic Region*/
SELECT @sHistoricRegionTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegionTable''
	AND ParameterType = ''PType_TableID''

IF @fSysSecMgr = 0
BEGIN
	SELECT @sHistoricRegionTableName = sysobjects.name
	FROM sysprotects
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + CONVERT(varchar(8000), childViewID) FROM ASRSysChildViews WHERE tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHRegionTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricRegionColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegion''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricRegionDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHRegionDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of regions we are to use */
IF @sStaticRegionColumnName is null
BEGIN
	IF (@sHistoricRegionTableName is null) OR (@sHistoricRegionColumnName is null) OR (@sHistoricRegionDateColumnName is null)
	BEGIN
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricRegion = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricRegion = 0
END

/* Get the WP Setup - Static WP*/
SELECT @sStaticWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

/* Get the Region Setup - Historic WP */
SELECT @sHistoricWPTableName = AsrSysTables.TableName
FROM AsrSysTables 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysTables.TableID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPatternTable''
	AND ParameterType = ''PType_TableID''

IF @fSysSecMgr = 0
BEGIN
	SELECT @sHistoricWPTableName = sysobjects.name
	FROM sysprotects
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
	AND sysprotects.protectType <> 206
	AND sysprotects.action = 193
	AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + CONVERT(varchar(8000), childViewID) FROM ASRSysChildViews WHERE tableID = (SELECT AsrSysModuleSetup.ParameterValue FROM ASRSysModuleSetup WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' AND ParameterKey = ''Param_FieldsHWorkingPatternTable'' AND ParameterType = ''PType_TableID''))
END

SELECT @sHistoricWPColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPattern''
	AND ParameterType = ''PType_ColumnID''

SELECT @sHistoricWPDateColumnName = AsrSysColumns.ColumnName 
FROM AsrSysColumns 
INNER JOIN AsrSysModuleSetup 
	ON AsrSysColumns.ColumnID = AsrSysModuleSetup.ParameterValue
	WHERE AsrSysModuleSetup.ModuleKey = ''MODULE_PERSONNEL'' 
	AND ParameterKey = ''Param_FieldsHWorkingPatternDate''
	AND ParameterType = ''PType_ColumnID''

/* Set flag to indicate what type of wp we are to use */
IF @sStaticWPColumnName is null
BEGIN
	IF (@sHistoricWPTableName is null) OR (@sHistoricWPColumnName is null) OR (@sHistoricWPDateColumnName is null)
	BEGIN
		RETURN
	END
	ELSE
	BEGIN
		SET @fHistoricWP = 1
	END
END
ELSE
BEGIN	
	SET @fHistoricWP = 0
END

/* Make sure the variables are nice sql dates */
SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
SET @pdtEndDate = convert(datetime, convert(varchar(20), @pdtEndDate, 101))

SET @dtCurrentDate = @pdtStartDate 

/* If we are using static wp and static region, do it the simple way */
IF (@fHistoricRegion = 0) AND (@fHistoricWP = 0)
BEGIN
	/* Get The Employees Working Pattern */
	SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
	SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
	EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT

	/* If we are including bank holidays, get the region information */
	IF @fBHolSetupOK = 1 
	BEGIN			
		/* Get The Employees Region */
		SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT

		/* Get the Region ID for the persons Region */
		SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +  '' FROM '' + @sBHolRegionTableName + '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
		SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
	END

	/* Determine which days are work days from the given work pattern. */
	SET @fWorkOnSundayAM = 0
	SET @fWorkOnSundayPM = 0
	SET @fWorkOnMondayAM = 0
	SET @fWorkOnMondayPM = 0
	SET @fWorkOnTuesdayAM = 0
	SET @fWorkOnTuesdayPM = 0
	SET @fWorkOnWednesdayAM = 0
	SET @fWorkOnWednesdayPM = 0
	SET @fWorkOnThursdayAM = 0
	SET @fWorkOnThursdayPM = 0
	SET @fWorkOnFridayAM = 0
	SET @fWorkOnFridayPM = 0
	SET @fWorkOnSaturdayAM = 0
	SET @fWorkOnSaturdayPM = 0

	IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
	IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
	IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
	IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
	IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
	IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
	IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
	IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
	IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
	IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
	IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
	IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
	IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
	IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

	/* Loop through absence, only counting dates btwn the rpt dates */
	WHILE @dtCurrentDate <= @pdtEndDate
	BEGIN

		/* Check if the current date is a work day. */
		SET @fWorkAM = 0
		SET @fWorkPM = 0
		SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)

		/* Set work am/pm variables */
		IF @iDayOfWeek = 1
		BEGIN
			SET @fWorkAM = @fWorkOnSundayAM
			SET @fWorkPM = @fWorkOnSundayPM
		END
		IF @iDayOfWeek = 2
		BEGIN
			SET @fWorkAM = @fWorkOnMondayAM
			SET @fWorkPM = @fWorkOnMondayPM
		END
		IF @iDayOfWeek = 3
		BEGIN
			SET @fWorkAM = @fWorkOnTuesdayAM
			SET @fWorkPM = @fWorkOnTuesdayPM
		END
		IF @iDayOfWeek = 4
		BEGIN
			SET @fWorkAM = @fWorkOnWednesdayAM
			SET @fWorkPM = @fWorkOnWednesdayPM
		END
		IF @iDayOfWeek = 5
		BEGIN
			SET @fWorkAM = @fWorkOnThursdayAM
			SET @fWorkPM = @fWorkOnThursdayPM
	       	END
		IF @iDayOfWeek = 6
		BEGIN
			SET @fWorkAM = @fWorkOnFridayAM
			SET @fWorkPM = @fWorkOnFridayPM
		END
		IF @iDayOfWeek = 7
		BEGIN
			SET @fWorkAM = @fWorkOnSaturdayAM
			SET @fWorkPM = @fWorkOnSaturdayPM
		END

		/* If its a working day */
		IF (@fWorkAM = 1) OR (@fWorkPM = 1)
		BEGIN

			/* If we are including bank holidays, check for Bhols */
			IF @fBHolSetupOK = 1
			BEGIN
	
				/* Check that the current date is not a company holiday. */
				SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' + '' FROM '' + @sBHolTableName + 
							'' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
   							'''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
				SET @sParamDefinition = N''@count int OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT
		
				IF @iCount = 0
				BEGIN
					IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
					IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
				END
			END	
			ELSE
			BEGIN
				/* We arent using BHols, so just add to the result without checking the bhol table */
				IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
				IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
			END
		END

		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1
	END
END
ELSE
BEGIN

	/* either historic wp or region, so do this...*/
	/* Loop through absence, only counting dates btwn the rpt dates */
	WHILE @dtCurrentDate <= @pdtEndDate
	BEGIN

	/* We are using a historic region, so ensure we have the right region for the @dtCurrentDate */
	If @fHistoricRegion = 1
	BEGIN			

		/* Only bother checking we have the right region if we dont know the nxt chg date or the current date is equal to nxt chg date */
		IF (@dtnextchange_region IS NULL) OR ((@dtCurrentDate >= @dtNextChange_Region) And (@dtCurrentDate <> ''12/31/9999''))
		BEGIN
						
			/* Get The Employees Region For @dCurrentDate */
			SET @sCommandString = ''SELECT TOP 1 @psPersonnelRegion = '' + @sHistoricRegionColumnName +
						  '' FROM '' + @sHistoricRegionTableName +
						  '' WHERE '' + @sHistoricRegionDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
						  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
						  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' DESC''
			SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT
	
			/* Get the Region ID for the persons Region */
			SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
						 '' FROM '' + @sBHolRegionTableName + 
				              '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
			SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
							
			/* Get the date of next change for the Region */
			SET @sCommandString = ''SELECT TOP 1 @dtTempDate = '' + @sHistoricRegionDateColumnName +
						  '' FROM '' + @sHistoricRegionTableName +
 						  '' WHERE '' + @sHistoricRegionDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
						  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
						  '' ORDER BY '' + @sHistoricRegionDateColumnName + '' ASC''
			SET @sParamDefinition = N''@dtTempDate datetime OUTPUT''
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dtTempDate OUTPUT
							
			IF @dtTempDate IS NULL
			BEGIN
				SET @dtNextChange_Region = ''12/31/9999''
			END
			ELSE
			BEGIN
				SET @dtNextChange_Region = @dtTempDate
			END
		END	
	END	
	ELSE
	BEGIN
		/* We are using a static region, so get it */
		SET @sCommandString = ''SELECT @psPersonnelRegion = '' + @sStaticRegionColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psPersonnelRegion varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psPersonnelRegion OUTPUT
	
		/* Get the Region ID for the persons Region */
		SET @sCommandString = ''SELECT @iBHolRegionID = ID '' +
			               '' FROM '' + @sBHolRegionTableName + 
			               '' WHERE '' + @sBHolRegionColumnName + '' = '''''' + @psPersonnelRegion + ''''''''
		SET @sParamDefinition = N''@iBHolRegionID int OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iBHolRegionID OUTPUT
	END
					
	/* We are using a historic wp so ensure we are getting the right wp for @dCurrentDate */
	IF @fHistoricWP = 1  
	BEGIN
		IF (@dtnextchange_WP IS NULL) OR ((@dtCurrentDate >= @dtNextChange_WP) And (@dtCurrentDate <> ''12/31/9999''))
		BEGIN
			/* Get The Employees WP For @dCurrentDate */
			SET @sCommandString = ''SELECT TOP 1 @psWorkPattern = '' + @sHistoricWPColumnName +
						  '' FROM '' + @sHistoricWPTableName +
						  '' WHERE '' + @sHistoricWPDateColumnName + '' <= '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
						  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
						  '' ORDER BY '' + @sHistoricWPDateColumnName + '' DESC'' 
			SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
	
			/* Get The next change date for WP */
			SET @sCommandString = ''SELECT TOP 1 @dtTempDate = '' + @sHistoricWPDateColumnName +
						  '' FROM '' + @sHistoricWPTableName +
						  '' WHERE '' + @sHistoricWPDateColumnName + '' > '''''' + convert(varchar(255), @dtCurrentDate,101) + '''''''' +
						  '' AND ID_'' + convert(varchar(20),@iPersonnelTableID) + '' = '' + convert(varchar(20),@iPersonnelID) +
						  '' ORDER BY '' + @sHistoricWPDateColumnName + '' ASC'' 
			SET @sParamDefinition = N''@dtTempDate datetime OUTPUT''
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dtTempDate OUTPUT
	
			IF @dtTempDate IS NULL
			BEGIN
				SET @dtNextChange_WP = ''12/31/9999''
			END
			ELSE
			BEGIN
				SET @dtNextChange_WP = @dtTempDate
			END
		END
	END
	ELSE
	BEGIN
		/* We are using a static wp, so get it */
		SET @sCommandString = ''SELECT @psWorkPattern = '' + @sStaticWPColumnName + '' FROM '' + @sPersonnelTable + '' WHERE ID = '' + convert(varchar(255), @iPersonnelID)
		SET @sParamDefinition = N''@psWorkPattern varchar(255) OUTPUT''
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psWorkPattern OUTPUT
	END
		
	/* Determine which days are work days from the given work pattern. */
	SET @fWorkOnSundayAM = 0
	SET @fWorkOnSundayPM = 0
	SET @fWorkOnMondayAM = 0
	SET @fWorkOnMondayPM = 0
	SET @fWorkOnTuesdayAM = 0
	SET @fWorkOnTuesdayPM = 0
	SET @fWorkOnWednesdayAM = 0
	SET @fWorkOnWednesdayPM = 0
	SET @fWorkOnThursdayAM = 0
	SET @fWorkOnThursdayPM = 0
	SET @fWorkOnFridayAM = 0
	SET @fWorkOnFridayPM = 0
	SET @fWorkOnSaturdayAM = 0
	SET @fWorkOnSaturdayPM = 0

	IF LEN(@psWorkPattern) > 0 IF SUBSTRING(@psWorkPattern, 1, 1) <> '' '' SET @fWorkOnSundayAM = 1
	IF LEN(@psWorkPattern) > 1 IF SUBSTRING(@psWorkPattern, 2, 1) <> '' '' SET @fWorkOnSundayPM = 1
	IF LEN(@psWorkPattern) > 2 IF SUBSTRING(@psWorkPattern, 3, 1) <> '' '' SET @fWorkOnMondayAM = 1
	IF LEN(@psWorkPattern) > 3 IF SUBSTRING(@psWorkPattern, 4, 1) <> '' '' SET @fWorkOnMondayPM = 1
	IF LEN(@psWorkPattern) > 4 IF SUBSTRING(@psWorkPattern, 5, 1) <> '' '' SET @fWorkOnTuesdayAM = 1
	IF LEN(@psWorkPattern) > 5 IF SUBSTRING(@psWorkPattern, 6, 1) <> '' '' SET @fWorkOnTuesdayPM = 1
	IF LEN(@psWorkPattern) > 6 IF SUBSTRING(@psWorkPattern, 7, 1) <> '' '' SET @fWorkOnWednesdayAM = 1
	IF LEN(@psWorkPattern) > 7 IF SUBSTRING(@psWorkPattern, 8, 1) <> '' '' SET @fWorkOnWednesdayPM = 1
	IF LEN(@psWorkPattern) > 8 IF SUBSTRING(@psWorkPattern, 9, 1) <> '' '' SET @fWorkOnThursdayAM = 1
	IF LEN(@psWorkPattern) > 9 IF SUBSTRING(@psWorkPattern, 10, 1) <> '' '' SET @fWorkOnThursdayPM = 1
	IF LEN(@psWorkPattern) > 10 IF SUBSTRING(@psWorkPattern, 11, 1) <> '' '' SET @fWorkOnFridayAM = 1
	IF LEN(@psWorkPattern) > 11 IF SUBSTRING(@psWorkPattern, 12, 1) <> '' '' SET @fWorkOnFridayPM = 1
	IF LEN(@psWorkPattern) > 12 IF SUBSTRING(@psWorkPattern, 13, 1) <> '' '' SET @fWorkOnSaturdayAM = 1
	IF LEN(@psWorkPattern) > 13 IF SUBSTRING(@psWorkPattern, 14, 1) <> '' '' SET @fWorkOnSaturdayPM = 1

	/* Loop through absence, only counting dates btwn the rpt dates
	WHILE @dtCurrentDate <= @pdtEndDate
	BEGIN*/

		/* Check if the current date is a work day. */
		SET @fWorkAM = 0
		SET @fWorkPM = 0
		SET @iDayOfWeek = DATEPART(weekday, @dtCurrentDate)

		IF @iDayOfWeek = 1 
		BEGIN
			SET @fWorkAM = @fWorkOnSundayAM
			SET @fWorkPM = @fWorkOnSundayPM
		END
		IF @iDayOfWeek = 2
		BEGIN
			SET @fWorkAM = @fWorkOnMondayAM
			SET @fWorkPM = @fWorkOnMondayPM
		END
		IF @iDayOfWeek = 3
		BEGIN
			SET @fWorkAM = @fWorkOnTuesdayAM
			SET @fWorkPM = @fWorkOnTuesdayPM
		END
		IF @iDayOfWeek = 4
		BEGIN
			SET @fWorkAM = @fWorkOnWednesdayAM
			SET @fWorkPM = @fWorkOnWednesdayPM
		END
		IF @iDayOfWeek = 5
		BEGIN
			SET @fWorkAM = @fWorkOnThursdayAM
			SET @fWorkPM = @fWorkOnThursdayPM
		END
		IF @iDayOfWeek = 6
		BEGIN
			SET @fWorkAM = @fWorkOnFridayAM
			SET @fWorkPM = @fWorkOnFridayPM
		END
		IF @iDayOfWeek = 7
		BEGIN
			SET @fWorkAM = @fWorkOnSaturdayAM
			SET @fWorkPM = @fWorkOnSaturdayPM
		END

		IF (@fWorkAM = 1) OR (@fWorkPM = 1)
		BEGIN
			IF @fBHolSetupOK = 1
			BEGIN
				/* Check that the current date is not a company holiday. */
				SET @sCommandString = ''SELECT @count = COUNT('' + @sBHolDateColumnName + '')'' +
						      '' FROM '' + @sBHolTableName + 
					              '' WHERE convert(varchar(20), '' + @sBHolDateColumnName + '', 101) = '''''' + convert(varchar(20), @dtCurrentDate, 101) + 
  						      '''''' AND '' + @sBHolTableName + ''.ID_'' + convert(varchar(20),@iBHolRegionTableID) + '' = '' + convert(varchar(20),@iBHolRegionID)
				SET @sParamDefinition = N''@count int OUTPUT''
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iCount OUTPUT

				IF @iCount = 0
				BEGIN
					IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
					IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
				END
			END
			ELSE
			BEGIN
				/* We arent using Bholidays, so just add to the result */
				IF @fWorkAM = 1 SET @pdblResult = @pdblResult + 0.5
				IF @fWorkPM = 1 SET @pdblResult = @pdblResult + 0.5
			END
		END

		/* Move onto the next date. */
		SET @dtCurrentDate = @dtCurrentDate + 1

	END

END  /* end of the if all static else historic condition */

END')


/* ---------------------------------------- */
/* Drop and recreate sp_ASRFn_ServiceMonths */
/* ---------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ServiceMonths]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ServiceMonths]

EXEC('CREATE PROCEDURE sp_ASRFn_ServiceMonths 
(
	@piResult		integer OUTPUT,
	@pdtFirstDate 		datetime,
	@pdtSecondDate	datetime
)
AS
BEGIN

	DECLARE @dtTempDate	datetime

	/* If start date is in the future then return zero */
	IF datediff(d,@pdtFirstDate,getdate()) < 1 or @pdtFirstDate IS null
	BEGIN
		SET @piResult = 0
	END
	ELSE
	BEGIN

		/* If leaving date is in the future or blank then calculate from todays date minus start date */
		IF datediff(d,@pdtSecondDate,getdate()) < 1 or @pdtSecondDate IS null
		BEGIN
			SET @dtTempDate = getdate()
		END
		ELSE
		/* If leaving date is in past then calculate from leaving date minus start date */
		BEGIN
			SET @dtTempDate = @pdtSecondDate
		END

		EXEC sp_ASRFn_WholeMonthsBetweenTwoDates @piResult OUTPUT, @pdtFirstDate, @dtTempDate

		/* NOTE % 12 means divide by 12 and return the remainder */
		/* Remove any whole years from the result */
		SET @piResult = @piResult % 12

	END

END')

/* --------------------------------------- */
/* Drop and recreate sp_ASRFn_ServiceYears */
/* --------------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ServiceYears]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ServiceYears]

EXEC('CREATE PROCEDURE sp_ASRFn_ServiceYears 
(
	@piResult		integer OUTPUT,
	@pdtFirstDate 		datetime,
	@pdtSecondDate	datetime
)
AS
BEGIN

	DECLARE @pdtTempDate	datetime

	/* If start date is in the future then return zero */
	IF datediff(d,@pdtFirstDate,getdate()) < 1 or @pdtFirstDate IS null
	BEGIN
		SET @piResult = 0
	END
	ELSE
	BEGIN
		/* If leaving date is in the future or blank then calculate from todays date minus start date */
		IF datediff(d,@pdtSecondDate,getdate()) < 1 or @pdtSecondDate IS null
		BEGIN
			SET @pdtTempDate = getdate()
		END
		ELSE
		/* If leaving date is in past then calculate from leaving date minus start date */
		BEGIN
			SET @pdtTempDate = @pdtSecondDate
		END
		EXEC sp_ASRFn_WholeYearsBetweenTwoDates @piResult OUTPUT, @pdtFirstDate, @pdtTempDate
	END

END')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Set the flag to refresh the stored procedures               */
/* ----------------------------------------------------------- */

UPDATE ASRSysConfig
SET databaseVersion = 14,
	systemManagerVersion = '1.1.12',
	securityManagerVersion = '1.1.12',
	dataManagerVersion = '1.1.12',
        intranetversion = '0.0.3',
	refreshStoredProcedures = 1

/* RH Note : intranet version is prob already 0.0.3 as JPD releases fixed sp's to QA as an when done */
