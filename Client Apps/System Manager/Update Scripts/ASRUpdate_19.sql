/* -------------------------------------------------- */
/* Update the database from version 18 to version 19. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

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

/* Exit if the database is not version 18 or 19. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 18) or (@iDBVersion > 19)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* --------------------------------- */
/* Amend sp_ASRIntGetAvailableLogins */
/* --------------------------------- */

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetAvailableLogins]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_ASRIntGetAvailableLogins]

EXEC('CREATE PROCEDURE sp_ASRIntGetAvailableLogins 
AS
BEGIN
	/* Return a recordset of the SQL ''logins'' not already attached to ''users'' in the HR Pro database. */
	SELECT name
	FROM master.dbo.syslogins 
	WHERE name NOT IN (''sa'', ''probe'', ''SQLExec'', ''repl_publisher'', ''repl_subscriber'')
		AND isntName = 0
		AND name NOT IN(
			SELECT name 
			FROM sysusers
			WHERE isSQLRole = 0
				AND gid > 0
				AND sid IN (
					SELECT sid
					FROM master.dbo.syslogins 
					WHERE name NOT IN (''sa'', ''probe'',''SQLExec'',''repl_publisher'',''repl_subscriber'')
						AND isntname = 0
				)
		)
END')


/* ------------------------------ */
/* Amend sp_ASRIntGetFindRecords2 */
/* ------------------------------ */

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetFindRecords2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_ASRIntGetFindRecords2]

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
				AND sysobjects.name IN (SELECT ''ASRSysChildView_'' + convert(varchar(8000), childViewID) 
					FROM ASRSysChildViews where tableID = @piTableID)
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
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
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
						@sRealSource + ''.'' + @sColumnName
				END
				ELSE
				BEGIN
					/* Order column. */
					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						@sRealSource + ''.'' + @sColumnName +
						CASE 
							WHEN @fAscending = 0 THEN '' DESC'' 
							ELSE '''' 
						END				
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
						@sColumnTableName + ''.'' + @sColumnName
				END
				ELSE
				BEGIN
					/* Order column. */
					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						@sColumnTableName + ''.'' + @sColumnName + 
						CASE 
							WHEN @fAscending = 0 THEN '' DESC'' 
							ELSE '''' 
						END				
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
								WHEN @iDataType = 11 THEN ''convert(datetime, '' + @sSubString + '')''
								ELSE @sSubString 
							END
					END
					ELSE
					BEGIN
						/* Order column. */
						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN '','' 
								ELSE '''' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN ''convert(datetime, '' + @sSubString + '')''
								ELSE @sSubString 
							END + 
							CASE 
								WHEN @fAscending = 0 THEN '' DESC'' 
								ELSE '''' 
							END				
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


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

UPDATE ASRSysConfig
SET databaseVersion = 19,
	systemManagerVersion = '1.1.17',
	securityManagerVersion = '1.1.17',
	dataManagerVersion = '1.1.17'
/*,
	intranetversion = '0.0.6'*/

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 19 Has Converted Your HR Pro Database To Use V1.1.17 Of HR Pro'
