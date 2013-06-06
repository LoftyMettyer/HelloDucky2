/* -------------------------------------------------- */
/* Update the database from version 16 to version 17. */
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

/* Exit if the database is not version 16 or 17. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 16) or (@iDBVersion > 17)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* --------------------------------*/
/* Update sp_ASRIntGetFindRecords2 */
/* ------------------------------- */

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
						CASE 
							WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sRealSource + ''.'' + @sColumnName + '',103) AS ['' + @sColumnName + '']''
							ELSE @sRealSource + ''.'' + @sColumnName
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
							WHEN @iDataType = 11 THEN ''convert(datetime, '' + @sRealSource + ''.'' + @sColumnName + '')'' 
							ELSE @sRealSource + ''.'' + @sColumnName 
						END + 
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
						CASE 
							WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sColumnTableName + ''.'' + @sColumnName + '',103) AS ['' + @sColumnName + '']''
							ELSE @sColumnTableName + ''.'' + @sColumnName
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
							WHEN @iDataType = 11 THEN ''convert(datetime, '' + @sColumnTableName + ''.'' + @sColumnName + '')''
							ELSE @sColumnTableName + ''.'' + @sColumnName
						END + 
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
								WHEN @iDataType = 11 THEN ''convert(varchar(8000), '' + @sSubString + '',103) AS ['' + @sColumnName + '']''
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


/* ----------------------------------- */
/* Update sp_ASRFn_ConvertToPropercase */
/* ----------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ConvertToPropercase]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ConvertToPropercase]


EXEC('CREATE PROCEDURE sp_ASRFn_ConvertToPropercase 
(
	@psResult		varchar(8000) OUTPUT,
	@psStringToConvert 	varchar(8000)
)
AS
BEGIN
	DECLARE @iPosition		integer,
		@sLastCharacter		varchar(1)

	IF LEN(@psStringToConvert) > 1

	BEGIN
		/* Convert the first letter to uppercase, everything else to lowercase. */
		SET @psResult = UPPER(LEFT(@psStringToConvert, 1)) + LOWER(SUBSTRING(@psStringToConvert, 2, LEN(@psStringToConvert)-1))

		SET @iPosition = 2
		WHILE @iPosition <= LEN(@psStringToConvert)
		BEGIN
			SET @sLastCharacter = SUBSTRING(@psStringToConvert, @iPosition - 1, 1)

			/* Convert to uppercase any charatcers that follow non-alphabetic characters. eg. the N in O''Neill, or the L in Jean-Louis. */
			IF ((ASCII(@sLastCharacter) < ASCII(''A'')) OR (ASCII(@sLastCharacter) > ASCII(''Z''))) 
				AND ((ASCII(@sLastCharacter) < ASCII(''a'')) OR (ASCII(@sLastCharacter) > ASCII(''z''))) 
				AND ((ASCII(@sLastCharacter) < ASCII(''À'')) OR (ASCII(@sLastCharacter) > ASCII(''Ö''))) 
				AND ((ASCII(@sLastCharacter) < ASCII(''Ù'')) OR (ASCII(@sLastCharacter) > ASCII(''Ý''))) 
				AND ((ASCII(@sLastCharacter) < ASCII(''ß'')) OR (ASCII(@sLastCharacter) > ASCII(''ö''))) 
				AND ((ASCII(@sLastCharacter) < ASCII(''ù'')) OR (ASCII(@sLastCharacter) > ASCII(''ÿ''))) 
			BEGIN
             		 		SET @psResult = LEFT(@psResult, @iPosition - 1) + UPPER(SUBSTRING(@psResult, @iPosition, 1)) + RIGHT(@psResult, LEN(@psResult) - @iPosition)
			END
			ELSE
			BEGIN
				IF @iPosition > 2 
				BEGIN
				              /* Catch the McName.*/
             					IF (ASCII(SUBSTRING(@psResult, @iPosition - 2, 1)) = ASCII(''M'')) AND (ASCII(@sLastCharacter) = ASCII(''c'')) 
					BEGIN
						SET @psResult = LEFT(@psResult, @iPosition - 1) + UPPER(SUBSTRING(@psResult, @iPosition, 1)) + RIGHT(@psResult, LEN(@psResult) - @iPosition)
					END
				END
			END

			SET @iPosition = @iPosition + 1
		END
	END
	ELSE
	BEGIN
		SET @psResult = UPPER(@psStringToConvert)
	END
END')


/* ---------------------------------- */
/* Update sp_ASRIntGetHistoryMainMenu */
/* ---------------------------------- */

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
					WHERE ASRSysScreens.tableID = childScreens.tableID 
						AND ASRSysScreens.quickEntry = 0
						AND ASRSysScreens.screenID IN 
							(
								SELECT historyScreenID 
								FROM ASRSysHistoryScreens 
								WHERE parentScreenID = parentScreens.screenID
							)
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


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Set the flag to refresh the stored procedures               */
/* ----------------------------------------------------------- */

UPDATE ASRSysConfig
SET databaseVersion = 17,
	systemManagerVersion = '1.1.15',
	securityManagerVersion = '1.1.15',
	dataManagerVersion = '1.1.15',
	intranetversion = '0.0.5'

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 17 Has Converted Your HR Pro Database To Use V1.1.15 Of HR Pro'
