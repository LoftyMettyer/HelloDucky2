CREATE PROCEDURE [dbo].[sp_ASRIntAddFromWaitingList] (
	@piEmpRecordID		integer,
	@piCourseRecordID	integer,
	@psStatus			varchar(MAX)
)
AS
BEGIN
	DECLARE @iUserGroupID		integer,
		@sUserGroupName			sysname,
		@fSysSecMgr				bit,
		@sColumnList			varchar(MAX),
		@sValueList				varchar(MAX),
		@iChildViewID 			integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@sExecString			nvarchar(MAX),
		@iEmpTableID			integer,
		@iCourseTableID			integer,
		@iTBTableID				integer,
		@sTBTableName			sysname,
		@sTBRealSource			varchar(MAX),
		@iTBStatusColumnID		integer,
		@sTBStatusColumnName	sysname,
		@sTempTBColumnName		sysname,
		@iWLTableID				integer,
		@sWLTableName			sysname,
		@sWLRealSource			varchar(MAX),
		@sTempWLColumnName		sysname,
		@sAddedColumns			varchar(MAX),
		@iSourceColumnID		integer,
		@iDestinationColumnID	integer,
		@iIndex					integer,
		@fGranted				bit,
		@iWLCourseTitleColumnID	integer,
		@sWLCourseTitleColumnName	sysname,
		@sCourseTitle			varchar(MAX),
		@iCourseTitleColumnID	integer,
		@sCourseTitleColumnName	sysname,
		@sCourseSource			sysname,
		@sActualUserName		sysname;

	/* Clean the input string parameters. */
	IF len(@psStatus) > 0 SET @psStatus = replace(@psStatus, '''', '''''');

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) = 'SA'
	BEGIN
		SET @fSysSecMgr = 1;
	END
	ELSE
	BEGIN	

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
			AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS';
	END

	/* Get the EMPLOYEE table information. */
	SELECT @iEmpTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable';
	IF @iEmpTableID IS NULL SET @iEmpTableID = 0;

	/* Get the COURSE table information. */
	SELECT @iCourseTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTable';
	IF @iCourseTableID IS NULL SET @iCourseTableID = 0;

	SELECT @iCourseTitleColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTitle';
	IF @iCourseTitleColumnID IS NULL SET @iCourseTitleColumnID = 0;
	
	IF @iCourseTitleColumnID > 0 
	BEGIN
		SELECT @sCourseTitleColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iCourseTitleColumnID;
	END
	IF @sCourseTitleColumnName IS NULL SET @sCourseTitleColumnName = '';

	/* Get the @sCourseTitle value for the given course record. */
	DECLARE courseSourceCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT sysobjects.name
	FROM sysprotects
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
	WHERE sysprotects.uid = @iUserGroupID
		AND sysprotects.action = 193 
		AND (sysprotects.protectType = 205 OR sysprotects.protectType = 204)
		AND syscolumns.name = @sCourseTitleColumnName
		AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
			ASRSysTables.tableID = @iCourseTableID 
			UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iCourseTableID)
		AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
		OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
	OPEN courseSourceCursor;
	FETCH NEXT FROM courseSourceCursor INTO @sCourseSource;
	WHILE (@@fetch_status = 0) AND (@sCourseTitle IS null)
	BEGIN
		SET @sTempExecString = 'SELECT @sCourseTitle = ' + @sCourseTitleColumnName + 
			' FROM ' + @sCourseSource +
			' WHERE id = ' + convert(nvarchar(100), @piCourseRecordID);
		SET @sTempParamDefinition = N'@sCourseTitle varchar(MAX) OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sCourseTitle OUTPUT;

		FETCH NEXT FROM courseSourceCursor INTO @sCourseSource;
	END
	CLOSE courseSourceCursor;
	DEALLOCATE courseSourceCursor;

	IF @sCourseTitle IS null
	BEGIN
		SET @sCourseTitle = '';
	END

	/* Get the TRAINING BOOKING table information. */
	SELECT @iTBTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookTable';
	IF @iTBTableID IS NULL SET @iTBTableID = 0;

	SELECT @sTBTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iTBTableID;

	SELECT @iTBStatusColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookStatus';
	IF @iTBStatusColumnID IS NULL SET @iTBStatusColumnID = 0;

	SELECT @sTBStatusColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iTBStatusColumnID;

	/* Get the waiting list table information. */
	SELECT @iWLTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListTable';
	IF @iWLTableID IS NULL SET @iWLTableID = 0;

	SELECT @sWLTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iWLTableID;

	SELECT @iWLCourseTitleColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListCourseTitle';
	IF @iWLCourseTitleColumnID IS NULL SET @iWLCourseTitleColumnID = 0;
	
	IF @iWLCourseTitleColumnID > 0 
	BEGIN
		SELECT @sWLCourseTitleColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iWLCourseTitleColumnID;
	END
	IF @sWLCourseTitleColumnName IS NULL SET @sWLCourseTitleColumnName = '';

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iWLTableID
		AND role = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sWLRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sWLTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sWLRealSource = left(@sWLRealSource, 255);
	END

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iTBTableID
		AND role = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sTBRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sTBRealSource = left(@sTBRealSource, 255);
	END

	/* Initialise the insert strings with the basic values.*/
	SET @sColumnList = 'id_' + convert(varchar(128), @iEmpTableID) + ',' +
		'id_' + convert(varchar(128), @iCourseTableID) + ',' +
		@sTBStatusColumnName;
	SET @sValueList = convert(varchar(128), @piEmpRecordID) + ',' +
		convert(varchar(128), @piCourseRecordID) + ',' +
		'''' + @psStatus + '''';
	SET @sAddedColumns = ',' + convert(varchar(MAX), @iTBStatusColumnID) + ',';

	/* Get the TB and WL column permissions. */
	IF @fSysSecMgr = 0
	BEGIN
		CREATE TABLE #columnPermissions
		(
			columnID	int,
			action		int,		
			granted		bit		
		);

		INSERT INTO #columnPermissions
		SELECT 
			ASRSysColumns.columnID,
			sysprotects.action,
			CASE protectType
				WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END AS [protectType]
		FROM sysprotects
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
		INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
			AND (ASRSysColumns.tableID = @iTBTableID
				OR ASRSysColumns.tableID = @iWLTableID))
		WHERE sysprotects.uid = @iUserGroupID
			AND (sysprotects.action = 193 OR sysprotects.action = 197)
			AND (sysobjects.name = @sTBRealSource
				OR sysobjects.name = @sWLRealSource)
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
	END

	/* Get the Waiting List - Training Booking related columns. */
	DECLARE relatedColumns_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT sourceColumnID, destColumnID
	FROM ASRSysModuleRelatedColumns
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TBWLRelatedColumns';
	OPEN relatedColumns_cursor;
	FETCH NEXT FROM relatedColumns_cursor INTO @iSourceColumnID, @iDestinationColumnID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @iIndex = charindex(',' + convert(varchar(MAX), @iDestinationColumnID) + ',', @sAddedColumns);

		IF @iIndex = 0
		BEGIN
			/* Check that the user has read permission on the WL column, and update permission on the TB column. */
			SET @fGranted = 1;

			IF @fSysSecMgr = 0
			BEGIN
				SELECT @fGranted = granted
				FROM #columnPermissions
				WHERE columnID = @iDestinationColumnID
					AND action = 193;

				IF @fGranted IS null SET @fGranted = 0;

				IF @fGranted = 1
				BEGIN
					SELECT @fGranted = granted
					FROM #columnPermissions
					WHERE columnID = @iSourceColumnID
						AND action = 197;

					IF @fGranted IS null SET @fGranted = 0;
				END
			END

			IF @fGranted = 1
			BEGIN
				SELECT @sTempTBColumnName = columnName
				FROM ASRSysColumns
				WHERE columnID = @iSourceColumnID;

				SELECT @sTempWLColumnName = columnName
				FROM ASRSysColumns 
				WHERE columnID = @iDestinationColumnID;

				SET @sColumnList = @sColumnList + ',' + @sTempTBColumnName;
				SET @sValueList = @sValueList + ',' + @sTempWLColumnName;
			END
			
			SET @sAddedColumns = @sAddedColumns + convert(varchar(MAX), @iSourceColumnID) + ',';
		END

		FETCH NEXT FROM relatedColumns_cursor INTO @iSourceColumnID, @iDestinationColumnID;
	END
	CLOSE relatedColumns_cursor;
	DEALLOCATE relatedColumns_cursor;

	IF @fSysSecMgr = 0
	BEGIN
		/* Drop temporary tables no longer required. */
		DROP TABLE #columnPermissions;
	END

	/* Create the booking record. */
	SET @sExecString = 'INSERT INTO ' + @sTBRealSource + 
		'(' + @sColumnList + ')' +
		' SELECT TOP 1 ' + @sValueList + 
		' FROM ' + @sWLRealSource + 
		' WHERE ' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle,'''','''''') + '''' + 
		' AND id_' + convert(nvarchar(MAX), @iEmpTableID) + ' = ' + convert(nvarchar(MAX), @piEmpRecordID);
	EXEC sp_executesql @sExecString;

	/* Delete the old Waiting List record(s). */
	SET @sExecString = 'DELETE FROM ' + @sWLRealSource +
		' WHERE ' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle,'''','''''') + '''' + 
		' AND id_' + convert(nvarchar(MAX), @iEmpTableID) + ' = ' + convert(nvarchar(MAX), @piEmpRecordID);
	EXEC sp_executesql @sExecString;
END