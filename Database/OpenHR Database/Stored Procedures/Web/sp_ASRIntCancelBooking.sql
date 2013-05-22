CREATE PROCEDURE [dbo].[sp_ASRIntCancelBooking] (
	@pfTransferBookings	bit,
	@piTBRecordID		integer,
	@psErrorMessage		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	
		@iCount						integer,
		@iUserGroupID				integer,
		@sUserGroupName				sysname,
		@fSysSecMgr					bit,
		@iTBTableID					integer,
		@sTBTableName				sysname,
		@sTBRealSource				sysname,
		@iWLTableID					integer,
		@sWLTableName				sysname,
		@sWLRealSource				sysname,
		@iChildViewID				integer,
		@iTBStatusColumnID			integer,
		@sTBStatusColumnName		sysname,
		@sExecString				nvarchar(MAX),
		@sCommand					nvarchar(MAX),
		@sParamDefinition			nvarchar(500),
		@sTBStatus					varchar(MAX),
		@iEmpID						integer,
		@iCourseID					integer,
		@iEmpTableID				integer,
		@iCourseTableID				integer,
		@iStatusCount				integer,
		@fTBProvisionalStatusExists		bit,
		@iCourseTitleColumnID		integer,
		@sCourseTitleColumnName		sysname,
		@sTempExecString			nvarchar(MAX),
		@sTempParamDefinition		nvarchar(500),
		@sCourseTitle				varchar(MAX),
		@iTBCancelDateColumnID		integer,
		@fTBCancelDateColumnUpdate	bit,
		@sTBCancelDateColumnName	sysname,
		@iWLCourseTitleColumnID		integer,
		@sWLCourseTitleColumnName	sysname,
		@sColumnList				varchar(MAX),
		@sValueList					varchar(MAX),
		@sAddedColumns				varchar(MAX),
		@sCourseSource				sysname,
		@iSourceColumnID			integer, 
		@iDestinationColumnID		integer,
		@iIndex						integer,
		@fGranted					bit,
		@sTempTBColumnName			sysname,
		@sTempWLColumnName			sysname,
		@sActualUserName			sysname;

	SET @psErrorMessage = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

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

	/* Get the EMPLOYEE table information. */
	SELECT @iEmpTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable'
	IF @iEmpTableID IS NULL SET @iEmpTableID = 0;

	/* Get the COURSE table information. */
	SELECT @iCourseTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTable'
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
		AND parameterKey = 'Param_TrainBookStatus'
	IF @iTBStatusColumnID IS NULL SET @iTBStatusColumnID = 0;

	/* Get the training booking status column name. */
	SELECT @sTBStatusColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iTBStatusColumnID;

	SET @fTBProvisionalStatusExists = 0
	SET @sCommand = 'SELECT @iStatusCount = COUNT(*)' +
		' FROM ASRSysColumnControlValues' +
		' WHERE columnID = ' + convert(nvarchar(255), @iTBStatusColumnID) +
		' AND value = ''P''';
	SET @sParamDefinition = N'@iStatusCount integer OUTPUT';
	EXEC sp_executesql @sCommand, @sParamDefinition, @iStatusCount OUTPUT;
	IF @iStatusCount > 0 SET @fTBProvisionalStatusExists = 1;

	SELECT @iTBCancelDateColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookCancelDate';
	IF @iTBCancelDateColumnID IS NULL SET @iTBCancelDateColumnID = 0;

	IF @iTBCancelDateColumnID > 0 
	BEGIN
		SELECT @sTBCancelDateColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iTBCancelDateColumnID;
	END
	IF @sTBCancelDateColumnName IS NULL SET @sTBCancelDateColumnName = '';

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

	/* Get the realSource of the training booking table. */
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

	/* Get the status, employee ID and course ID from the given TB record. */
	/* NB. If we've reached this point we already know that we have 'read' permision on the Trining Booking 'status' and id columns. */	
	SET @sCommand = 'SELECT @sTBStatus = ' + @sTBStatusColumnName +
		', @iEmpID = id_' + convert(nvarchar(255), @iEmpTableID) +
		', @iCourseID = id_' + convert(nvarchar(255), @iCourseTableID) +
		' FROM ' + @sTBRealSource +
		' WHERE id = ' + convert(varchar(100), @piTBRecordID);

	SET @sParamDefinition = N'@sTBStatus varchar(MAX) OUTPUT, @iEmpID integer OUTPUT, @iCourseID integer OUTPUT';
	EXEC sp_executesql @sCommand, @sParamDefinition, @sTBStatus OUTPUT, @iEmpID OUTPUT, @iCourseID OUTPUT;

	/* Check the employee ID, course ID and status are valid. */
	IF (@sTBStatus <> 'B') AND (@sTBStatus <> 'P') 
	BEGIN
		SET @psErrorMessage = 'Bookings can only be cancelled if they have ''Booked''';

		IF @fTBProvisionalStatusExists = 1
		BEGIN
			SET @psErrorMessage = @psErrorMessage + ' or ''Provisional''';
		END
		SET @psErrorMessage = @psErrorMessage + ' status.';
		RETURN;
	END

	IF NOT (@iEmpID > 0) 
	BEGIN
		SET @psErrorMessage = 'The selected Training Booking record has no associated Employee record.';
		RETURN;
	END

	IF NOT (@iCourseID > 0) 
	BEGIN
		SET @psErrorMessage = 'The selected Training Booking record has no associated Course record.';
		RETURN;
	END

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
			' WHERE id = ' + convert(nvarchar(255), @iCourseID);
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

	/* Check the current user's column permissions on the current course table/view. */
	SET @fTBCancelDateColumnUpdate = 0;

	IF @fSysSecMgr = 1
	BEGIN
		SET @fTBCancelDateColumnUpdate = 1;
	END
	ELSE
	BEGIN
	
		/* Create a temporary table of the column permissions. */
		DECLARE @tbColumnPermissions TABLE(
			columnID	int,
			action		int,		
			granted		bit);

		INSERT INTO @tbColumnPermissions
		SELECT 
			ASRSysColumns.columnID,
			sysprotects.action,
			CASE protectType
				WHEN 205 THEN 1
				WHEN 204 THEN 1
				ELSE 0
			END 
		FROM sysprotects
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
		INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
			AND ASRSysColumns.tableID = @iTBTableID
			AND ASRSysColumns.columnID = @iTBCancelDateColumnID)
		WHERE sysprotects.uid = @iUserGroupID
			AND (sysprotects.action = 197)
			AND sysobjects.name = @sTBRealSource
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		SELECT @fTBCancelDateColumnUpdate = granted
		FROM @tbColumnPermissions
		WHERE columnID =  @iTBCancelDateColumnID;
		IF @fTBCancelDateColumnUpdate IS NULL SET @fTBCancelDateColumnUpdate = 0;

	END

	/* Update the TrainingBooking record. */
	/* Already checked that we have 'update' permission on the ststaus column. */
	SET @sCommand = 'UPDATE ' + @sTBRealSource +
		' SET ' + @sTBStatusColumnName + ' = ''C''';

	IF (@iTBCancelDateColumnID > 0) AND (@fTBCancelDateColumnUpdate = 1)
	BEGIN
		/* Add the 'cancel date' column to the update string if the user has permission to. */
		SET @sCommand = @sCommand +
			', ' + @sTBCancelDateColumnName + ' = getdate()';
	END

	SET @sCommand = @sCommand +
		' WHERE id = ' + convert(varchar(100), @piTBRecordID);
	SET @sParamDefinition = N'@sTBStatus varchar(MAX) OUTPUT, @iEmpID integer OUTPUT, @iCourseID integer OUTPUT';
	EXEC sp_executesql @sCommand, @sParamDefinition, @sTBStatus OUTPUT, @iEmpID OUTPUT, @iCourseID OUTPUT;

	/* Create Waiting List record if required. */
	IF @pfTransferBookings = 1
	BEGIN
		/* Check if there is already a WL record for the course. */
		SET @sCommand = 'SELECT @iCount = COUNT(*)' +
			' FROM ' + @sWLRealSource + 
			' WHERE ' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle, '''', '''''') + '''' + 
			' AND id_' + convert(nvarchar(255), @iEmpTableID) + ' = ' + convert(nvarchar(MAX), @iEmpID);
		SET @sParamDefinition = N'@iCount integer OUTPUT';
		EXEC sp_executesql @sCommand, @sParamDefinition, @iCount OUTPUT;

		IF @iCount = 0
		BEGIN
			/* Initialise the insert strings with the basic values.*/
			/* NB. To reach this point we've already checked the user has 'update' permission on the 'courseTitle' column in the Waiting List table. */
			SET @sColumnList = 'id_' + convert(nvarchar(255), @iEmpTableID) + ',' +	@sWLCourseTitleColumnName;
			SET @sValueList = convert(nvarchar(255), @iEmpID) + ',' +	'''' + replace(@sCourseTitle, '''', '''''') + '''';
			SET @sAddedColumns = ',' + convert(varchar(255), @iWLCourseTitleColumnID) + ',';

			/* Get the TB and WL column permissions. */
			IF @fSysSecMgr = 0
			BEGIN
				DECLARE @columnPermissions TABLE(
					columnID	int,
					[action]	int,		
					granted		bit);

				INSERT INTO @columnPermissions
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
				SET @iIndex = charindex(',' + convert(varchar(255), @iDestinationColumnID) + ',', @sAddedColumns);

				IF @iIndex = 0
				BEGIN
					/* Check that the user has read permission on the WL column, and update permission on the TB column. */
					SET @fGranted = 1;

					IF @fSysSecMgr = 0
					BEGIN
						SELECT @fGranted = granted
						FROM @columnPermissions
						WHERE columnID = @iSourceColumnID
							AND [action] = 193;

						IF @fGranted IS null SET @fGranted = 0

						IF @fGranted = 1
						BEGIN
							SELECT @fGranted = granted
							FROM @columnPermissions
							WHERE columnID = @iDestinationColumnID
								AND [action] = 197;

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

						SET @sColumnList = @sColumnList + ',' + @sTempWLColumnName;
						SET @sValueList = @sValueList + ',' + @sTempTBColumnName;
					END
			
					SET @sAddedColumns = @sAddedColumns + convert(varchar(255), @iSourceColumnID) + ',';
				END

				FETCH NEXT FROM relatedColumns_cursor INTO @iSourceColumnID, @iDestinationColumnID;
			END
			CLOSE relatedColumns_cursor;
			DEALLOCATE relatedColumns_cursor;

			/* Create the WL record. */
			SET @sExecString = 'INSERT INTO ' + @sWLRealSource + 
				'(' + @sColumnList + ')' +
				' SELECT TOP 1 ' + @sValueList + 
				' FROM ' + @sTBRealSource + 
				' WHERE id = ' + convert(nvarchar(255), @piTBRecordID);
			EXEC sp_executesql @sExecString;
		END
	END
END