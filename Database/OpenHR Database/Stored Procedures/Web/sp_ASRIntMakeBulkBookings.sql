CREATE PROCEDURE [dbo].[sp_ASRIntMakeBulkBookings] (
	@piCourseRecordID		integer,
	@psEmployeeRecordIDs	varchar(MAX),
	@psStatus				varchar(MAX)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUserGroupID		integer,
		@sUserGroupName			sysname,
		@iEmpTableID			integer,
		@iEmployeeID			integer,
		@iCourseTableID			integer,
		@iCourseTitleColumnID	integer,
		@sCourseTitleColumnName	sysname,
		@sCourseTitle			varchar(MAX),
		@sCourseSource			sysname,
		@iTBTableID				integer,
		@sTBTableName			sysname,
		@sTBRealSource			varchar(MAX),
		@iTBStatusColumnID		integer,
		@sTBStatusColumnName	sysname,
		@iWLTableID				integer,
		@sWLTableName			sysname,
		@sWLRealSource			varchar(MAX),
		@iWLCourseTitleColumnID	integer,
		@sWLCourseTitleColumnName	sysname,
		@iIndex					integer,
		@iChildViewID			integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@sActualUserName		sysname;

	/* Clean the input string parameters. */
	IF len(@psEmployeeRecordIDs) > 0 SET @psEmployeeRecordIDs = replace(@psEmployeeRecordIDs, '''', '''''');
	IF len(@psStatus) > 0 SET @psStatus = replace(@psStatus, '''', '''''');

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

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

	SELECT @sCourseTitleColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iCourseTitleColumnID;

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

	/* Get the WAITING LIST table information. */
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

	SELECT @sWLCourseTitleColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iWLCourseTitleColumnID;

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iWLTableID
		AND [role] = @sUserGroupName;
		
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
		AND [role] = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sTBRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sTBRealSource = left(@sTBRealSource, 255);
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
			' WHERE id = ' + convert(nvarchar(100), @piCourseRecordID);
		SET @sTempParamDefinition = N'@sCourseTitle varchar(MAX) OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sCourseTitle OUTPUT;

		FETCH NEXT FROM courseSourceCursor INTO @sCourseSource;
	END
	CLOSE courseSourceCursor;
	DEALLOCATE courseSourceCursor;

	WHILE len(@psEmployeeRecordIDs) > 0
	BEGIN
		/* Rip out the individual empoyee record ID from the given comma-delimited string of employee IDs. */
		SELECT @iIndex = charindex(',', @psEmployeeRecordIDs);
		IF @iIndex > 0
		BEGIN
			SET  @iEmployeeID = substring(@psEmployeeRecordIDs, 1, @iIndex - 1);
			SELECT @psEmployeeRecordIDs = substring(@psEmployeeRecordIDs, @iIndex + 1, len(@psEmployeeRecordIDs));
		END
		ELSE
		BEGIN
			SET  @iEmployeeID = @psEmployeeRecordIDs;
			SET @psEmployeeRecordIDs = '';
		END

		/* Create the new booking record. */
		SET @sTempExecString = 'INSERT INTO ' + @sTBRealSource + 
			' (' + @sTBStatusColumnName +
			', id_' + convert(nvarchar(100), @iEmpTableID) +
			', id_' + convert(nvarchar(100), @iCourseTableID) +
			') VALUES (''' + @psStatus + '''' +
			', ' + convert(nvarchar(100), @iEmployeeID) +
			', ' + convert(nvarchar(100), @piCourseRecordID) + ')';
		EXEC sp_executesql @sTempExecString;

		/* Remove any Waiting List records. */
		SET @sTempExecString = 'DELETE FROM ' + @sWLRealSource + 
			' WHERE id_' + convert(nvarchar(MAX), @iEmpTableID) + ' = ' + convert(nvarchar(100), @iEmployeeID) +
			' AND ' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle, '''', '''''') + '''';
		EXEC sp_executesql @sTempExecString;
	END
END