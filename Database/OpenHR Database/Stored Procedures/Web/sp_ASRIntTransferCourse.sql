CREATE PROCEDURE [dbo].[sp_ASRIntTransferCourse] (
	@piTBRecordID		integer,
	@piCourseRecordID	integer
)
AS
BEGIN
	DECLARE @iUserGroupID		integer,
		@sUserGroupName			sysname,
		@iEmpTableID			integer,
		@iEmpRecordID			integer,
		@iCourseTableID			integer,
		@iOriginalCourseRecordID	integer,
		@iTBTableID				integer,
		@sTBRealSource			varchar(MAX),
		@iTBStatusColumnID		integer,
		@sTBStatusColumnName	sysname,
		@iTBCancelDateColumnID	integer,
		@sTBCancelDateColumnName	sysname,
		@sBookingStatus			varchar(MAX),
		@fTStatusExists			bit,
		@iCount					integer,
		@iChildViewID			integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@sTBTableName			sysname,
		@sActualUserName		sysname;

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	/* Get the EMPLOYEE table information. */
	SELECT @iEmpTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable'
	IF @iEmpTableID IS NULL SET @iEmpTableID = 0

	/* Get the COURSE table information. */
	SELECT @iCourseTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTable'
	IF @iCourseTableID IS NULL SET @iCourseTableID = 0

	/* Get the TRAINING BOOKING table information. */
	SELECT @iTBTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookTable'
	IF @iTBTableID IS NULL SET @iTBTableID = 0

	SELECT @sTBTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iTBTableID

	SELECT @iTBStatusColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookStatus'
	IF @iTBStatusColumnID IS NULL SET @iTBStatusColumnID = 0

	SELECT @sTBStatusColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iTBStatusColumnID

	SELECT @iTBCancelDateColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookCancelDate'
	IF @iTBCancelDateColumnID IS NULL SET @iTBCancelDateColumnID = 0

	SELECT @sTBCancelDateColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @iTBCancelDateColumnID

	/* Check if the 'T' status code exists. */
	SET @fTStatusExists = 0
	SELECT @iCount = count(value)
	FROM ASRSysColumnControlValues
	WHERE columnID = @iTBStatusColumnID
		AND value = 'T'
	IF @iCount > 0 SET @fTStatusExists = 1

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iTBTableID
		AND role = @sUserGroupName
		
	IF @iChildViewID IS null SET @iChildViewID = 0
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sTBRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_')
		SET @sTBRealSource = left(@sTBRealSource, 255)
	END

	SET @sTempExecString = 'SELECT @iEmpRecordID = ID_' + convert(nvarchar(100), @iEmpTableID) +
		', @iOriginalCourseRecordID = ID_' + convert(nvarchar(100), @iCourseTableID) +
		', @sBookingStatus = ' + @sTBStatusColumnName +
		' FROM ' + @sTBRealSource +
		' WHERE id = ' + convert(nvarchar(100), @piTBRecordID)
	SET @sTempParamDefinition = N'@iEmpRecordID integer OUTPUT, @iOriginalCourseRecordID integer OUTPUT, @sBookingStatus varchar(MAX) OUTPUT'
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iEmpRecordID OUTPUT, @iOriginalCourseRecordID OUTPUT, @sBookingStatus OUTPUT


	IF @iEmpRecordID IS null SET @iEmpRecordID = 0
	IF @iOriginalCourseRecordID IS null SET @iOriginalCourseRecordID = 0

	/* Create the new booking record. */
	SET @sTempExecString = 'INSERT INTO ' + @sTBRealSource + 
		' (' + @sTBStatusColumnName +
		', id_' + convert(nvarchar(100), @iEmpTableID) +
		', id_' + convert(nvarchar(100), @iCourseTableID) +
		') VALUES (''' + @sBookingStatus + '''' +
		', ' + convert(nvarchar(100), @iEmpRecordID) +
		', ' + convert(nvarchar(100), @piCourseRecordID) + ')'
	EXEC sp_executesql @sTempExecString

	/* Update the old booking record. */     
	SET @sTempExecString = 'UPDATE ' + @sTBRealSource + 
		' SET ' + @sTBStatusColumnName + ' = ' + 
		CASE @fTStatusExists
			WHEN 1 THEN '''T'''
			ELSE '''C'''
		END

	IF len(@sTBCancelDateColumnName) > 0 
	BEGIN
		SET @sTempExecString = @sTempExecString +
			', ' + @sTBCancelDateColumnName + ' = getdate()'
	END

	SET @sTempExecString = @sTempExecString +
		' WHERE id = ' + convert(nvarchar(100), @piTBRecordID)

	EXEC sp_executesql @sTempExecString
END







GO

