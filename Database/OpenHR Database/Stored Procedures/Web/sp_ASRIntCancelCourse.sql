CREATE PROCEDURE [dbo].[sp_ASRIntCancelCourse] (
	@piNumberOfBookings			integer	OUTPUT,
	@piCourseRecordID			integer,
	@piTrainBookTableID			integer,
	@piCourseTableID			integer,
	@piTrainBookStatusColumnID	integer,
	@psCourseRealSource			varchar(MAX),
	@psErrorMessage				varchar(MAX) OUTPUT,
	@psCourseTitle				varchar(MAX) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sCommand					nvarchar(MAX),
			@sParamDefinition			nvarchar(500),
			@sRealSource				sysname,
			@iCourseTableID				integer,
			@fTransferProvisionals		bit,
			@sTrainBookStatusColumnName	sysname,
			@iUserGroupID				integer,
			@sUserGroupName				sysname,
			@fSysSecMgr					bit,
			@iChildViewID				integer,
			@sTemp						varchar(MAX),
			@iCourseTitleColumnID			integer,
			@fCourseTitleColumnSelect		bit,
			@iCourseCancelDateColumnID		integer,
			@fCourseCancelDateColumnUpdate	bit,
			@iCourseCancelByColumnID		integer,
			@fCourseCancelByColumnUpdate	bit,
			@sCourseTitleColumnName		sysname,
			@sTBTableName				sysname,
			@sActualUserName			sysname,
			@sCleanCourseRealSource		sysname;

	/* Clean the input string parameters. */
	SET @sCleanCourseRealSource = @psCourseRealSource;
	IF len(@sCleanCourseRealSource) > 0 SET @sCleanCourseRealSource = replace(@sCleanCourseRealSource, '''', '''''');

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

	SELECT @sTBTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @piTrainBookTableID;
	
	/* Get the realSource of the training booking table. */
	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @piTrainBookTableID
		AND [role] = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sRealSource = left(@sRealSource, 255);
	END

	/* Check if we need to transfer provisional bookings. */
	SET @sTemp = '';
	SELECT @sTemp = convert(varchar(MAX), parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTransferProvisionals';
		
	IF @sTemp IS NULL SET @sTemp = '';
	IF @sTemp = 'TRUE'
	BEGIN
		SET @fTransferProvisionals = 1;
	END
	ELSE
	BEGIN
		SET @fTransferProvisionals = 0;
	END

	/* Get the training booking status column name. */
	SELECT @sTrainBookStatusColumnName = columnName
	FROM ASRSysColumns
	WHERE columnID = @piTrainBookStatusColumnID;

	/* Get the number of training booking records for the current course. */
	SET @sCommand = 'SELECT @iValue = COUNT(ID) ' + 
		' FROM ' + @sRealSource +
		' WHERE id_' + convert(varchar(100), @piCourseTableID) + ' = ' + convert(varchar(100), @piCourseRecordID);

	IF @fTransferProvisionals = 1 
	BEGIN
		SET @sCommand = @sCommand +
			' AND (LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''' +
			' OR LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''P'')';
	END
	ELSE
	BEGIN
		SET @sCommand = @sCommand +
			' AND LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''';
	END

	SET @sParamDefinition = N'@iValue integer OUTPUT';
	EXEC sp_executesql @sCommand, @sParamDefinition, @piNumberOfBookings OUTPUT;

	IF @piNumberOfBookings IS NULL SET @piNumberOfBookings = 0;

	/* Check the current user's column permissions on the current course table/view. */
	/* Get the IDs of the required columns. */
	SELECT @iCourseTitleColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTitle';
	IF @iCourseTitleColumnID IS NULL SET @iCourseTitleColumnID = 0;

	SELECT @iCourseCancelDateColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseCancelDate';
	IF @iCourseCancelDateColumnID IS NULL SET @iCourseCancelDateColumnID = 0;

	SELECT @iCourseCancelByColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseCancelledBy';
	IF @iCourseCancelByColumnID IS NULL SET @iCourseCancelByColumnID = 0;

	IF @fSysSecMgr = 1
	BEGIN
		SET @fCourseTitleColumnSelect = 1;
		SET @fCourseCancelDateColumnUpdate = 1;
		SET @fCourseCancelByColumnUpdate = 1;
	END
	ELSE
	BEGIN
		/* Create a temporary table of the column permissions. */
		DECLARE @courseColumnPermissions TABLE(
			columnID	int,
			[action]		int,		
			granted		bit);

		INSERT INTO @courseColumnPermissions
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
			AND ASRSysColumns.tableID = @piCourseTableID
			AND ((ASRSysColumns.columnID = @iCourseTitleColumnID) 
				OR (ASRSysColumns.columnID = @iCourseCancelDateColumnID)
				OR (ASRSysColumns.columnID = @iCourseCancelByColumnID)))
		WHERE sysprotects.uid = @iUserGroupID
			AND (sysprotects.action = 193 OR sysprotects.action = 197)
			AND sysobjects.name = @psCourseRealSource
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		SELECT @fCourseTitleColumnSelect = granted
		FROM @courseColumnPermissions
		WHERE columnID =  @iCourseTitleColumnID
			AND [action] = 193;
		IF @fCourseTitleColumnSelect IS NULL SET @fCourseTitleColumnSelect = 0;

		SELECT @fCourseCancelDateColumnUpdate = granted
		FROM @courseColumnPermissions
		WHERE columnID =  @iCourseCancelDateColumnID
			AND action = 197;
		IF @fCourseCancelDateColumnUpdate IS NULL SET @fCourseCancelDateColumnUpdate = 0;

		SELECT @fCourseCancelByColumnUpdate = granted
		FROM @courseColumnPermissions
		WHERE columnID =  @iCourseCancelByColumnID
			AND action = 197;
		IF @fCourseCancelByColumnUpdate IS NULL SET @fCourseCancelByColumnUpdate = 0;

	END

	IF @iCourseTitleColumnID = 0 SET @psErrorMessage = 'Unable to find the Course Title column.';
	IF ((LEN(@psErrorMessage) = 0) AND (@fCourseTitleColumnSelect = 0)) SET @psErrorMessage = 'You do not have ''read'' permission on the Course Title column in the current table/view.';
	IF ((LEN(@psErrorMessage) = 0) AND (@iCourseCancelDateColumnID = 0)) SET @psErrorMessage = 'Unable to find the Course Cancel Date column.';
	IF ((LEN(@psErrorMessage) = 0) AND (@fCourseCancelDateColumnUpdate = 0)) SET @psErrorMessage = 'You do not have ''edit'' permission on the Course Cancel Date column in the current table/view.';
	IF ((LEN(@psErrorMessage) = 0) AND (@iCourseCancelByColumnID > 0) AND (@fCourseCancelByColumnUpdate = 0)) SET @psErrorMessage = 'You do not have ''edit'' permission on the Course Cancel By column in the current table/view.';

	SET @psCourseTitle = '';
	IF (@iCourseTitleColumnID > 0) AND (@fCourseTitleColumnSelect = 1)
	BEGIN
		SELECT @sCourseTitleColumnName = columnName
		FROM [dbo].[ASRSysColumns]
		WHERE columnID = @iCourseTitleColumnID;

		IF @sCourseTitleColumnName IS NULL SET @sCourseTitleColumnName = '';
		IF LEN(@sCourseTitleColumnName) > 0
		BEGIN
			SET @sCommand = 'SELECT @sValue = ' + @sCourseTitleColumnName +
				' FROM ' + @sCleanCourseRealSource +
				' WHERE id = ' + convert(varchar(100), @piCourseRecordID);

			SET @sParamDefinition = N'@sValue varchar(MAX) OUTPUT';
			EXEC sp_executesql @sCommand, @sParamDefinition, @psCourseTitle OUTPUT;
			IF @psCourseTitle IS NULL SET @psCourseTitle = '';
		END
	END
END