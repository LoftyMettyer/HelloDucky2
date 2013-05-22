CREATE PROCEDURE [dbo].[sp_ASRIntGetTrainingBookingParameters] (
	@piEmployeeTableID			integer	OUTPUT,
	@piCourseTableID			integer	OUTPUT,
	@piCourseCancelDateColumnID	integer	OUTPUT,
	@piTBTableID				integer	OUTPUT,
	@pfTBTableSelect			bit		OUTPUT,
	@pfTBTableInsert			bit		OUTPUT,
	@pfTBTableUpdate			bit		OUTPUT,
	@piTBStatusColumnID			integer	OUTPUT,
	@pfTBStatusColumnUpdate		bit		OUTPUT,
	@piTBCancelDateColumnID		integer	OUTPUT,
	@pfTBCancelDateColumnUpdate	bit		OUTPUT,
	@pfTBProvisionalStatusExists	bit	OUTPUT,
	@piWaitListTableID			integer	OUTPUT,
	@pfWaitListTableInsert			bit	OUTPUT,
	@pfWaitListTableDelete			bit	OUTPUT,
	@piWaitListCourseTitleColumnID		integer	OUTPUT,
	@pfWaitListCourseTitleColumnUpdate	bit	OUTPUT,
	@pfWaitListCourseTitleColumnSelect	bit	OUTPUT,
	@piBulkBookingDefaultViewID		integer	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the given screen's definition and table permission info. */
	DECLARE @fOK			bit,
		@fSysSecMgr			bit,
		@iUserGroupID		integer,
		@sUserGroupName		sysname,
		@sTempName			sysname,
		@iTempAction		integer,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@sRealSource		sysname,
		@iStatusCount		integer,
		@iChildViewID		integer,
		@sTBTableName		sysname,
		@sWLTableName		sysname,
		@sActualUserName	sysname;
		
	/* Training Booking information. */
	SET @fOK = 1

	SET @piEmployeeTableID = 0

	SET @piCourseTableID = 0
	SET @piCourseCancelDateColumnID = 0

	SET @piTBTableID = 0
	SET @pfTBTableSelect = 0
	SET @pfTBTableInsert = 0
	SET @pfTBTableUpdate = 0
	SET @piTBStatusColumnID = 0
	SET @pfTBStatusColumnUpdate = 0
	SET @piTBCancelDateColumnID = 0
	SET @pfTBCancelDateColumnUpdate = 0
	SET @pfTBProvisionalStatusExists = 0

	SET @piWaitListTableID = 0
	SET @pfWaitListTableInsert = 0
	SET @pfWaitListTableDelete = 0
	SET @piWaitListCourseTitleColumnID = 0
	SET @pfWaitListCourseTitleColumnUpdate = 0
	SET @pfWaitListCourseTitleColumnSelect = 0

	SET @piBulkBookingDefaultViewID = 0
	
	/* Get the current user's group id. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

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

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'TRAINING', @fOK OUTPUT

	/* Get the required training booking module paramaters. */
	IF @fOK = 1
	BEGIN
		/* Get the EMPLOYEE table information. */
		SELECT @piEmployeeTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_EmployeeTable'
		IF @piEmployeeTableID IS NULL SET @piEmployeeTableID = 0

		/* Get the COURSE table information. */
		SELECT @piCourseTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseTable'
		IF @piCourseTableID IS NULL SET @piCourseTableID = 0

		IF @piCourseTableID > 0
		BEGIN
			SELECT @piCourseCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_CourseCancelDate'
			IF @piCourseCancelDateColumnID IS NULL SET @piCourseCancelDateColumnID = 0
		END

		/* Get the TRAINING BOOKING table information. */
		SELECT @piTBTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_TrainBookTable'
		IF @piTBTableID IS NULL SET @piTBTableID = 0


		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID]				int,
								   [columns]		varbinary(8000),
								   [Action]			tinyint,
								   [ProtectType]	tinyint)
		INSERT INTO @SysProtects
		SELECT [ID],[Columns],[Action],[ProtectType] FROM #sysprotects
			WHERE [Action] IN (193, 195, 196, 197)

		IF @piTBTableID > 0
		BEGIN
			SELECT @sTBTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piTBTableID

			SELECT @piTBStatusColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookStatus'
			IF @piTBStatusColumnID IS NULL SET @piTBStatusColumnID = 0

			SELECT @piTBCancelDateColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_TrainBookCancelDate'
			IF @piTBCancelDateColumnID IS NULL SET @piTBCancelDateColumnID = 0

			SET @sCommand = 'SELECT @iStatusCount = COUNT(*)' +
				' FROM ASRSysColumnControlValues' +
				' WHERE columnID = ' + convert(nvarchar(100), @piTBStatusColumnID) +
				' AND value = ''P'''
			SET @sParamDefinition = N'@iStatusCount integer OUTPUT'
			EXEC sp_executesql @sCommand, @sParamDefinition, @iStatusCount OUTPUT
			IF @iStatusCount > 0 SET @pfTBProvisionalStatusExists = 1

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				/* System/Security managers must have all permissions granted. */
				SET @pfTBTableSelect = 1
				SET @pfTBTableInsert = 1
				SET @pfTBTableUpdate = 1
				SET @pfTBStatusColumnUpdate = 1
				SET @pfTBCancelDateColumnUpdate = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piTBTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN

					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sTBTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (193, 195, 197)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN

						IF @iTempAction = 193
						BEGIN
							SET @pfTBTableSelect = 1
						END
						IF @iTempAction = 195
						BEGIN
							SET @pfTBTableInsert = 1
						END
						IF @iTempAction = 197
						BEGIN
							SET @pfTBTableUpdate = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @tbColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @tbColumnPermissions
					SELECT 
						ASRSysColumns.columnID,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piTBTableID
						AND (ASRSysColumns.columnID = @piTBStatusColumnID
							OR ASRSysColumns.columnID = @piTBCancelDateColumnID))
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfTBStatusColumnUpdate = granted

					FROM @tbColumnPermissions
					WHERE columnID = @piTBStatusColumnID
						AND action = 197
					IF @pfTBStatusColumnUpdate IS NULL SET @pfTBStatusColumnUpdate = 0

					SELECT @pfTBCancelDateColumnUpdate = granted
					FROM @tbColumnPermissions
					WHERE columnID = @piTBCancelDateColumnID
						AND action = 197
					IF @pfTBCancelDateColumnUpdate IS NULL SET @pfTBCancelDateColumnUpdate = 0

				END
			END
		END

		/* Get the waiting list table information. */
		SELECT @piWaitListTableID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_WaitListTable'
		IF @piWaitListTableID IS NULL SET @piWaitListTableID = 0

		IF @piWaitListTableID > 0
		BEGIN
			SELECT @sWLTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piWaitListTableID

			SELECT @piWaitListCourseTitleColumnID = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_WaitListCourseTitle'
			IF @piWaitListCourseTitleColumnID IS NULL SET @piWaitListCourseTitleColumnID = 0

			/* Check what permissions the current user has on the table. */
			IF @fSysSecMgr = 1
			BEGIN
				SET @pfWaitListTableInsert = 1
				SET @pfWaitListTableDelete = 1
				SET @pfWaitListCourseTitleColumnUpdate = 1
				SET @pfWaitListCourseTitleColumnSelect = 1
			END
			ELSE
			BEGIN
				SET @sRealSource = ''

				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @piWaitListTableID
					AND role = @sUserGroupName
					
				IF @iChildViewID IS null SET @iChildViewID = 0
					
				IF @iChildViewID > 0 
				BEGIN
					SET @sRealSource = 'ASRSysCV' + 
						convert(varchar(1000), @iChildViewID) +
						'#' + replace(@sWLTableName, ' ', '_') +
						'#' + replace(@sUserGroupName, ' ', '_')
					SET @sRealSource = left(@sRealSource, 255)

					DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT sysobjects.name, p.action
						FROM @SysProtects p
						INNER JOIN sysobjects ON p.id = sysobjects.id
						WHERE p.protectType <> 206
							AND p.action IN (195, 196)
							AND sysobjects.name = @sRealSource

					OPEN tableInfo_cursor
					FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iTempAction = 195
						BEGIN
							SET @pfWaitListTableInsert = 1
						END
						IF @iTempAction = 196
						BEGIN
							SET @pfWaitListTableDelete = 1
						END
						FETCH NEXT FROM tableInfo_cursor INTO @sTempName, @iTempAction

					END
					CLOSE tableInfo_cursor
					DEALLOCATE tableInfo_cursor
				END

				IF LEN(@sRealSource) > 0
				BEGIN
					/* Check the current user's column permissions. */
					/* Create a temporary table of the column permissions. */
					DECLARE @waitListColumnPermissions TABLE
					(
						columnID	int,
						action		int,		
						granted		bit		
					)

					INSERT INTO @waitListColumnPermissions
					SELECT 
						ASRSysColumns.columnID,
						p.action,
						CASE protectType
							WHEN 205 THEN 1
							WHEN 204 THEN 1
							ELSE 0
						END 
					FROM @SysProtects p
					INNER JOIN sysobjects ON p.id = sysobjects.id
					INNER JOIN syscolumns ON p.id = syscolumns.id
					INNER JOIN ASRSysColumns ON (syscolumns.name = ASRSysColumns.columnName
						AND ASRSysColumns.tableID = @piWaitListTableID
						AND ASRSysColumns.columnID = @piWaitListCourseTitleColumnID)
					WHERE p.action IN (193, 197)
						AND sysobjects.name = @sRealSource
						AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
						OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
						AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))

					SELECT @pfWaitListCourseTitleColumnUpdate = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 197
					IF @pfWaitListCourseTitleColumnUpdate IS NULL SET @pfWaitListCourseTitleColumnUpdate = 0

					SELECT @pfWaitListCourseTitleColumnSelect = granted
					FROM @waitListColumnPermissions
					WHERE columnID =  @piWaitListCourseTitleColumnID
						AND action = 193
					IF @pfWaitListCourseTitleColumnSelect IS NULL SET @pfWaitListCourseTitleColumnSelect = 0

				END
			END
		END

		/* Get the Bulk Booking default view. */
		SELECT @piBulkBookingDefaultViewID = convert(integer, parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_BulkBookingDefaultView'
		IF @piBulkBookingDefaultViewID IS NULL SET @piBulkBookingDefaultViewID = 0
	END
END











GO

