CREATE PROCEDURE [dbo].[sp_ASRIntCancelCoursePart2] (
	@piEmployeeTableID					integer,
	@piCourseTableID					integer,
	@psCourseRealSource					varchar(MAX),
	@piCourseRecordID					integer,
	@piTransferCourseRecordID			integer,
	@piCourseCancelDateColumnID			integer,
	@psCourseTitle						varchar(MAX),
	@piTrainBookTableID					integer,
	@pfTrainBookTableInsert				bit,
	@piTrainBookStatusColumnID			integer,
	@piTrainBookCancelDateColumnID		integer,
	@piWaitListTableID					integer,
	@pfWaitListTableInsert				bit,
	@piWaitListCourseTitleColumnID		integer,
	@pfWaitListCourseTitleColumnUpdate	bit,
	@pfCreateWaitListRecords			bit,
	@psErrorMessage						varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* This stored procedure actually does the course cancellation, and the associated transferrals, etc. */
	/* Returns an error string if anything went wrong. */
	DECLARE	@sCommand						nvarchar(MAX),
			@sParamDefinition				nvarchar(500),
			@iUserGroupID					integer,
			@sUserGroupName					sysname,
			@iChildViewID					integer,
			@sTemp							varchar(MAX),
			@iCount							integer,
			@fTransferProvisionals			bit,
			@fCStatusExists					bit,
			@fCCStatusExists				bit,
			@sCourseCancelDateColumnName	sysname,
			@iCourseCancelledByColumnID		integer,
			@sCourseCancelledByColumnName	sysname,
			@sWLTableName					sysname,
			@sWaitListRealSource			sysname,
			@sWaitListCourseTitleColumnName	sysname,
			@sTBTableName					sysname,
			@sTrainBookRealSource			sysname,
			@sTrainBookStatusColumnName		sysname,
			@sTrainBookCancelDateColumnName	sysname,
			@sActualUserName				sysname,
			@iSourceColumnID				integer,
			@iDestinationColumnID			integer,
			@sAddedColumns					varchar(MAX),
			@iIndex							integer,
			@fGranted						bit,
			@sTempTBColumnName				sysname,
			@sTempWLColumnName				sysname,
			@sColumnList					varchar(MAX),
			@sValueList						varchar(MAX),
			@fSysSecMgr						bit;

	BEGIN TRANSACTION

	/* Clean the input string parameters. */
	IF len(@psCourseRealSource) > 0 SET @psCourseRealSource = replace(@psCourseRealSource, '''', '''''');
	IF len(@psCourseTitle) > 0 SET @psCourseTitle = replace(@psCourseTitle, '''', '''''');

	SET @psErrorMessage = '';

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

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
		SET @sTrainBookRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sTrainBookRealSource = left(@sTrainBookRealSource, 255);
	END
	ELSE
	BEGIN
		SET @psErrorMessage = 'Unable to determine the Training Booking child view.';
	END

	IF LEN(@psErrorMessage) = 0 
	BEGIN
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

		/* Get the Course Cancelled Date column name if there is a column defined. */
		IF @piCourseCancelDateColumnID > 0 
		BEGIN
			SELECT @sCourseCancelDateColumnName = columnName
			FROM ASRSysColumns
			WHERE columnID = @piCourseCancelDateColumnID;
		END

		IF @sCourseCancelDateColumnName IS NULL SET @sCourseCancelDateColumnName = '';
		IF LEN(@sCourseCancelDateColumnName) = 0 SET @psErrorMessage = 'Unable to find the Course Cancel Date column.';
	END

	IF LEN(@psErrorMessage) = 0
	BEGIN
		/* Get the Course Cancelled By column name if there is a column defined. */
		SELECT @iCourseCancelledByColumnID = convert(integer, parameterValue)
		FROM [dbo].[ASRSysModuleSetup]
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseCancelledBy';
			
		IF @iCourseCancelledByColumnID IS NULL SET @iCourseCancelledByColumnID = 0;
		IF @iCourseCancelledByColumnID > 0
		BEGIN
			SELECT @sCourseCancelledByColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @iCourseCancelledByColumnID;
			
			IF @sCourseCancelledByColumnName IS NULL SET @sCourseCancelledByColumnName = '';
			IF LEN(@sCourseCancelledByColumnName) = 0 SET @psErrorMessage = 'Unable to find the Course Cancel Date column.';
		END
	END

	IF LEN(@psErrorMessage) = 0 		
	BEGIN
		/* Update the current course record. */
		/* NB. The 'sp_ASRIntCancelCourse' stored procedure is run before this one, and checks certain permissions.
			If the user does not have update permission on the Course Cancelled Date column 
			or Course Cancelled By column (if there is one) then this point won't be reached. */
		SET @sCommand = 'UPDATE ' + @psCourseRealSource +
			' SET ' + @sCourseCancelDateColumnName + ' = getdate()';
		IF @iCourseCancelledByColumnID > 0 SET @sCommand = @sCommand + ', ' + @sCourseCancelledByColumnName + ' = SYSTEM_USER';
		SET @sCommand = @sCommand + ' WHERE id = ' + convert(nvarchar(255), @piCourseRecordID);
		EXEC sp_executesql @sCommand;
	END
	
	IF LEN(@psErrorMessage) = 0 		
	BEGIN	
		/* Get the Training Booking Status column name. */
		IF @piTrainBookStatusColumnID > 0
		BEGIN
			SELECT @sTrainBookStatusColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @piTrainBookStatusColumnID;
		END
		IF @sTrainBookStatusColumnName IS NULL SET @psErrorMessage = 'Unable to find the Training Booking Status column.';
	END

	IF LEN(@psErrorMessage) = 0 		
	BEGIN	
		/* Transfer course records if required. */
		IF @piTransferCourseRecordID > 0 
		BEGIN
			IF @pfTrainBookTableInsert = 0 SET @psErrorMessage = 'You do not have ''new'' permission on the Training Booking table.';
			IF LEN(@psErrorMessage) = 0 		
			BEGIN
				/* Create the transferred booking records. */
				/* NB. Insert permission on the table is checked above.
					The 'sp_ASRIntGetTrainingBookingParameters' stored procedure is run before this one, 
					as the user logs into the intranet module, and checks certain permissions. 
					If the user does not have update permission on the Status column 
					then this point won't be reached. 
				The checks for overbooking, unavailability, overlapped bookings and pre-requisites
				are made as the user selects the new course. */
				SET @sCommand = 'INSERT INTO ' + @sTrainBookRealSource +
					' (' + @sTrainBookStatusColumnName + ', ' +
					'id_' + convert(nvarchar(255), @piEmployeeTableID) + ', ' +
					'id_' + convert(nvarchar(255), @piCourseTableID) + ')' +
					' (SELECT ' +
						@sTrainBookStatusColumnName + ', ' +
						'id_' + convert(nvarchar(255), @piEmployeeTableID) + ', ' +
						convert(nvarchar(255), @piTransferCourseRecordID) +
						' FROM ' + @sTrainBookRealSource +
						' WHERE id_' + convert(nvarchar(255), @piCourseTableID) + ' = ' + convert(nvarchar(255), @piCourseRecordID);
				IF @fTransferProvisionals = 1
				BEGIN
					SET @sCommand = @sCommand +
						' AND (LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''' +
						' OR LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''P''))';
				END	
				ELSE
				BEGIN
					SET @sCommand = @sCommand +
						' AND LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''';
				END
				EXEC sp_executesql @sCommand;
			END
		END
	END

	IF (LEN(@psErrorMessage) = 0)
		AND (@piTrainBookCancelDateColumnID > 0)
	BEGIN
		/* Change the Cancellation Date of the existing bookings. */	
		/* NB. Update permission on the table is checked above.
			The 'sp_ASRIntGetTrainingBookingParameters' stored procedure is run before this one, 
				as the user logs into the intranet module, and checks certain permissions. 
				If the user does not have update permission on the Status column or Cancel Date column (if there is one)
				then this point won't be reached. */

		SELECT @sTrainBookCancelDateColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @piTrainBookCancelDateColumnID;
			
		SET @sCommand = 'UPDATE ' + @sTrainBookRealSource +
			' SET '+ @sTrainBookCancelDateColumnName + ' = getdate()' +
			' WHERE id_' + convert(nvarchar(255), @piCourseTableID) + ' = ' + convert(nvarchar(255), @piCourseRecordID);
       
		IF @fTransferProvisionals  = 1
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
		EXEC sp_executesql @sCommand;
	END

	IF (LEN(@psErrorMessage) = 0) AND (@piTransferCourseRecordID = 0) AND (@pfCreateWaitListRecords = 1)
	BEGIN	
	
		/* Moved from below to get @sWaitListCourseTitleColumnName in time*/
		IF @piWaitListCourseTitleColumnID > 0
		BEGIN
			/* Get the Waiting List Course Title column name. */
			SELECT @sWaitListCourseTitleColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @piWaitListCourseTitleColumnID;
		END
		IF @sWaitListCourseTitleColumnName IS NULL SET @psErrorMessage = 'Unable to find the Waiting List Course Title column.';

		/*-------------------------------------------------------------------------------------------------------------*/
		/*NPG20080422 Faults 13024 and 13025																		   */
		/*-------------------------------------------------------------------------------------------------------------*/
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

		/* Initialise the insert strings with the basic values.*/

		SET @sColumnList = @sWaitListCourseTitleColumnName + ',' +
			'id_' + convert(nvarchar(255), @piEmployeeTableID);
		SET @sValueList = '''' + @psCourseTitle  + ''',' +
			'id_' + convert(nvarchar(255), @piEmployeeTableID);
		SET @sAddedColumns = ',' + convert(varchar(255), @piWaitListCourseTitleColumnID) + ',';

		/* Get the TB and WL column permissions. */
		IF @fSysSecMgr = 0
		BEGIN
			DECLARE @columnPermissions TABLE(
				columnID	int,
				[action]		int,		
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
				AND (ASRSysColumns.tableID = @piTrainBookTableID
					OR ASRSysColumns.tableID = @piWaitListTableID))
			WHERE sysprotects.uid = @iUserGroupID
				AND (sysprotects.action = 193 OR sysprotects.action = 197)
				AND (sysobjects.name = @sTrainBookRealSource
					OR sysobjects.name = @sWaitListRealSource)
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
					WHERE columnID = @iDestinationColumnID
						AND [action] = 193;
					IF @fGranted IS null SET @fGranted = 0;
					IF @fGranted = 1
					BEGIN
						SELECT @fGranted = granted
						FROM @columnPermissions
						WHERE columnID = @iSourceColumnID
							AND [action] = 197;
						IF @fGranted IS null SET @fGranted = 0;
					END
				END
				IF @fGranted = 1
				BEGIN
					SELECT @sTempTBColumnName = columnName
					FROM [dbo].[ASRSysColumns]
					WHERE columnID = @iSourceColumnID;
					
					SELECT @sTempWLColumnName = columnName
					FROM [dbo].[ASRSysColumns]
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
		
		/* Create Waiting List records if required. */
		IF LEN(@psErrorMessage) = 0 
		BEGIN
			IF @pfWaitListCourseTitleColumnUpdate = 0 SET @psErrorMessage = 'You do not have ''edit'' permission on the Waiting List Course Title column.';
		END
		IF LEN(@psErrorMessage) = 0 
		BEGIN
			IF @pfWaitListTableInsert = 0 SET @psErrorMessage = 'You do not have ''new'' permission on the Waiting List table.';
		END
		IF LEN(@psErrorMessage) = 0 
		BEGIN
			SELECT @sWLTableName = tableName
			FROM [dbo].[ASRSysTables]
			WHERE tableID = @piWaitListTableID;
			
			SELECT @iChildViewID = childViewID
			FROM [dbo].[ASRSysChildViews2]
			WHERE tableID = @piWaitListTableID
				AND [role] = @sUserGroupName;
				
			IF @iChildViewID IS null SET @iChildViewID = 0;
				
			IF @iChildViewID > 0 
			BEGIN
				SET @sWaitListRealSource = 'ASRSysCV' + 
					convert(varchar(1000), @iChildViewID) +
					'#' + replace(@sWLTableName, ' ', '_') +
					'#' + replace(@sUserGroupName, ' ', '_');
				SET @sWaitListRealSource = left(@sWaitListRealSource, 255);
			END
			ELSE
			BEGIN
				SET @psErrorMessage = 'Unable to determine the Waiting List child view.';
			END
		END

		IF LEN(@psErrorMessage) = 0 		
		BEGIN	

			/* Create Waiting List records if required. */
			/* NB. Insert permission on the table is checked above.
				The 'sp_ASRIntGetTrainingBookingParameters' stored procedure is run before this one, 
				as the user logs into the intranet module, and checks certain permissions. 
				If the user does not have update permission on the Course Title column 
				then this point won't be reached. */

			SET @sCommand = 'INSERT INTO ' + @sWaitListRealSource + 
				'(' + @sColumnList + ')' +
				' (SELECT ' + @sValueList + 
				' FROM ' + @sTrainBookRealSource + 
				' WHERE id_' + convert(nvarchar(255), @piCourseTableID) + ' = ' + convert(nvarchar(255), @piCourseRecordID) +  
				' AND id_' + convert(nvarchar(255), @piEmployeeTableID) + ' > 0' +
				' AND ''' + @psCourseTitle + ''' NOT IN (SELECT ' + @sWaitListRealSource + '.'+ @sWaitListCourseTitleColumnName +
				' FROM ' + @sWaitListRealSource + 
				' WHERE ' + @sWaitListRealSource + '.id_' + convert(nvarchar(255), @piEmployeeTableID) + ' = ' + @sTrainBookRealSource + '.id_' + convert(nvarchar(255), @piEmployeeTableID) + ')';

			IF @fTransferProvisionals  = 1
			BEGIN
				SET @sCommand = @sCommand +
					' AND (LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''' +
					' OR LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''P''))';
			END
			ELSE
			BEGIN
				SET @sCommand = @sCommand +
					' AND LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B'')';
			END
			EXEC sp_executesql @sCommand;
		END
	END

	IF LEN(@psErrorMessage) = 0 		
	BEGIN
		/* Check if the 'CC' status code exists. */
		SET @fCCStatusExists = 0;
		
		SELECT @iCount = count(value)
		FROM [dbo].[ASRSysColumnControlValues]
		WHERE columnID = @piTrainBookStatusColumnID
			AND value = 'CC';
			
		IF @iCount > 0 SET @fCCStatusExists = 1;
		/* Check if the 'C' status code exists. */
		SET @fCStatusExists = 0;
		
		SELECT @iCount = count(value)
		FROM [dbo].[ASRSysColumnControlValues]
		WHERE columnID = @piTrainBookStatusColumnID
			AND value = 'C';
			
		IF @iCount > 0 SET @fCStatusExists = 1;
		IF @fCStatusExists = 0 SET @psErrorMessage = 'The Training Booking Status column does not have ''C'' as a valid value.';
	END

	IF LEN(@psErrorMessage) = 0 		
	BEGIN
		/* Update the existing training booking records. */	
		/* NB. Update permission on the table is checked above.
			The 'sp_ASRIntGetTrainingBookingParameters' stored procedure is run before this one, 
				as the user logs into the intranet module, and checks certain permissions. 
				If the user does not have update permission on the Status column or Cancel Date column (if there is one)
				then this point won't be reached. */
		SET @sCommand = 'UPDATE ' + @sTrainBookRealSource +
			' SET ' + @sTrainBookStatusColumnName + ' = ' +
			CASE 
				WHEN @fCCStatusExists = 1 THEN '''CC'''
				ELSE '''C'''
			END;

		SET @sCommand = @sCommand +
			' WHERE id_' + convert(nvarchar(255), @piCourseTableID) + ' = ' + convert(nvarchar(255), @piCourseRecordID);
       
		IF @fTransferProvisionals  = 1
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
		EXEC sp_executesql @sCommand;
	END

	IF LEN(@psErrorMessage) > 0 		
	BEGIN
		RAISERROR(@psErrorMessage, 16, 1);
		ROLLBACK;
	END
	ELSE COMMIT TRANSACTION;
END