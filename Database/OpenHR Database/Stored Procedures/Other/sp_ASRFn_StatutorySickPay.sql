CREATE PROCEDURE [dbo].[sp_ASRFn_StatutorySickPay]
(
	@piAbsenceRecordID		int
)
AS
BEGIN
	/* Refresh the SSP fields in the Absence records for the Personnel record that is the parent of the given Absence record ID. */

	/* Absence module - Personnel table variables. */
	DECLARE @iPersonnelTableID				integer,
		@sPersonnelTableName 				varchar(128),
		@sWorkingDaysNum_ColumnName 		varchar(128),
		@sWorkingDaysPattern_ColumnName 	varchar(128),
		@sDateOfBirth_ColumnName 			varchar(128);

	/* Personnel record variables. */
	DECLARE @iPersonnelRecordID 			integer,
		@iWorkingDaysPerWeek 				integer,
		@sWorkingPattern 					varchar(MAX),
		@dtDateOfBirth						datetime,
		@dtRetirementDate					datetime,
		@dtSixteenthBirthday				datetime;

	/* Absence module - Absence table variables. */
	DECLARE @sAbsenceTableName				varchar(128),
		@sAbsence_StartDateColumnName		varchar(128),
		@sAbsence_EndDateColumnName			varchar(128),
		@sAbsence_StartSessionColumnName	varchar(128),
		@sAbsence_EndSessionColumnName		varchar(128),
		@sAbsence_TypeColumnName			varchar(128),
		@sAbsence_SSPAppliesColumnName 		varchar(128),
		@sAbsence_QualifyingDaysColumnName 	varchar(128),
		@sAbsence_WaitingDaysColumnName 	varchar(128),
		@sAbsence_PaidDaysColumnName 		varchar(128),
		@iAbsence_WorkingDaysType 			integer;

	/* Absence record variables. */
	DECLARE @cursAbsenceRecords			cursor,
		@cursFollowingAbsenceRecords	cursor,
		@iAbsenceRecordID 				integer,
		@dtStartDate 					datetime,
		@dtEndDate 						datetime,
		@sStartSession					varchar(MAX),
		@sEndSession 					varchar(MAX),
		@dtWholeStartDate 				datetime,
		@dtWholeEndDate 				datetime,
		@dtFollowingStartDate 			datetime,
		@dtFollowingEndDate 			datetime,
		@sFollowingStartSession	 		varchar(MAX),
		@sFollowingEndSession 			varchar(MAX),
		@dtFollowingWholeStartDate 		datetime,
		@dtFollowingWholeEndDate 		datetime,
		@fOriginalSSPApplies			bit,
		@dblOriginalQualifyingDays		float,
		@dblOriginalWaitingDays			float,
		@dblOriginalPaidDays			float;

	/* Absence module - Absence Type table variables. */
	DECLARE @sAbsenceTypeTableName			varchar(128),
		@sAbsenceType_TypeColumnName		varchar(128),
		@sAbsenceType_SSPAppliesColumnName	varchar(128);

	/* General procedure handling variables. */
	DECLARE @fOK	 					bit,
		@iLoop							integer,
		@iIndex							integer,
		@sCommandString					nvarchar(MAX),
		@sParamDefinition				nvarchar(500),
		@dblWaitEntitlement 			float,
		@dblAbsenceEntitlement 			float,
		@dblQualifyingDays 				float,
		@dblWaitingDays 				float,
		@dblPaidDays 					float,
		@fSSPApplies					bit,
		@dtTempDate						datetime,
		@fAddOK							bit,
		@dblAddAmount					float,
		@fContinue 						bit,
		@iConsecutiveRecords			integer,
		@dtConsecutiveStartDate 		datetime,
		@dtConsecutiveEndDate 			datetime,
		@dtConsecutiveWholeStartDate 	datetime,
		@dtConsecutiveWholeEndDate 		datetime,
		@sConsecutiveStartSession 		varchar(MAX),
		@sConsecutiveEndSession 		varchar(MAX),
		@dtLastWholeEndDate 			datetime,
		@dtFirstLinkedWholeStartDate 	datetime,
		@iYearDifference 				integer;

	SET @fOK = 1;

	/* Get the Absence module parameters. */
	/* Get the Personnel table name and ID. */
	SELECT @iPersonnelTableID = convert(integer, parameterValue), 
		@sPersonnelTableName = ASRSysTables.tableName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysTables 
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
	WHERE moduleKey = 'MODULE_PERSONNEL'
	AND parameterKey = 'Param_TablePersonnel';

	/* Get the Personnel - Date Of Birth column name. */
	SELECT @sDateOfBirth_ColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_PERSONNEL'
	AND parameterKey = 'Param_FieldsDateOfBirth';

	/* Get the Absence table name. */
	SELECT @sAbsenceTableName = ASRSysTables.tableName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysTables 
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_TableAbsence';

	/* Get the Absence - Start Date column name. */
	SELECT @sAbsence_StartDateColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldStartDate';

	/* Get the Absence - End Date column name. */
	SELECT @sAbsence_EndDateColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldEndDate';

	/* Get the Absence - Start Session column name. */
	SELECT @sAbsence_StartSessionColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldStartSession';

	/* Get the Absence - End Session column name. */
	SELECT @sAbsence_EndSessionColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldEndSession';

	/* Get the Absence - Type column name. */
	SELECT @sAbsence_TypeColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldType';

	/* Get the Absence - SSP Applies column name. */
	SELECT @sAbsence_SSPAppliesColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldSSPApplies';

	/* Get the Absence - Qualifying Days column name. */
	SELECT @sAbsence_QualifyingDaysColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldQualifyingDays';

	/* Get the Absence - Waiting Days column name. */
	SELECT @sAbsence_WaitingDaysColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldWaitingDays';

	/* Get the Absence - Paid Days column name. */
	SELECT @sAbsence_PaidDaysColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldPaidDays';

	/* Get the Absence - Working Days selection type. */
	SELECT @iAbsence_WorkingDaysType = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_WorkingDaysType';

	/* Get the Absence Type table name. */
	SELECT @sAbsenceTypeTableName = ASRSysTables.tableName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysTables
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_TableAbsenceType';

	/* Get the Absence Type - Type column name. */
	SELECT @sAbsenceType_TypeColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldTypeType';

	/* Get the Absence Type - SSP Applies column name. */
	SELECT @sAbsenceType_SSPAppliesColumnName = ASRSysColumns.columnName
	FROM ASRSysModuleSetup
	INNER JOIN ASRSysColumns
		ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
	WHERE moduleKey = 'MODULE_ABSENCE'
	AND parameterKey = 'Param_FieldTypeSSP';

	/* Validate the Absence module variables. */
	IF (@iPersonnelTableID IS null)
		OR (@sPersonnelTableName IS null)
		OR (@sAbsenceTableName IS null) 
		OR (@sAbsence_StartDateColumnName IS null) 
		OR (@sAbsence_EndDateColumnName IS null) 
		OR (@sAbsence_StartSessionColumnName IS null) 
		OR (@sAbsence_EndSessionColumnName IS null)  
		OR (@sAbsence_TypeColumnName IS null)   
		OR (@sAbsence_SSPAppliesColumnName IS null)
		OR (@sAbsence_QualifyingDaysColumnName IS null)
		OR (@sAbsence_WaitingDaysColumnName IS null)
		OR (@sAbsence_PaidDaysColumnName IS null)
		OR (@iAbsence_WorkingDaysType IS null)
		OR (@sAbsenceTypeTableName IS null)   
		OR (@sAbsenceType_TypeColumnName IS null)    
		OR (@sAbsenceType_SSPAppliesColumnName IS null) SET @fOK = 0;

	IF @fOK = 1
	BEGIN
		/* Get the ID  of the associated record in the Personnel table. */
		SET @sParamDefinition = N'@recordID integer OUTPUT';
		SET @sCommandString = 'SELECT @recordID = id_' + convert(varchar(128), @iPersonnelTableID) + 
			' FROM ' + @sAbsenceTableName + 
			' WHERE id = ' + convert(varchar(128), @piAbsenceRecordID);
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iPersonnelRecordID OUTPUT;

		IF (@iPersonnelRecordID IS null) OR (@iPersonnelRecordID <= 0) SET @fOK = 0;
	END

	IF (@fOK = 1) AND (NOT @sDateOfBirth_ColumnName IS null) 
	BEGIN
		/* Get the retirement date, and the date of the person's sixteenth birthday. */
		SET @sParamDefinition = N'@dateOfBirth datetime OUTPUT';
		SET @sCommandString = 'SELECT @dateOfBirth = convert(datetime, convert(varchar(20), ' + @sDateOfBirth_ColumnName + ', 101))' +
			' FROM ' + @sPersonnelTableName + 
			' WHERE id = ' + convert(varchar(128), @iPersonnelRecordID);
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dtDateOfBirth OUTPUT;

		IF (NOT @dtDateOfBirth IS null) SET @dtRetirementDate = dateadd(yy, 65, @dtDateOfBirth);
		IF (NOT @dtDateOfBirth IS null) SET @dtSixteenthBirthday = dateadd(yy, 16, @dtDateOfBirth);
	END

	IF @fOK = 1 
	BEGIN
		/* Get the number of working days per week. */
		SET @iWorkingDaysPerWeek = 0;
		SET @sWorkingPattern = '';

		IF @iAbsence_WorkingDaysType = 0	/* The Working Days are an straight numeric value. */
		BEGIN
			SELECT @iWorkingDaysPerWeek = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_ABSENCE'
			AND parameterKey = 'Param_WorkingDaysNum';

			IF @iWorkingDaysPerWeek IS null SET @fOK = 0;
		END

		IF @iAbsence_WorkingDaysType = 1	/* The Working Days are an straight working pattern value. */
		BEGIN
			SELECT @sWorkingPattern = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_ABSENCE'
			AND parameterKey = 'Param_WorkingDaysPattern';
			
			IF @sWorkingPattern IS null SET @fOK = 0;
		END

		IF @iAbsence_WorkingDaysType = 2	/* The Working Days are a numeric field reference. */
		BEGIN
			SELECT @sWorkingDaysNum_ColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = 'MODULE_ABSENCE'
			AND parameterKey = 'Param_FieldWorkingDays';

			IF @sWorkingDaysNum_ColumnName IS null SET @fOK = 0;

			IF @fOK = 1
			BEGIN
				SET @sParamDefinition = N'@workingDays varchar(MAX) OUTPUT'
				SET @sCommandString = 'SELECT @workingDays = ' + @sWorkingDaysNum_ColumnName + 
					' FROM ' + @sPersonnelTableName + 
					' WHERE id = ' + convert(varchar(128), @iPersonnelRecordID);
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iWorkingDaysPerWeek OUTPUT;

				IF (@iWorkingDaysPerWeek IS null) SET @fOK = 0;
			END
		END

		IF @iAbsence_WorkingDaysType = 3	/* The Working Days are an working pattern field. */
		BEGIN
			SELECT @sWorkingDaysPattern_ColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = 'MODULE_ABSENCE'
			AND parameterKey = 'Param_FieldWorkingDays';

			IF @sWorkingDaysPattern_ColumnName IS null SET @fOK = 0;

			IF @fOK = 1
			BEGIN
				SET @sParamDefinition = N'@workingDaysPattern varchar(MAX) OUTPUT'
				SET @sCommandString = 'SELECT @workingDaysPattern = ' + @sWorkingDaysNum_ColumnName + 
					' FROM ' + @sPersonnelTableName + 
					' WHERE id = ' + convert(varchar(128), @iPersonnelRecordID);
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @sWorkingPattern OUTPUT;

				IF (@sWorkingPattern IS null) SET @fOK = 0;
			END
		END

		IF @fOK = 1
		BEGIN
			/* Calculate the number of qualifying days per week. */
			IF len(@sWorkingPattern) > 0
			BEGIN
				SET @iLoop = 1;

				WHILE (len(@sWorkingPattern) >= (@iLoop * 2)) AND (@iLoop <=14)
				BEGIN
					IF (substring(@sWorkingPattern, @iLoop, 1) <> ' ') AND (substring(@sWorkingPattern, @iLoop + 1, 1) <> ' ')
					BEGIN
						SET @iWorkingDaysPerWeek = @iWorkingDaysPerWeek + 1;
					END
				
					SET @iLoop = @iLoop + 2;
				END
			END

			IF @iWorkingDaysPerWeek <= 0 SET @fOK = 0;
		END
	END

	IF @fOK = 1
	BEGIN
		SET @iConsecutiveRecords = 0;
		SET @dtLastWholeEndDate = null;

		/* Create a cursor of the absence records for the current person. */
		SET @sParamDefinition = N'@absenceRecs cursor OUTPUT'
		SET @sCommandString = 'SET @absenceRecs = CURSOR  LOCAL FAST_FORWARD FOR' +
			' SELECT ' + @sAbsenceTableName + '.id, ' + 
				'convert(datetime, convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', 101)), ' + 
				'convert(datetime, convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_EndDateColumnName + ', 101)), ' +
				'upper(left(' + @sAbsenceTableName + '.' + @sAbsence_StartSessionColumnName + ', 2)), ' +
				'upper(left(' + @sAbsenceTableName + '.' + @sAbsence_EndSessionColumnName + ', 2)), ' + 
				@sAbsenceTableName + '.' + @sAbsence_SSPAppliesColumnName + ', ' +
				@sAbsenceTableName + '.' + @sAbsence_QualifyingDaysColumnName + ', ' +
				@sAbsenceTableName + '.' + @sAbsence_WaitingDaysColumnName + ', ' +
				@sAbsenceTableName + '.' + @sAbsence_PaidDaysColumnName + 
			' FROM ' + @sAbsenceTableName + 
			' INNER JOIN ' + @sAbsenceTypeTableName + ' ON ' + @sAbsenceTableName + '.' + @sAbsence_TypeColumnName + ' = ' + @sAbsenceTypeTableName + '.' + @sAbsenceType_TypeColumnName +
			' WHERE ' + @sAbsenceTableName + '.id_' + convert(varchar(128), @iPersonnelTableID) + ' = ' + convert(varchar(128), @iPersonnelRecordID) +
			' AND ' + @sAbsenceTypeTableName + '.' + @sAbsenceType_SSPAppliesColumnName + ' = 1' +
			' ORDER BY ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', ' + @sAbsenceTableName + '.id' +
			' OPEN @absenceRecs';
		EXECUTE sp_executesql @sCommandString, @sParamDefinition, @cursAbsenceRecords OUTPUT;

		/* Loop through the absence records, calculating SSP for each record. 
		NB. We check if any periods of absence are consecutive before checking for SSP application. */
		FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays;
		WHILE (@@fetch_status = 0)
		BEGIN
			/* Ignore incomplete absence records. */
			IF (NOT @dtStartDate IS null) AND (NOT @dtEndDate IS null)
			BEGIN
				/* Ignore absence after retirement. */
				IF NOT @dtRetirementDate IS null
				BEGIN
					IF (@dtRetirementDate < @dtEndDate) 
					BEGIN
						SET @dtEndDate = @dtRetirementDate;
						SET @sEndSession = 'PM';
					END
				END
				/* Ignore absence before the sixteenth birthday. */
				IF NOT @dtSixteenthBirthday IS null
				BEGIN
					IF (@dtSixteenthBirthday > @dtStartDate) 
					BEGIN
						SET @dtStartDate = @dtSixteenthBirthday;
						SET @sStartSession = 'AM';
					END
				END

				/* Get the start and end dates (whole days only) of the current absence record. */
				SET @dtWholeStartDate = @dtStartDate;
				SET @dtWholeEndDate = @dtEndDate;
				IF @sStartSession = 'PM' SET @dtWholeStartDate = @dtWholeStartDate + 1;
				IF @sEndSession = 'AM' SET @dtWholeEndDate = @dtWholeEndDate - 1;

				IF @iConsecutiveRecords = 0 
				BEGIN
					SET @dtConsecutiveStartDate = @dtStartDate;
					SET @dtConsecutiveEndDate = @dtEndDate;
					SET @sConsecutiveStartSession = @sStartSession;
					SET @sConsecutiveEndSession = @sEndSession;
					SET @dtConsecutiveWholeStartDate = @dtWholeStartDate;
					SET @dtConsecutiveWholeEndDate = @dtWholeEndDate;

					/* Create a cursor of the absence records for the current person that follow the current absence record. */
					SET @sParamDefinition = N'@followingAbsenceRecs cursor OUTPUT';
					SET @sCommandString = 'SET @followingAbsenceRecs = CURSOR  LOCAL FAST_FORWARD FOR' +
						' SELECT convert(datetime, convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', 101)), ' + 
							'convert(datetime, convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_EndDateColumnName + ', 101)), ' +
							'upper(left(' + @sAbsenceTableName + '.' + @sAbsence_StartSessionColumnName + ', 2)), ' +
							'upper(left(' + @sAbsenceTableName + '.' + @sAbsence_EndSessionColumnName + ', 2)) ' + 
						' FROM ' + @sAbsenceTableName + 
						' INNER JOIN ' + @sAbsenceTypeTableName + ' ON ' + @sAbsenceTableName + '.' + @sAbsence_TypeColumnName + ' = ' + @sAbsenceTypeTableName + '.' + @sAbsenceType_TypeColumnName +
						' WHERE ' + @sAbsenceTableName + '.id_' + convert(varchar(128), @iPersonnelTableID) + ' = ' + convert(varchar(128), @iPersonnelRecordID) +
						' AND ' + @sAbsenceTypeTableName + '.' + @sAbsenceType_SSPAppliesColumnName + ' = 1' +
						' AND (NOT ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ' IS null)' + 
						' AND (NOT ' + @sAbsenceTableName + '.' + @sAbsence_EndDateColumnName + ' IS null)' +
						' AND ((convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', 112) > ' + convert(varchar(20), @dtStartDate, 112) + ')' +
						' OR ((convert(varchar(20), ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', 112) = ' + convert(varchar(20), @dtStartDate, 112) + ') AND (' + @sAbsenceTableName + '.id > ' + convert(varchar(128), @iAbsenceRecordID) + ')))' +
						' ORDER BY ' + @sAbsenceTableName + '.' + @sAbsence_StartDateColumnName + ', ' + @sAbsenceTableName + '.id' +
						' OPEN @followingAbsenceRecs';
					EXECUTE sp_executesql @sCommandString, @sParamDefinition, @cursFollowingAbsenceRecords OUTPUT;

					SET @fContinue = 1;
					FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession;
					WHILE (@@fetch_status = 0) AND (@fContinue = 1)
					BEGIN
						SET @fContinue = 0;
			
						/* Get the start and end dates (whole days only) of the current absence records. */
						SET @dtFollowingWholeStartDate = @dtFollowingStartDate;
						SET @dtFollowingWholeEndDate = @dtFollowingEndDate;
						IF @sFollowingStartSession = 'PM' SET @dtFollowingWholeStartDate = @dtFollowingWholeStartDate + 1;
						IF @sFollowingEndSession = 'AM' SET @dtFollowingWholeEndDate = @dtFollowingWholeEndDate - 1;

						IF ((@dtConsecutiveEndDate = @dtFollowingStartDate) AND (@sConsecutiveEndSession = 'AM') AND (@sFollowingStartSession = 'PM'))
							OR (@dtConsecutiveWholeEndDate + 1 >= @dtFollowingWholeStartDate)
						BEGIN
							SET @iConsecutiveRecords = @iConsecutiveRecords + 1;
							SET @dtConsecutiveEndDate = @dtFollowingEndDate;
							SET @sConsecutiveEndSession = @sFollowingEndSession;
							SET @dtConsecutiveWholeEndDate = @dtFollowingWholeEndDate;
							SET @fContinue = 1;
						END

						FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession;
					END

					CLOSE @cursFollowingAbsenceRecords;
					DEALLOCATE @cursFollowingAbsenceRecords;

				END
				ELSE
				BEGIN
					SET @iConsecutiveRecords = @iConsecutiveRecords - 1;
				END

				/* SSP Applies if the absence period is greater than 3 days. */
				SET @fSSPApplies = 0;
				IF (datediff(dd, @dtConsecutiveWholeStartDate, @dtConsecutiveWholeEndDate) + 1) > 3 SET @fSSPApplies = 1;

				IF @fSSPApplies = 1
				BEGIN
					/* Check if 56 days have passed since the previous absence period. */
					IF @dtLastWholeEndDate IS null
					BEGIN
						/* First absence record so use default values. */
						SET @dblWaitEntitlement = 3;
						SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28;
						SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate;
					END
					ELSE
					BEGIN
						IF (datediff(dd, @dtLastWholeEndDate, @dtWholeStartDate) - 1) > 56
						BEGIN
							/* More than 56 days since the previous absence record so use default values. */
							SET @dblWaitEntitlement = 3;
							SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28;
							SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate;
						END
					END
		
					/* Calculate SSP qualifying, waiting and paid days.
					NB. The start and end dates should already take into account the start and end periods (AM/PM)
					so that only whole absence days are used. */
					SET @dblQualifyingDays = 0;

					/* Loop from the start date to the end date, incrementing the number of qualifying days for each date that qualifies. */
					SET @dtTempDate = @dtStartDate;

					WHILE (@dtTempDate <= @dtEndDate)
					BEGIN
						SET @fAddOK = 0;
						SET @dblAddAmount = 0;

						IF len(@sWorkingPattern) = 0
						BEGIN
							/* No working pattern passed in, so use the 'daysPerWeek' variable. */
							IF (@iWorkingDaysPerWeek = 7) OR 
								((datepart(dw, @dtTempDate) >= 2) AND (datepart(dw, @dtTempDate) <= 6))
							BEGIN
								/* The current date qualifies if 7 days per week are worked, or if the current date is a weekday. */
								SET @fAddOK = 1;
							END
						END
						ELSE	
						BEGIN
							/* Use the working pattern. */
							SET @iIndex = (2 * datepart(dw, @dtTempDate)) -1;
							IF len(@sWorkingPattern) >= (@iIndex +1)
							BEGIN
								/* The current date qualifies if its 'day of the week' is worked in the working pattern.
								NB. Both AM and PM sessions must be worked for the day to qualify. */
								IF (substring(@sWorkingPattern, @iIndex, 1) <> ' ') AND (substring(@sWorkingPattern, @iIndex + 1, 1) <> ' ')
								BEGIN
									SET @fAddOK = 1;
								END
							END
						END

						IF @fAddOK = 1 
						BEGIN
							/* If the person is older than retirement age, then the day does not qualify. */
							IF NOT @dtRetirementDate IS null
							BEGIN
								IF @dtTempDate > @dtRetirementDate SET @fAddOK = 0;
							END
						END

						IF @fAddOK = 1 
						BEGIN
							/* If the person is less than sixteen then the day does not qualify. */
							IF (NOT @dtSixteenthBirthday IS null) 
							BEGIN
								IF @dtTempDate < @dtSixteenthBirthday SET @fAddOK = 0;
							END
						END

						IF @fAddOK = 1 
						BEGIN
							/* Days linked after 3 years from the start of the link do not count. */
							exec sp_ASRFn_WholeYearsBetweenTwoDates @iYearDifference OUTPUT, @dtFirstLinkedWholeStartDate, @dtTempDate;
							IF @iYearDifference >= 3  SET @fAddOK = 0;
						END

						/* Calculate how much to add to the Qualifying Days. */
						IF @fAddOK = 1 
						BEGIN
							SET @dblAddAmount = 0;

							IF @dtTempDate < @dtWholeStartDate
							BEGIN
								/* The current date is the half day before the whole dated period starts.
								A half day qualifies only if this period of absence consecutively follows another. */
								IF (@dtConsecutiveStartDate < @dtStartDate) OR 
									((@dtConsecutiveStartDate = @dtStartDate) AND (@sConsecutiveStartSession <> @sStartSession)) SET @dblAddAmount = 0.5;
							END
							ELSE
							BEGIN
								IF @dtTempDate > @dtWholeEndDate
								BEGIN
									/* The current date is the half day after the whole dated period end.
									A half day qualifies only if this period of absence is consecutively followed by another. */
									IF (@dtConsecutiveEndDate > @dtEndDate) OR 
										((@dtConsecutiveEndDate = @dtEndDate) AND (@sConsecutiveEndSession <> @sStartSession)) SET @dblAddAmount = 0.5;
								END
								ELSE
								BEGIN
									/* The current date lies within the whole dated period, so a whole day qualifies. */
									SET @dblAddAmount = 1;
								END
							END
						END


						/* Increment the number of qualifying days. */
						SET @dblQualifyingDays = @dblQualifyingDays + @dblAddAmount;

						SET @dtTempDate = @dtTempDate + 1;
					END

					/* Take off any waiting entitlement. */
					IF @dblWaitEntitlement > @dblQualifyingDays
					BEGIN
						SET @dblWaitingDays = @dblQualifyingDays;
						SET @dblWaitEntitlement = @dblWaitEntitlement - @dblQualifyingDays;
					END
					ELSE
					BEGIN
						SET @dblWaitingDays = @dblWaitEntitlement;
						SET @dblWaitEntitlement = 0;
					END

					/* Paid days is the difference providing there is enough entitlement. */
					SET @dblPaidDays = @dblQualifyingDays - @dblWaitingDays;

					IF @dblPaidDays > @dblAbsenceEntitlement
					BEGIN
						SET @dblPaidDays = @dblAbsenceEntitlement;
						SET @dblAbsenceEntitlement = 0;
					END
					ELSE
					BEGIN
						SET @dblAbsenceEntitlement = @dblAbsenceEntitlement - @dblPaidDays;
					END	

					SET @dtLastWholeEndDate = @dtWholeEndDate;

					/* Update the SSP fields in the current absence record if required. */
					IF (@fOriginalSSPApplies IS null) OR
						(@fOriginalSSPApplies = 0) OR
						(@dblOriginalQualifyingDays IS null) OR
						(@dblOriginalQualifyingDays <> @dblQualifyingDays) OR
						(@dblOriginalWaitingDays IS null) OR
						(@dblOriginalWaitingDays <> @dblWaitingDays) OR
						(@dblOriginalPaidDays IS null) OR
						(@dblOriginalPaidDays <> @dblPaidDays)
					BEGIN
						SET @sCommandString = 'UPDATE ' + @sAbsenceTableName +
							' SET ' + @sAbsence_SSPAppliesColumnName + ' = 1, ' +
							@sAbsence_QualifyingDaysColumnName + ' = ' + convert(varchar(MAX), @dblQualifyingDays) + ', ' +
							@sAbsence_WaitingDaysColumnName + ' = ' + convert(varchar(MAX), @dblWaitingDays) + ', ' +
							@sAbsence_PaidDaysColumnName + ' = ' + convert(varchar(MAX), @dblPaidDays) + 
							' WHERE id = ' + convert(varchar(128), @iAbsenceRecordID);
						exec sp_executesql @sCommandString;
					END
				END
				ELSE			
				BEGIN
					/* Update the SSP fields in the current absence record. */
					IF (@fOriginalSSPApplies IS null) OR
						(@fOriginalSSPApplies = 1) OR
						(@dblOriginalQualifyingDays IS null) OR
						(@dblOriginalQualifyingDays <> 0) OR
						(@dblOriginalWaitingDays IS null) OR
						(@dblOriginalWaitingDays <> 0) OR
						(@dblOriginalPaidDays IS null) OR
						(@dblOriginalPaidDays <> 0)
					BEGIN
						SET @sCommandString = 'UPDATE ' + @sAbsenceTableName +
							' SET ' + @sAbsence_SSPAppliesColumnName + ' = 0, ' +
							@sAbsence_QualifyingDaysColumnName + ' = 0, ' +
							@sAbsence_WaitingDaysColumnName + ' = 0, ' +
							@sAbsence_PaidDaysColumnName + ' = 0' + 
							' WHERE id = ' + convert(varchar(128), @iAbsenceRecordID);
						exec sp_executesql @sCommandString;
					END
				END
			END
			ELSE
			BEGIN
				/* Update the SSP fields in the current absence record. */
				IF (@fOriginalSSPApplies IS null) OR
					(@fOriginalSSPApplies = 1) OR
					(@dblOriginalQualifyingDays IS null) OR
					(@dblOriginalQualifyingDays <> 0) OR
					(@dblOriginalWaitingDays IS null) OR
					(@dblOriginalWaitingDays <> 0) OR
					(@dblOriginalPaidDays IS null) OR
					(@dblOriginalPaidDays <> 0)
				BEGIN
					SET @sCommandString = 'UPDATE ' + @sAbsenceTableName +
						' SET ' + @sAbsence_SSPAppliesColumnName + ' = 0, ' +
						@sAbsence_QualifyingDaysColumnName + ' = 0, ' +
						@sAbsence_WaitingDaysColumnName + ' = 0, ' +
						@sAbsence_PaidDaysColumnName + ' = 0' + 
						' WHERE id = ' + convert(varchar(128), @iAbsenceRecordID);
					exec sp_executesql @sCommandString;
				END
			END

			FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays;
		END
		CLOSE @cursAbsenceRecords;
		DEALLOCATE @cursAbsenceRecords;
	END
END