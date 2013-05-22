CREATE PROCEDURE [dbo].[sp_ASRIntValidateBulkBookings] (
	@piCourseRecordID				integer,
	@psEmployeeRecordIDs			varchar(MAX),
	@psBookingStatus				varchar(MAX),
	@piResultCode					integer			OUTPUT,
	@psErrorMessage					varchar(MAX)	OUTPUT,
	@psWhoFailedPreReqCheck			varchar(MAX)	OUTPUT,
	@psWhoFailedUnavailabilityCheck	varchar(MAX)	OUTPUT,
	@psWhoFailedOverlapCheck		varchar(MAX)	OUTPUT,
	@psWhoFailedOverbookingCheck	varchar(MAX)	OUTPUT
)
AS
BEGIN
	/* This stored procedure run the pre-requisite, overbooking, overlapped booking and unavailability checks
	on booking being made. 
	Return codes are :
		0 - completely valid
		If non-zero then the result code is composed as abcd,
		where a is the result of the OVERBOOKING check
			b is the result of the PRE-REQUISITES check
			c is the result of the AVAILABILITY check
			d is the result of the OVERLAPPED BOOKING check.
		the values of which can be :
			0 if the check PASSED
			1 if the check FAILED and CANNOT be overridden
			2 if the check FAILED but CAN be overridden

		eg. if the current record passed the overbooking or overlapped bookings checks, but failed the availability check (overridable),
		and the pre-requisite check (not overridable) then the result code would be 0120.
	*/
	DECLARE	
			@fPreReqsOverridden		bit,
			@fUnavailOverridden		bit,
			@fOverlapOverridden		bit,
			@iCount					integer,
			@iNumberBooked			integer,
			@iResult				integer,
			@iIndex					integer,
			@fFailure				bit,
			@sCurrentCode			varchar(1),
			@sCodeString			varchar(10),
			@iEmployeeID			integer,
			@fDoPreReqCheck			bit,
			@iPreReqTableID			integer,
			@fDoUnavailabilityCheck	bit,
			@iUnavailTableID		integer,
			@fDoOverlapCheck		bit,
			@fDoOverbookingCheck	bit,
			@sTemp					varchar(MAX),
			@piTableID				integer,
			@psRecordDescription	varchar(MAX),
			@psEmployeeName			varchar(MAX);

	SET @piResultCode = 0
	SET @fPreReqsOverridden = 0
	SET @fUnavailOverridden = 0
	SET @fOverlapOverridden = 0
	SET @iNumberBooked = 0
	SET @fDoPreReqCheck = 0
	SET @fDoUnavailabilityCheck = 0
	SET @fDoOverlapCheck = 0
	SET @fDoOverbookingCheck = 0
	SET @psErrorMessage = ''
	SET @psWhoFailedPreReqCheck = ''
	SET @psWhoFailedUnavailabilityCheck = ''
	SET @psWhoFailedOverlapCheck = ''
	SET @psWhoFailedOverbookingCheck = ''
	
	/* Clean the input string parameters. */
	IF len(@psEmployeeRecordIDs) > 0 SET @psEmployeeRecordIDs = replace(@psEmployeeRecordIDs, '''', '''''')

	/* Check if we need to do the overbooking check. */
	SELECT @sTemp = parameterValue
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseIncludeProvisionals'
	IF @psBookingStatus = 'B' OR @sTemp = 'TRUE'
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM sysobjects
		WHERE id = object_id('sp_ASR_TBCheckOverbooking')
			AND sysstat & 0xf = 4

		IF @iCount > 0 SET @fDoOverbookingCheck = 1
	END
	
	/* Check if we need to do the pre-requisite check. */
	SELECT @iPreReqTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_PreReqTable'
	IF @iPreReqTableID IS NULL SET @iPreReqTableID = 0

	IF @iPreReqTableID > 0 
	BEGIN
		/* Check if the pre-req stored procedure exists. */
		SELECT @iCount = COUNT(*) 
		FROM sysobjects
		WHERE id = object_id('sp_ASR_TBCheckPreRequisites')
			AND sysstat & 0xf = 4
		
		IF @iCount > 0 SET @fDoPreReqCheck = 1
	END

	/* Check if we need to do the unavailibility check. */
	SELECT @iUnavailTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_UnavailTable'
	IF @iUnavailTableID IS NULL SET @iUnavailTableID = 0

	IF @iUnavailTableID > 0 
	BEGIN
		/* Check if the unavailibility stored procedure exists. */
		SELECT @iCount = COUNT(*) 
		FROM sysobjects
		WHERE id = object_id('sp_ASR_TBCheckUnavailability')
			AND sysstat & 0xf = 4
		
		IF @iCount > 0 SET @fDoUnavailabilityCheck = 1
	END

	/* Check if we need to do the overlap check. */
	/* Check if the unavailibility stored procedure exists. */
	SELECT @iCount = COUNT(*) 
	FROM sysobjects
	WHERE id = object_id('sp_ASR_TBCheckOverlappedBooking')
		AND sysstat & 0xf = 4
	
	IF @iCount > 0 SET @fDoOverlapCheck = 1

	/* Loop thourgh the given employee records. */
	SET @fFailure = 0
	SET @piResultCode = 0
	SET @iNumberBooked = 0

	WHILE len(@psEmployeeRecordIDs) > 0
	BEGIN
		/* Rip out the individual employee record ID from the given comma-delimited string of employee IDs. */
		SELECT @iIndex = charindex(',', @psEmployeeRecordIDs)
		IF @iIndex > 0
		BEGIN
			SET  @iEmployeeID = substring(@psEmployeeRecordIDs, 1, @iIndex - 1)
			SELECT @psEmployeeRecordIDs = substring(@psEmployeeRecordIDs, @iIndex + 1, len(@psEmployeeRecordIDs))
		END
		ELSE
		BEGIN
			SET  @iEmployeeID = @psEmployeeRecordIDs
			SET @psEmployeeRecordIDs = ''
		END

		BEGIN
			SET @piTableID = 1 /* Need to derive the Personnel table */
			EXECUTE dbo.spASRRecordDescription 1, @iEmployeeID, @psRecordDescription OUTPUT
			SET  @psEmployeeName = @psRecordDescription
		END

		SET @sCodeString = convert(varchar(10), @piResultCode)

		IF @fDoPreReqCheck = 1
		BEGIN
			IF LEN(@sCodeString) >= 3
			BEGIN
				SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 3), 1)
			END
			ELSE SET @sCurrentCode = 0
			
			IF @sCurrentCode <> '1'
			BEGIN
				/* Return 0 if the given record in the personnel table has satisfied the pre-requisite criteria for the given course record.
				Return 1 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record.
				Return 2 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record but the user can override this failure. */
				exec sp_ASR_TBCheckPreRequisites @piCourseRecordID,  @iEmployeeID, @iResult OUTPUT
				IF @iResult = 1 
				BEGIN
					SET @psWhoFailedPreReqCheck = @psWhoFailedPreReqCheck 
						+ CASE
							WHEN len(@psWhoFailedPreReqCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END
				IF @iResult = 2 
				BEGIN
					SET @psWhoFailedPreReqCheck = @psWhoFailedPreReqCheck
						+ CASE
							WHEN len(@psWhoFailedPreReqCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END
				SET @piResultCode = @piResultCode - (100 * convert(integer, @sCurrentCode)) + (100 * @iResult)
			END
		END

		IF @fDoUnavailabilityCheck = 1
		BEGIN
			IF len(@sCodeString) >= 2
			BEGIN 
				SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 2), 1)
			END
			ELSE SET @sCurrentCode = 0
			
			IF @sCurrentCode <> '1'
			BEGIN
				exec sp_ASR_TBCheckUnavailability @piCourseRecordID,  @iEmployeeID, @iResult OUTPUT
				/* Return 0 if the given record in the personnel table IS available for the given course record.
				Return 1 if the given record in the personnel table is NOT available for the given course record.
				Return 2 if the given record in the personnel table is NOT available for the given course record but the user can override this failure. */
				IF @iResult = 1 
				BEGIN
					SET @psWhoFailedUnavailabilityCheck = @psWhoFailedUnavailabilityCheck
						+ CASE
							WHEN len(@psWhoFailedUnavailabilityCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END	
				IF @iResult = 2 
				BEGIN
					SET @psWhoFailedUnavailabilityCheck = @psWhoFailedUnavailabilityCheck
						+ CASE
							WHEN len(@psWhoFailedUnavailabilityCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END				
				ELSE 
				SET @piResultCode = @piResultCode - (10 * convert(integer, @sCurrentCode)) + (10 * @iResult)
			END
		END

		IF @fDoOverlapCheck = 1
		BEGIN
			IF len(@sCodeString) >= 1
			BEGIN
				SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 1), 1)
			END
			ELSE SET @sCurrentCode = 0

			IF @sCurrentCode <> '1'
			BEGIN
				exec sp_ASR_TBCheckOverlappedBooking @piCourseRecordID,  @iEmployeeID, 0, @iResult OUTPUT
				  /* Return 0 if the given course does NOT overlap with another course that the given delegate is booked on.
				  Return 1 if the given course DOES overlap with another course that the given delegate is booked on.
				  Return 2 if the given course does NOT overlap with another course that the given delegate is booked on, but the user can override this failure. */
				IF @iResult = 1 
				BEGIN
					SET @psWhoFailedOverlapCheck = @psWhoFailedOverlapCheck
						+ CASE
							WHEN len(@psWhoFailedOverlapCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END
				IF @iResult = 2 
				BEGIN
					SET @psWhoFailedOverlapCheck = @psWhoFailedOverlapCheck
						+ CASE
							WHEN len(@psWhoFailedOverlapCheck) > 0 THEN ', '
							ELSE ''
						END 
						+ '"' + @psEmployeeName + '"'
				END
				SET @piResultCode = @piResultCode - convert(integer, @sCurrentCode) + @iResult
			END
		END

		SET @iNumberBooked = @iNumberBooked + 1
	END

	IF LEN(@psErrorMessage) = 0 AND (@fDoOverbookingCheck = 1)
	BEGIN
		exec sp_ASR_TBCheckOverbooking @piCourseRecordID, 0, @iNumberBooked, @iResult OUTPUT

		IF @iResult = 1 /* Course fully booked (error). */
		BEGIN
			SET @piResultCode = @piResultCode + 1000
		END
		IF @iResult = 2 /* Course fully booked (over-rideable by the user). */
		BEGIN
			SET @piResultCode = @piResultCode + 2000
		END
	END
END