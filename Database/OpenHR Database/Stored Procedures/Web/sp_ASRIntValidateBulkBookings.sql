CREATE PROCEDURE [dbo].[sp_ASRIntValidateBulkBookings] (
	@piCourseRecordID				integer,
	@psEmployeeRecordIDs			varchar(MAX),
	@psBookingStatus				varchar(MAX),
	@psErrorMessage				varchar(MAX)	OUTPUT,
	@psPreReqCheckFailsCount			integer	  	OUTPUT,
	@psUnavailabilityCheckFailCount	integer		OUTPUT,
	@psOverlapCheckFailCount			integer		OUTPUT,
	@psCourseOverbooked				integer		OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* This stored procedure run the pre-requisite, overbooking, overlapped booking and unavailability checks on booking being made. 
	Return codes are:
		000 - completely valid
		If non-zero then the result code is composed as abc,
		where a is the result of the OVERBOOKING check
			a is the result of the PRE-REQUISITES check
			b is the result of the AVAILABILITY check
			c is the result of the OVERLAPPED BOOKING check.
		the values of which can be :
			0 if the check PASSED
			1 if the check FAILED and CANNOT be overridden
			2 if the check FAILED but CAN be overridden

		eg. if the current record passed the overlapped bookings check, but failed the availability check (overridable),
		and the pre-requisite check (not overridable) then the result code would be 012.

		The SP returns a table indicating the result code for each Employee record (see below) plus a parameter that indicates if the course is overbooked

		EmployeeID  ResultCode
		----------- ----------
		1094        020		--> ResultCode is codified as described above
		961         000	     --> ResultCode is codified as described above

		The output @psCourseOverbooked parameter will be set to one of the following values:
		  0 - No overbooking
		  1 - Course fully booked
		  2 - Course fully booked (over-rideable by the user).
	*/

	DECLARE	
			@fPreReqsOverridden		bit,
			@fUnavailOverridden		bit,
			@fOverlapOverridden		bit,
			@iCount				integer,
			@iNumberBooked			integer,
			@iResult				integer,
			@iIndex				integer,
			@fFailure				bit,
			@iEmployeeID			integer,
			@fDoPreReqCheck		bit,
			@iPreReqTableID		integer,
			@fDoUnavailabilityCheck	bit,
			@iUnavailTableID		integer,
			@fDoOverlapCheck		bit,
			@fDoOverbookingCheck	bit,
			@sTemp				varchar(MAX),
			@piTableID			integer,
			@psRecordDescription	varchar(MAX),
			@psEmployeeName		varchar(MAX),
			@piResultCode			varchar(10)


	SET @piResultCode = ''
	SET @fPreReqsOverridden = 0
	SET @fUnavailOverridden = 0
	SET @fOverlapOverridden = 0
	SET @iNumberBooked = 0
	SET @fDoPreReqCheck = 0
	SET @fDoUnavailabilityCheck = 0
	SET @fDoOverlapCheck = 0
	SET @fDoOverbookingCheck = 0
	SET @psErrorMessage = ''
	SET @psPreReqCheckFailsCount = 0
	SET @psUnavailabilityCheckFailCount = 0
	SET @psOverlapCheckFailCount = 0
	SET @psCourseOverbooked = 0
	
	DECLARE @TempTable TABLE (EmployeeID INTEGER, EmployeeName VARCHAR(MAX), ResultCode VARCHAR(5))

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
	SET @piResultCode = ''
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

		IF @fDoPreReqCheck = 1
		BEGIN
				/* Return 0 if the given record in the personnel table has satisfied the pre-requisite criteria for the given course record.
				Return 1 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record.
				Return 2 if the given record in the personnel table has NOT satisfied the pre-requisite criteria for the given course record but the user can override this failure. */
				exec sp_ASR_TBCheckPreRequisites @piCourseRecordID,  @iEmployeeID, @iResult OUTPUT
				IF @iResult = 1 OR @iResult = 2
				BEGIN
	   				SET @psPreReqCheckFailsCount = @psPreReqCheckFailsCount + 1
				END
				SET @piResultCode = @piResultCode + CONVERT(varchar(10), @iResult)
		END

		IF @fDoUnavailabilityCheck = 1
		BEGIN
				exec sp_ASR_TBCheckUnavailability @piCourseRecordID,  @iEmployeeID, @iResult OUTPUT
				/* Return 0 if the given record in the personnel table IS available for the given course record.
				Return 1 if the given record in the personnel table is NOT available for the given course record.
				Return 2 if the given record in the personnel table is NOT available for the given course record but the user can override this failure. */
				IF @iResult = 1  OR @iResult = 2
				BEGIN
					SET @psUnavailabilityCheckFailCount = @psUnavailabilityCheckFailCount + 1
				END				
				 
				SET @piResultCode = @piResultCode + CONVERT(varchar(10), @iResult)
		END

		IF @fDoOverlapCheck = 1
		BEGIN
				exec sp_ASR_TBCheckOverlappedBooking @piCourseRecordID,  @iEmployeeID, 0, @iResult OUTPUT
				  /* Return 0 if the given course does NOT overlap with another course that the given delegate is booked on.
				  Return 1 if the given course DOES overlap with another course that the given delegate is booked on.
				  Return 2 if the given course does NOT overlap with another course that the given delegate is booked on, but the user can override this failure. */
				IF @iResult = 1  OR @iResult = 2
				BEGIN
					SET @psOverlapCheckFailCount = @psOverlapCheckFailCount + 1
				END
				SET @piResultCode = @piResultCode + CONVERT(varchar(10), @iResult)
		END

		INSERT INTO @TempTable VALUES (@iEmployeeID, @psEmployeeName, @piResultCode)

		SET @iNumberBooked = @iNumberBooked + 1

		SET @piResultCode = ''
	END

	IF LEN(@psErrorMessage) = 0 AND (@fDoOverbookingCheck = 1)
	BEGIN
		exec sp_ASR_TBCheckOverbooking @piCourseRecordID, 0, @iNumberBooked, @iResult OUTPUT

		SET @psCourseOverbooked = @iResult
	END

	SELECT EmployeeID, EmployeeName, ResultCode FROM @TempTable
END