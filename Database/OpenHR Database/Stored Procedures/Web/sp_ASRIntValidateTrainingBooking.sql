CREATE PROCEDURE [dbo].[sp_ASRIntValidateTrainingBooking] (
	@piResultCode		integer OUTPUT,
	@piEmpRecID			integer,
	@piCourseRecID		integer,
	@psBookingStatus	varchar(MAX),
	@piTBRecID			integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Perform the Training Booking validation on the given insert/update SQL string.
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
	DECLARE @fOK				bit,
		@fIncludeProvisionals	bit,
		@sIncludeProvisionals	varchar(MAX),
		@iCount					integer,
		@iResult				integer,
		@iTemp					integer;

	SET @fOK = 1
	SET @piResultCode = 0

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'TRAINING', @fOK OUTPUT
		
	IF @fOK = 0
	BEGIN
		/* Do not perform any training Booking checks if the module is not licenced. */
		RETURN
	END

	IF (@piCourseRecID > 0) AND ((@psBookingStatus = 'B') OR (@psBookingStatus = 'P'))
	BEGIN  
		SELECT @sIncludeProvisionals = parameterValue
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseIncludeProvisionals'
		IF @sIncludeProvisionals IS NULL SET @sIncludeProvisionals = 'FALSE'
		IF @sIncludeProvisionals = 'FALSE'
		BEGIN
			SET @sIncludeProvisionals = 0
		END
		ELSE
		BEGIN
			SET @sIncludeProvisionals = 1
		END

		/* Only check that the selected course is not fully booked if the new booking is included in the number booked. */
		IF (@fIncludeProvisionals = 1) OR (@psBookingStatus = 'B') 
		BEGIN
			/* Check if the overbooking stored procedure exists. */
			SELECT @iCount = COUNT(*) 
			FROM sysobjects
			WHERE id = object_id('sp_ASR_TBCheckOverbooking')
				AND sysstat & 0xf = 4

			IF @iCount > 0
			BEGIN
				exec sp_ASR_TBCheckOverbooking @piCourseRecID, @piTBRecID, 1, @iResult OUTPUT

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
      
		IF @piEmpRecID > 0
		BEGIN
			/* Check that the employee has satisfied the pre-requisite criteria for the selected course. */
			/* First check if the pre-requisite table is configured. If not, we do not need to do the pre-req check. */
			SELECT @iTemp = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_PreReqTable'
			IF @iTemp IS NULL SET @iTemp = 0

			IF @iTemp > 0 
			BEGIN
				/* Check if the pre-req stored procedure exists. */
				SELECT @iCount = COUNT(*) 
				FROM sysobjects
				WHERE id = object_id('sp_ASR_TBCheckPreRequisites')
					AND sysstat & 0xf = 4

				IF @iCount > 0
				BEGIN
					exec sp_ASR_TBCheckPreRequisites @piCourseRecID, @piEmpRecID, @iResult OUTPUT

					IF @iResult = 1 /* Pre-requisites not satisfied (error). */
					BEGIN
						SET @piResultCode = @piResultCode + 100
					END
					IF @iResult = 2 /* Pre-requisites not satisfied (over-rideable by the user). */
					BEGIN
						SET @piResultCode = @piResultCode + 200
					END 
				END
			END

			/* Check that the employee is available for the selected course. */
			/* First check if the unavailability table is configured. If not, we do not need to do the unavailability check. */
			SELECT @iTemp = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_TRAININGBOOKING'
				AND parameterKey = 'Param_UnavailTable'
			IF @iTemp IS NULL SET @iTemp = 0

			IF @iTemp > 0 
			BEGIN
				/* Check if the unavailability stored procedure exists. */
				SELECT @iCount = COUNT(*) 
				FROM sysobjects
				WHERE id = object_id('sp_ASR_TBCheckUnavailability')
					AND sysstat & 0xf = 4

				IF @iCount > 0
				BEGIN
					exec sp_ASR_TBCheckUnavailability @piCourseRecID, @piEmpRecID, @iResult OUTPUT

					IF @iResult = 1 /* Employee unavailable (error). */
					BEGIN
						SET @piResultCode = @piResultCode + 10
					END
					IF @iResult = 2 /* Employee unavailable (over-rideable by the user). */
					BEGIN
						SET @piResultCode = @piResultCode + 20
					END 
				END
			END

			/* Check if the overlapped booking stored procedure exists. */
			SELECT @iCount = COUNT(*) 
			FROM sysobjects
			WHERE id = object_id('sp_ASR_TBCheckOverlappedBooking')
				AND sysstat & 0xf = 4

			IF @iCount > 0
			BEGIN
				exec sp_ASR_TBCheckOverlappedBooking @piCourseRecID, @piEmpRecID, @piTBRecID, @iResult OUTPUT
				IF @iResult = 1 /* Overlapped booking (error). */
				BEGIN
					SET @piResultCode = @piResultCode + 1
				END
				IF @iResult = 2 /* Overlapped booking (over-rideable by the user). */
				BEGIN
					SET @piResultCode = @piResultCode + 2
				END 
			END
		END
	END
END