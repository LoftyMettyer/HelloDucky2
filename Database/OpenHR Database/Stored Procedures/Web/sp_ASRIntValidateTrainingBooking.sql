CREATE PROCEDURE [dbo].[sp_ASRIntValidateTrainingBooking] (
	@piResultCode		varchar(MAX) OUTPUT,
	@piEmpRecID		integer,
	@piCourseRecID		integer,
	@psBookingStatus	varchar(MAX),
	@piTBRecID		integer,
	@psCourseOverbooked integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Perform the Training Booking validation on the given insert/update SQL string.
	Return codes are :
		@piResultCode = '000' - completely valid
		If non-zero then the result code is composed as abc,
		where a is the result of the PRE-REQUISITES check
			b is the result of the AVAILABILITY check
			c is the result of the OVERLAPPED BOOKING check.
		the values of which can be :
			0 if the check PASSED
			1 if the check FAILED and CANNOT be overridden
			2 if the check FAILED but CAN be overridden

	The psCourseOverbooked parameter returns if the course is overbooked
	*/
	DECLARE	@fIncludeProvisionals	bit,
		@sIncludeProvisionals	varchar(MAX),
		@iCount					integer,
		@iResult				integer,
		@iTemp					integer,
		@piResultOverlapping   integer = 0,
		@piResultPrerequisites	integer = 0,
		@piResultUnavailability	integer = 0;

	SET @piResultCode = '';
	SET @psCourseOverbooked = 0;

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
				SET @psCourseOverbooked = @iResult -- @iResult = 1 -> Course fully booked (error). @iResult = 2 -> Course fully booked (over-rideable by the user).
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
					SET @piResultPrerequisites = @iResult -- @iResult = 1 -> Pre-requisites not satisfied (error). @iResult = 2 -> Pre-requisites not satisfied (over-rideable by the user). 
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
					SET @piResultUnavailability = @iResult -- @iResult = 1 -> Employee unavailable (error). @iResult = 2 -> Employee unavailable (over-rideable by the user).
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
				SET @piResultOverlapping = @iResult -- @iResult = 1 -> Overlapped booking (error). @iResult = 2 -> Overlapped booking (over-rideable by the user). 
			END
		END
		SET @piResultCode = CONVERT(VARCHAR(1), @piResultPrerequisites) + CONVERT(VARCHAR(1), @piResultUnavailability) + CONVERT(VARCHAR(1), @piResultOverlapping)
	END
END