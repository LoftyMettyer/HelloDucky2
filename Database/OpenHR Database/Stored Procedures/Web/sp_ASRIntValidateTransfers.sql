CREATE PROCEDURE [dbo].[sp_ASRIntValidateTransfers] (
	@piEmployeeTableID			integer,
	@piCourseTableID			integer,
	@piCourseRecordID			integer,
	@piTransferCourseRecordID	integer,
	@piTrainBookTableID			integer,
	@piTrainBookStatusColumnID	integer,
	@piResultCode				integer			OUTPUT,
	@psErrorMessage				varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	/* This stored procedure run the pre-requisite, overbooking, overlapped booking and unavailability checks
	on booking being transferred from one course (@piCourseRecordID) to another (@piTransferCourseRecordID). 
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
	DECLARE	@sCommand				nvarchar(MAX),
			@sParamDefinition		nvarchar(500),
			@iUserGroupID			integer,
			@sUserGroupName			sysname,
			@iChildViewID			integer,
			@fTransferProvisionals	bit,
			@sTemp					varchar(MAX),
			@fPreReqsOverridden		bit,
			@fUnavailOverridden		bit,
			@fOverlapOverridden		bit,
			@iCount					integer,
			@iNumberBooked			integer,
			@iResult				integer,
			@sTBTableName			sysname,
			@sTrainBookRealSource	sysname,
			@sTrainBookStatusColumnName		sysname,
			@fDoPreReqCheck			bit,
			@iPreReqTableID			integer,
			@fDoUnavailabilityCheck	bit,
			@iUnavailTableID		integer,
			@fDoOverlapCheck		bit,
			@fDoOverbookingCheck	bit,
			@sActualUserName		sysname;

	SET @piResultCode = 0
	SET @psErrorMessage = ''
	SET @fPreReqsOverridden = 0
	SET @fUnavailOverridden = 0
	SET @fOverlapOverridden = 0
	SET @iNumberBooked = 0
	SET @fDoPreReqCheck = 0
	SET @fDoUnavailabilityCheck = 0
	SET @fDoOverlapCheck = 0
	SET @fDoOverbookingCheck = 0

	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	SELECT @sTBTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @piTrainBookTableID

	/* Get the realSource of the training booking table. */
	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @piTrainBookTableID
		AND role = @sUserGroupName
		
	IF @iChildViewID IS null SET @iChildViewID = 0
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sTrainBookRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_')
		SET @sTrainBookRealSource = left(@sTrainBookRealSource, 255)
	END
	ELSE
	BEGIN
		SET @psErrorMessage = 'Unable to determine the Training Booking child view.'
	END

	IF LEN(@psErrorMessage) = 0
	BEGIN
		IF @piTrainBookStatusColumnID > 0
		BEGIN
			SELECT @sTrainBookStatusColumnName = columnName
				FROM ASRSysColumns
				WHERE columnID = @piTrainBookStatusColumnID
			IF @sTrainBookStatusColumnName IS NULL SET @psErrorMessage = 'Unable to find the Training Booking Status column.'
		END
	END

	IF LEN(@psErrorMessage) = 0 
	BEGIN
		/* Check if we need to transfer provisional bookings. */
		SET @sTemp = ''
		SELECT @sTemp = convert(varchar(MAX), parameterValue)
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_TRAININGBOOKING'
			AND parameterKey = 'Param_CourseTransferProvisionals'
		IF @sTemp IS NULL SET @sTemp = ''
		IF @sTemp = 'TRUE'
		BEGIN
			SET @fTransferProvisionals = 1
		END
		ELSE
		BEGIN
			SET @fTransferProvisionals = 0
		END
	END

	IF LEN(@psErrorMessage) = 0 
	BEGIN
		/* Check if the overbooking stored procedure exists. */
		SELECT @iCount = COUNT(*) 
		FROM sysobjects
		WHERE id = object_id('sp_ASR_TBCheckOverbooking')
			AND sysstat & 0xf = 4

		IF @iCount > 0 SET @fDoOverbookingCheck = 1

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
	END

	IF LEN(@psErrorMessage) = 0
	BEGIN
		SET @sCommand = 
			'DECLARE @iEmployeeID	integer,' +
			'	@sStatus		varchar(MAX),' + 
			'	@iResult		integer,' +
			'	@fFailure		bit,' +
			'	@sCodeString		varchar(10),' +
			'	@sCurrentCode		varchar(1)' +
			' SET @piResultCode = 0' +
			' SET @fFailure = 0' +
			' SET @piNumberBooked = 0' +
			' DECLARE transfersCursor CURSOR LOCAL FAST_FORWARD FOR ' + 
			' SELECT id_' + convert(nvarchar(100), @piEmployeeTableID) + 
			', ' + @sTrainBookStatusColumnName +
			' FROM ' + @sTrainBookRealSource +
			' WHERE id_' + convert(nvarchar(100), @piCourseTableID) + ' = ' + convert(nvarchar(100), @piCourseRecordID)

		IF @fTransferProvisionals = 1
		BEGIN
			SET @sCommand = @sCommand +
				' AND (LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B''' +
				' OR LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''P'')'
		END
		ELSE
		BEGIN
			SET @sCommand = @sCommand +
				' AND LEFT(UPPER(' + @sTrainBookStatusColumnName + '), 1) = ''B'''
		END

		SET @sCommand = @sCommand  + 
			' OPEN transfersCursor' +
			' FETCH NEXT FROM transfersCursor INTO @iEmployeeID, @sStatus' +
			' WHILE (@@fetch_status = 0) AND (@fFailure = 0)' +
			' BEGIN' +
			'	SET @sCodeString = convert(varchar(10), @piResultCode)'

		IF @fDoPreReqCheck = 1
		BEGIN
			SET @sCommand = @sCommand  + 
				'	IF LEN(@sCodeString) >= 3' +
				'	BEGIN' + 
				'		SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 3), 1)' +
				'	END' +
				'	ELSE SET @sCurrentCode = 0' +
				'	IF @sCurrentCode <> ''1''' +
				'	BEGIN' +
				'		exec sp_ASR_TBCheckPreRequisites ' + convert(nvarchar(100), @piTransferCourseRecordID) + ',  @iEmployeeID, @iResult OUTPUT' +
				'		IF @iResult = 1 SET @fFailure = 1 /* Pre-requisites not satisfied (error). */' +
				'		SET @piResultCode = @piResultCode - (100 * convert(integer, @sCurrentCode)) + (100 * @iResult)' +
				'	END'
		END

		IF @fDoUnavailabilityCheck = 1
		BEGIN
			SET @sCommand = @sCommand  + 
				'	IF len(@sCodeString) >= 2' +
				'	BEGIN' + 
				'		SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 2), 1)' +
				'	END' +
				'	ELSE SET @sCurrentCode = 0' +
				'	IF @sCurrentCode <> ''1''' +
				'	BEGIN' +
				'		exec sp_ASR_TBCheckUnavailability ' + convert(nvarchar(100), @piTransferCourseRecordID) + ',  @iEmployeeID, @iResult OUTPUT' +
				'		IF @iResult = 1 SET @fFailure = 1 /* Unavailability check not satisfied (error). */' +
				'		SET @piResultCode = @piResultCode - (10 * convert(integer, @sCurrentCode)) + (10 * @iResult)' +
				'	END'
		END

		IF @fDoOverlapCheck = 1
		BEGIN
			SET @sCommand = @sCommand  + 
				'	IF len(@sCodeString) >= 1' +
				'	BEGIN' + 
				'		SET @sCurrentCode = LEFT(RIGHT(@sCodeString, 1), 1)' +
				'	END' +
				'	ELSE SET @sCurrentCode = 0' +
				'	IF @sCurrentCode <> ''1''' +
				'	BEGIN' +
				'		exec sp_ASR_TBCheckOverlappedBooking ' + convert(nvarchar(100), @piTransferCourseRecordID) + ',  @iEmployeeID, 0, @iResult OUTPUT' +
				'		IF @iResult = 1 SET @fFailure = 1 /* Overlapped booking (error). */' +
				'		SET @piResultCode = @piResultCode - convert(integer, @sCurrentCode) + @iResult' +
				'	END'
		END

		IF @fTransferProvisionals = 1
		BEGIN
			SET @sCommand = @sCommand +
				' SET @piNumberBooked = @piNumberBooked + 1'
		END
		ELSE
		BEGIN
			SET @sCommand = @sCommand +
				' IF @sStatus = ''B'' SET @piNumberBooked = @piNumberBooked + 1'
		END

		SET @sCommand = @sCommand  + 
			'	FETCH NEXT FROM transfersCursor INTO @iEmployeeID, @sStatus' +
			' END' +
			' CLOSE transfersCursor' +
			' DEALLOCATE transfersCursor'

		SET @sParamDefinition = N'@piResultCode integer OUTPUT, @piNumberBooked integer OUTPUT'
		EXEC sp_executesql @sCommand,  @sParamDefinition, @piResultCode OUTPUT, @iNumberBooked OUTPUT
	END

	IF LEN(@psErrorMessage) = 0 AND (@fDoOverbookingCheck = 1)
	BEGIN
		exec sp_ASR_TBCheckOverbooking @piTransferCourseRecordID, 0, @iNumberBooked, @iResult OUTPUT

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