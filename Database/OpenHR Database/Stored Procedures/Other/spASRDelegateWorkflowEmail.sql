CREATE PROCEDURE [dbo].[spASRDelegateWorkflowEmail] 
(
	@psTo						varchar(MAX),
	@psCopyTo					varchar(MAX),
	@psMessage					varchar(MAX),
	@psMessage_HypertextLinks	varchar(MAX),
	@piStepID					integer,
	@psEmailSubject				varchar(MAX)
)
AS
BEGIN
	DECLARE	@sTo				varchar(MAX),
		@sAddress			varchar(MAX),
		@iInstanceID		integer,
		@curRecipients		cursor,
		@sEmailAddress		varchar(MAX),
		@fDelegated			bit,
		@sDelegatedTo		varchar(MAX),
		@fIsDelegate		bit,
		@sTemp		varchar(MAX),
		@fCopyDelegateEmail		bit;

	SET @psMessage = isnull(@psMessage, '');
	SET @psMessage_HypertextLinks = isnull(@psMessage_HypertextLinks, '');
	IF (len(ltrim(rtrim(@psTo))) = 0) RETURN;

	-- Get the instanceID of the given step
	SELECT @iInstanceID = instanceID
	FROM dbo.ASRSysWorkflowInstanceSteps
	WHERE ID = @piStepID;
		
    DECLARE @recipients TABLE (
		emailAddress	varchar(MAX),
		delegated		bit,
		delegatedTo		varchar(MAX),
		isDelegate		bit
    )

	exec [dbo].[spASRGetWorkflowDelegates] 
		@psTo, 
		@piStepID, 
		@curRecipients output;
		
	FETCH NEXT FROM @curRecipients INTO 
			@sEmailAddress,
			@fDelegated,
			@sDelegatedTo,
			@fIsDelegate
	WHILE (@@fetch_status = 0)
	BEGIN
		INSERT INTO @recipients
			(emailAddress,
			delegated,
			delegatedTo,
			isDelegate)
		VALUES (
			@sEmailAddress,
			@fDelegated,
			@sDelegatedTo,
			@fIsDelegate
		);
		
		FETCH NEXT FROM @curRecipients INTO 
				@sEmailAddress,
				@fDelegated,
				@sDelegatedTo,
				@fIsDelegate;
	END
	CLOSE @curRecipients;
	DEALLOCATE @curRecipients;

	-- Clear out the delegation record for the current step
	DELETE FROM [dbo].[ASRSysWorkflowStepDelegation]
	WHERE stepID = @piStepID;

	INSERT INTO [dbo].[ASRSysWorkflowStepDelegation] (delegateEmail, stepID)
	SELECT DISTINCT emailAddress, @piStepID
	FROM @recipients
	WHERE isDelegate = 1;

	SET @sTo = '';
	
	DECLARE toCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ltrim(rtrim(emailAddress))
	FROM @recipients
	WHERE len(ltrim(rtrim(emailAddress))) > 0
		AND delegated = 0
		AND ltrim(rtrim(emailAddress))  NOT IN
			(SELECT ltrim(rtrim(emailAddress))
			FROM @recipients
			WHERE len(ltrim(rtrim(emailAddress))) > 0
			AND delegated = 1);

	OPEN toCursor;
	FETCH NEXT FROM toCursor INTO @sAddress;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sTo = @sTo
			+ CASE 
				WHEN len(ltrim(rtrim(@sTo))) > 0 THEN ';'
				ELSE ''
			END 
			+ @sAddress;

		FETCH NEXT FROM toCursor INTO @sAddress;
	END
	CLOSE toCursor;
	DEALLOCATE toCursor;

	IF len(@sTo) > 0
	BEGIN
		INSERT [dbo].[ASRSysEmailQueue](
			RecordDesc,
			ColumnValue, 
			DateDue, 
			UserName, 
			[Immediate],
			RecalculateRecordDesc, 
			RepTo,
			MsgText,
			WorkflowInstanceID, 
			[Subject])
		VALUES ('',
			'',
			getdate(),
			'OpenHR Workflow',
			1,
			0, 
			@sTo,
			@psMessage + @psMessage_HypertextLinks,
			@iInstanceID,
			@psEmailSubject);
	END

	IF (len(@psCopyTo) > 0) AND (len(@psMessage) > 0)
	BEGIN
		INSERT ASRSysEmailQueue(
			RecordDesc,
			ColumnValue, 
			DateDue, 
			UserName, 
			[Immediate],
			RecalculateRecordDesc, 
			RepTo,
			MsgText,
			WorkflowInstanceID, 
			[Subject])
		VALUES ('',
			'',
			getdate(),
			'OpenHR Workflow',
			1,
			0, 
			@psCopyTo,
			'You have been copied in on the following OpenHR Workflow email with recipients:' + CHAR(13)
				+ CHAR(9) + @sTo + CHAR(13)	+ CHAR(13)
				+ @psMessage,
			@iInstanceID,
			@psEmailSubject);
	END

	SET @fCopyDelegateEmail = 1
	SELECT @sTemp = LTRIM(RTRIM(UPPER(ISNULL(parameterValue, ''))))
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_CopyDelegateEmail'

	IF @sTemp = 'TRUE'
	BEGIN
		DECLARE toCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ltrim(rtrim(emailAddress)), 
				ltrim(rtrim(delegatedTo))
			FROM @recipients
			WHERE len(ltrim(rtrim(emailAddress))) > 0
			AND delegated = 1;
			
		OPEN toCursor;
		FETCH NEXT FROM toCursor INTO @sAddress, @sDelegatedTo;
		WHILE (@@fetch_status = 0)
		BEGIN
			INSERT ASRSysEmailQueue(
				RecordDesc,
				ColumnValue, 
				DateDue, 
				UserName, 
				[Immediate],
				RecalculateRecordDesc, 
				RepTo,
				MsgText,
				WorkflowInstanceID, 
				[Subject])
			VALUES ('',
				'',
				getdate(),
				'OpenHR Workflow',
				1,
				0, 
				@sAddress,
				'The following email has been delegated to ' + @sDelegatedTo + char(13) + 
					'--------------------------------------------------' + char(13) +
					@psMessage + @psMessage_HypertextLinks,
				@iInstanceID,
				@psEmailSubject);

				
			FETCH NEXT FROM toCursor INTO @sAddress, @sDelegatedTo;
		END
		CLOSE toCursor;
		DEALLOCATE toCursor;
	END
END