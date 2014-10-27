CREATE PROCEDURE [dbo].[spASRGetWorkflowDelegates] 
(
	@psTo			varchar(MAX),
	@piStepID		integer,
	@results		cursor varying output
)
AS
BEGIN
	DECLARE
		@iDelegateEmailID	integer,
		@sTemp				varchar(MAX),
		@iDelegateRecordID	integer,
		@sDelegateTo		varchar(MAX),
		@iCount				integer,
		@sSQL				nvarchar(MAX),
		@iInstanceID		integer;

	IF len(ltrim(rtrim(@psTo))) = 0 RETURN;

    DECLARE @recipients TABLE (
        recordID		integer,
		emailAddress	varchar(MAX),
		delegated		bit,
		delegatedTo		varchar(MAX),
		processed		tinyint default 0,
		isDelegate		bit);
		
	-- Get the delegate email address definition. 
	SET @iDelegateEmailID = 0;
	SELECT @sTemp = ISNULL(parameterValue, '')
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_DelegateEmail'
	SET @iDelegateEmailID = convert(integer, @sTemp);
		
	IF @iDelegateEmailID = 0
	BEGIN
		INSERT INTO @recipients (
			recordID,
			emailAddress,
			delegated,
			delegatedTo,
			processed,
			isDelegate)
		VALUES (
			0, -- Personnel Record ID
			@psTo, -- Email Address(es)
			0, -- Delegated
			'', -- Delegate Email Address
			2, -- Processed
			0); -- Is Delegate
	END
	ELSE
	BEGIN
		INSERT INTO @recipients 
		SELECT         
			RECnew.ID,
			RECnew.emailAddress,
			RECnew.delegated,
			RECnew.delegatedTo,
			0,
			0
		FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@psTo) RECnew
		WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
			AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
				FROM @recipients RECold
				WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID);

		SELECT @iCount = COUNT(*)
		FROM @recipients
		WHERE processed = 0;

		WHILE @iCount > 0
		BEGIN
			-- Mark the new rows as 'being processed'.
			UPDATE @recipients
			SET processed = 1
			WHERE processed = 0;

			DECLARE delegatesCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT recordID
			FROM @recipients
			WHERE recordID > 0
				AND processed = 1
				AND delegated = 1;

			OPEN delegatesCursor;
			FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sDelegateTo = '';
				SET @sSQL = 'spASRSysEmailAddr';
	
				IF EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = @sSQL)
				BEGIN
					-- Get the delegate's email address
					EXEC @sSQL @sDelegateTo OUTPUT, @iDelegateEmailID, @iDelegateRecordID;
					IF @sDelegateTo IS null SET @sDelegateTo = '';
				END

				IF len(@sDelegateTo) > 0 
				BEGIN
					UPDATE @recipients 
					SET delegatedTo = @sDelegateTo
					WHERE recordID = @iDelegateRecordID;

					INSERT INTO @recipients 
					SELECT         
						RECnew.ID,
						RECnew.emailAddress,
						RECnew.delegated,
						RECnew.delegatedTo,
						0,
						1
					FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@sDelegateTo) RECnew
					WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
						AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
							FROM @recipients RECold
							WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID);
				END
				ELSE
				BEGIN
					UPDATE @recipients 
					SET delegated = 0
					WHERE recordID = @iDelegateRecordID;
				END

				FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID;
			END
			CLOSE delegatesCursor;
			DEALLOCATE delegatesCursor;

			-- Mark the processed rows as 'been processed'.
			UPDATE @recipients
			SET processed = 2
			WHERE processed = 1;

			SELECT @iCount = COUNT(*)
			FROM @recipients
			WHERE processed = 0;
		END
	END

	-- Return the cursor of succeeding elements. 
	SET @results = CURSOR FORWARD_ONLY STATIC FOR
        SELECT DISTINCT 
			emailAddress,
			delegated,
			delegatedTo,
			isDelegate
        FROM @recipients;

	OPEN @results;
END