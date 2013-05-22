CREATE PROCEDURE [dbo].[sp_ASRIntValidatePicklist2] (
	@psUtilName 		varchar(255),
	@piUtilID 			integer,
	@psAccess 			varchar(MAX),
	@piBaseTableID 		integer, 
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT /* 	0 = no errors, 
								1 = error, 
								2 = definition used in utilities owned by the current user. Prompt to make these hidden too. */
)
AS
BEGIN
	DECLARE	@iCount					integer,
			@sCurrentUser			sysname,
			@iHiddenCheckResult 	integer,
			@sHiddenCheckMessage	varchar(MAX);

	SELECT @sCurrentUser = SYSTEM_USER;
	SET @psErrorMsg = '';
	SET @piErrorCode = 0;

	/* Check that the picklist name is unique. */
	IF @piUtilID > 0
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM [dbo].[ASRSysPickListName]
		WHERE name = @psUtilName
			AND picklistID <> @piUtilID
			AND tableID = @piBaseTableID;
	END
	ELSE
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM [dbo].[ASRSysPickListName]
		WHERE name = @psUtilName
			AND tableID = @piBaseTableID;
	END

	IF @iCount > 0 
	BEGIN
		SET @psErrorMsg = 'A picklist called ''' + @psUtilName + ''' already exists.';
		SET @piErrorCode = 1;
	END

	IF (@piErrorCode = 0) AND (@psAccess = 'HD') AND (@piUtilID > 0)
	BEGIN
		/* Check that the picklist can be made hidden (ie. is not used in any utilities owned by other people. */
		exec [dbo].[sp_ASRIntCheckCanMakeHidden] 10, @piUtilID, @iHiddenCheckResult OUTPUT, @sHiddenCheckMessage OUTPUT;

		IF @iHiddenCheckResult = 1
		BEGIN
			/* picklist used only in utilities owned by the current user - we then need to prompt the user if they want to make these utilities hidden too. */
			SET @psErrorMsg = 'Changing the selected picklist to hidden will automatically make the following definition(s), of which you are the owner, hidden also :' + 
				'<BR><BR>' +
				@sHiddenCheckMessage +
				'<BR><BR>' +
				'Do you wish to continue ?';
			SET @piErrorCode = 2;
		END

		IF @iHiddenCheckResult = 2
		BEGIN
			/* picklist used in utilities which are in batch jobs not owned by the current user - Cannot therefore make the utility hidden. */
			SET @psErrorMsg = 'This picklist cannot be made hidden as it is used in definition(s) which are included in the following batch jobs of which you are not the owner :' + 
				'<BR><BR>' +
				@sHiddenCheckMessage;
			SET @piErrorCode = 1;
		END

		IF @iHiddenCheckResult = 3
		BEGIN
			/* picklist used in utilities which are not owned by the current user - Cannot therefore make the utility hidden. */
			SET @psErrorMsg = 'This picklist cannot be made hidden as it is used in the following definition(s), of which you are not the owner :' + 
				'<BR><BR>' +
				@sHiddenCheckMessage;
			SET @piErrorCode = 1;
		END

		IF @iHiddenCheckResult = 4
		BEGIN
			SET @psErrorMsg = 'This picklist cannot be made hidden as it is used in definition(s) which are included in the following batch jobs which are scheduled to be run by other user groups :' + 
				'<BR><BR>' +
				@sHiddenCheckMessage;
 				
			SET @piErrorCode = 1;
		END
	END
END