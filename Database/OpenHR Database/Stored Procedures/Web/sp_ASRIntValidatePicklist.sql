CREATE PROCEDURE [dbo].[sp_ASRIntValidatePicklist] (
	@psUtilName 		varchar(255), 
	@piUtilID 			integer, 
	@piTimestamp 		integer, 
	@psAccess 			varchar(MAX), 
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT /* 	0 = no errors, 
								1 = definition deleted or made read only by someone else,  but prompt to save as new definition 
								2 = definition changed by someone else, overwrite ? */
)
AS
BEGIN
	DECLARE	@iTimestamp			integer,
			@sAccess			varchar(MAX),
			@sOwner				varchar(255),
			@iCount				integer,
			@sCurrentUser		sysname;

	SELECT @sCurrentUser = SYSTEM_USER;
	SET @psErrorMsg = '';
	SET @piErrorCode = 0;

	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysPickListName]
		WHERE picklistID = @piUtilID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The picklist has been deleted by another user.<BR>Save as a new definition ?';
			SET @piErrorCode = 1;
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sAccess = access, 
				@sOwner = userName
			FROM [dbo].[ASRSysPickListName]
			WHERE picklistID = @piUtilID;

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The picklist has been amended by another user and is now Read Only.<BR>Save as a new definition ?';
					SET @piErrorCode = 1;
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The picklist has been amended by another user.<BR>Would you like to overwrite this definition ?';
					SET @piErrorCode = 2;
				END
			END
		END
	END
END