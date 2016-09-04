CREATE PROCEDURE [dbo].[spASRIntValidateOrganisationReport] (
	@psUtilName 		varchar(255), 
	@piUtilID 			integer, 
	@piTimestamp 		integer, 
	@piBaseViewID		integer, 	
	@piCategoryID 		integer,
	@psErrorMsg			varchar(MAX)	OUTPUT,
	@piErrorCode		varchar(MAX)	OUTPUT /* 	0 = no errors, 
								1 = error, 
								2 = definition deleted or made read only by someone else,  but prompt to save as new definition
								3 = definition changed by someone else, overwrite ? */
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@iTimestamp				integer,
			@sAccess				varchar(MAX),
			@sOwner					varchar(255),
			@iCount					integer,
			@sCurrentUser			sysname,					
			@fSysSecMgr				bit;
	
	SET @psErrorMsg = '';
	SET @piErrorCode = 0;	
	SELECT @sCurrentUser = SYSTEM_USER;

	exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;
	
	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysOrganisationReport
		WHERE ID = @piUtilID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The organisation report has been deleted by another user. Save as a new definition ?';
			SET @piErrorCode = 2;
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sOwner = userName
			FROM ASRSysOrganisationReport
			WHERE ID = @piUtilID;

			IF (@iTimestamp <> @piTimestamp)
			BEGIN
				exec spASRIntCurrentUserAccess 
					39, 
					@piUtilID,
					@sAccess	OUTPUT;
		
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW') AND (@iTimestamp <>@piTimestamp)
				BEGIN
					SET @psErrorMsg = 'The organisation report has been amended by another user and is now Read Only. Save as a new definition ?';
					SET @piErrorCode = 2;
				END
				ELSE
				BEGIN
					SET @psErrorMsg = 'The organisation report has been amended by another user. Would you like to overwrite this definition ?';
					SET @piErrorCode = 3;
				END
			END
			
		END
	END

	IF @piErrorCode = 0
	BEGIN
		/* Check that the report name is unique. */
		IF @piUtilID > 0
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSysOrganisationReport
			WHERE name = @psUtilName
				AND ID <> @piUtilID;
		END
		ELSE
		BEGIN
			SELECT @iCount = COUNT(*) 
			FROM ASRSysOrganisationReport
			WHERE name = @psUtilName;
		END

		IF @iCount > 0 
		BEGIN
			SET @psErrorMsg = 'An organisation report called ''' + @psUtilName + ''' already exists.';
			SET @piErrorCode = 1;
		END
	END

	IF (@piErrorCode = 0) AND (@piBaseViewID > 0)
	BEGIN
		/* Check that the Base View exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysViews 
		WHERE ViewID = @piBaseViewID;

		IF @iCount = 0
		BEGIN
			SET @psErrorMsg = 'The base view has been deleted by another user.';
			SET @piErrorCode = 1;
		END		
	END	
	

	IF (@piErrorCode = 0) AND (@piCategoryID > 0)
	BEGIN
		/* Check that the category exists. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysCategories
		WHERE id = @piCategoryID;

		IF @iCount = 1
		BEGIN
			SET @psErrorMsg = 'The category has been deleted by another user.';
			SET @piErrorCode = 1;
		END
	END	

END
