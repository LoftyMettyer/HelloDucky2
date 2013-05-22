CREATE PROCEDURE [dbo].[sp_ASRIntGetPicklistDefinition] (
	@piPicklistID 			integer, 
	@psAction				varchar(255),
	@psErrorMsg				varchar(MAX)	OUTPUT,
	@psPicklistName			varchar(255)	OUTPUT,
	@psPicklistOwner		varchar(255)	OUTPUT,
	@psPicklistDesc			varchar(MAX)	OUTPUT,
	@psAccess				varchar(255)	OUTPUT,
 	@piTimestamp			integer		OUTPUT
)
AS
BEGIN
	DECLARE	@iCount		integer,
		@sCurrentUser	sysname,
		@fSysSecMgr		bit;

	SET @psErrorMsg = '';
	SET @sCurrentUser = SYSTEM_USER;

	exec spASRIntSysSecMgr @fSysSecMgr OUTPUT
	
	/* Check the picklist exists. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysPicklistName 
	WHERE picklistID = @piPicklistID

	IF @iCount = 0
	BEGIN
		SET @psErrorMsg = 'picklist has been deleted by another user.'
		RETURN
	END

	SELECT @psPicklistName = name,
		@psPicklistOwner = userName,
		@psPicklistDesc = description,
		@psAccess = access,
		@piTimestamp = convert(integer, timestamp)
	FROM ASRSysPicklistName 
	WHERE picklistID = @piPicklistID

	/* Check the current user can view the report. */
	IF (@psAccess = 'HD') AND (@psPicklistOwner <> @sCurrentUser) AND (@fSysSecMgr = 0)
	BEGIN
		SET @psErrorMsg = 'picklist has been made hidden by another user.'
		RETURN
	END

	IF (@psAction <> 'view') AND (@psAction <> 'copy') AND (@psAccess = 'RO') AND (@psPicklistOwner <> @sCurrentUser)  AND (@fSysSecMgr = 0)
	BEGIN
		SET @psErrorMsg = 'picklist has been made read only by another user.'
		RETURN
	END

	IF @psAction = 'copy' 
	BEGIN
		SET @psPicklistName = left('copy of ' + @psPicklistName, 50)
		SET @psPicklistOwner = @sCurrentUser
	END

	/* Get the picklist records. */
	SELECT recordID
	FROM ASRSysPickListItems
	WHERE pickListID = @piPicklistID
END



















GO

