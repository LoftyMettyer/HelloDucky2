CREATE PROCEDURE [dbo].[spASRIntGetMessages]
	(@Logintime as datetime)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iUID		integer,
		@sLoginName		varchar(255),
		@sUserGroup		varchar(255);

	-- Get security info for this user
	EXEC [dbo].[spASRIntGetActualUserDetails] @sLoginName OUTPUT, @sUserGroup OUTPUT, @iUID OUTPUT

	-- Return a recordset of the messages for the current user.
	SELECT messagetime, messageFrom, [message] , messageSource
		FROM [dbo].[ASRSysMessages]
		WHERE loginName = @sLoginName	AND loginTime = @Logintime;

	-- Remove any orphaned messages.
	DELETE
	FROM [dbo].[ASRSysMessages]
	WHERE loginName = @sLoginName	AND loginTime = @Logintime;

END