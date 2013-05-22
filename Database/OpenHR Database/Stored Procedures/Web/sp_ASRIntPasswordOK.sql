CREATE PROCEDURE [dbo].[sp_ASRIntPasswordOK]
AS
BEGIN
	/* Update the current user's record into ASRSysPassword table.. */
	DECLARE @iCount		integer,
		@sCurrentUser	sysname;

	SET @sCurrentUser = system_user;

	/* Check that the current user has a record in the table. */
	SELECT @iCount = COUNT(userName)
	FROM ASRSysPasswords
	WHERE userName = @sCurrentUser;

	IF @iCount = 0
	BEGIN
		INSERT INTO ASRSysPasswords (userName, lastChanged, forceChange)
		VALUES (@sCurrentUser, GETDATE(), 0);
	END
	ELSE
	BEGIN
		UPDATE ASRSysPasswords 
		SET lastChanged = GETDATE(), 
			forceChange = 0
		WHERE userName = @sCurrentUser;
	END
END