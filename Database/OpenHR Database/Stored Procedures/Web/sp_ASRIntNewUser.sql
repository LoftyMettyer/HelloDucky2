CREATE PROCEDURE [dbo].[sp_ASRIntNewUser] (
	@psUserName	sysname)
AS
BEGIN
	/* Create an user associated with the given SQL login. 
	Put the new user in the current user's role.
	Return 1 if everything is done okay, else 0. */
	DECLARE @hResult 		integer,
		@sRoleName			sysname,
		@sActualUserName	sysname,
		@iUserGroupID		integer;

	/* Create a user in the database for the given login. */
	EXEC @hResult = sp_grantdbaccess @psUsername, @psUserName;
	IF @hResult <> 0 GOTO Done

	/* Determine the current user's role. */
	EXEC dbo.spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sRoleName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Put the new user in the same role as the current user. */
	EXEC @hResult = sp_addrolemember @sRoleName, @psUserName;
	IF @hResult <> 0 GOTO Err;

	/* Make the new user a dbo. */
	EXEC @hResult = sp_addrolemember 'db_owner', @psUserName;
	IF @hResult <> 0 GOTO Err;

	/* Jump over the error handling code. */
	GOTO Done;

Err:
	/* Remove the user from the database if it was added okay, but not assigned to a role. */
	EXEC sp_revokedbaccess @psUsername;

Done:
	RETURN (@hResult);

END