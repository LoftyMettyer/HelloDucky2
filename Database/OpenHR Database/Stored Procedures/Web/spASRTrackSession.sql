﻿CREATE PROCEDURE dbo.[spASRTrackSession](
	@LoggingIn		bit,
	@Application	varchar(255),
	@ClientMachine	varchar(255),
	@LoginTime			datetime OUTPUT)
AS
BEGIN

	DECLARE @sUserName		nvarchar(MAX), 
			@sUserGroup		nvarchar(MAX),
			@iUserGroupID	integer;

	EXEC [dbo].[spASRIntGetActualUserDetails] @sUserName OUTPUT, @sUserGroup OUTPUT, @iUserGroupID OUTPUT

	IF @sUserGroup IS NULL
		SET @sUserGroup = '<Unknown>';

	DELETE FROM dbo.ASRSysCurrentLogins WHERE Username = @sUserName AND [clientmachine] = @ClientMachine;

	SET @LoginTime = GETDATE();

	IF @LoggingIn = 1
	BEGIN
			
		INSERT dbo.ASRSysCurrentLogins ([username], [usergroup], [usergroupid], [usersid], [loginTime], [application], clientmachine)
			VALUES (@sUserName, @sUserGroup, @iUserGroupID, USER_SID(), @LoginTime, @Application, @ClientMachine);

		INSERT INTO [dbo].[ASRSysAuditAccess]
			(DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) 
			VALUES (@LoginTime, @sUserGroup, @sUserName, LOWER(HOST_NAME()), 'Intranet', 'Log In');
	END
	ELSE
	BEGIN

		INSERT INTO [dbo].[ASRSysAuditAccess]
			(DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) 
			VALUES (@LoginTime, @sUserGroup, @sUserName, LOWER(HOST_NAME()), 'Intranet', 'Log Out');

	END

END