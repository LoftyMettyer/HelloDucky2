CREATE PROCEDURE [dbo].[sp_ASRIntAuditAccess]
(
	@blnLoggingIn bit,
	@strUsername varchar(1000)
)
AS
BEGIN
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@sActualUserName	sysname;
		
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	IF @sUserGroupName IS NULL
	BEGIN
		SET @sUserGroupName = '<Unknown>';
	END

	/* Put an entry in the Audit Access Log */
	IF @blnLoggingIn <> 0
	BEGIN
		INSERT INTO [dbo].[ASRSysAuditAccess]
			(DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) 
			VALUES (GetDate(), @sUserGroupName, @strUserName, LOWER(HOST_NAME()), 'Intranet', 'Log In');
	END
	ELSE
	BEGIN
		INSERT INTO [dbo].[ASRSysAuditAccess]
			(DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) 
			VALUES (GetDate(), @sUserGroupName, @strUserName, LOWER(HOST_NAME()), 'Intranet', 'Log Out');
	END
END