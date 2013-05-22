CREATE PROCEDURE [dbo].[sp_ASRIntPoll] AS
BEGIN

	SET NOCOUNT ON;

	/* Update the ASRSysIntHit table to show that the database has been hit by an intranet user.
	Return a recordset of the messages for the user. */
	DECLARE @iCount		integer,
		@iDBID			integer,
		@iUID			integer,
		@dtLoginTime	datetime,
		@sLoginName		varchar(256);

	/* Check if the current user already has a record in the poll table. */
	SELECT @iCount = COUNT(*) 
	FROM [dbo].[ASRSysIntPoll]
	WHERE spid = @@spid;

	/* Get the current user's process information. */
	SELECT @iDBID = dbID,
		@iUID = uid,
		@dtLoginTime = login_time,
		@sLoginName = loginame
	FROM master..sysprocesses
	WHERE spid = @@spid;
	
	/* Create/update the current user's record in the poll table. */
	IF @iCount  = 0 
	BEGIN
		INSERT INTO [dbo].[ASRSysIntPoll] (spid, hitTime, dbID, uID, loginTime, loginName)
			VALUES (@@spid, getdate(), @iDBID, @iUID, @dtLoginTime, @sLoginName);
	END
	ELSE
	BEGIN
		UPDATE [dbo].[ASRSysIntPoll]
		SET hitTime = getdate(),
			dbID = @iDBID, 
			uID = @iUID, 
			loginTime = @dtLoginTime, 
			loginName = @sLoginName
		WHERE spid = @@spid;
	END

	/* Return a recordset of the messages for the current user. */
	SELECT 'Message from user ''' + ltrim(rtrim(messageFrom)) + 
		''' using ''' + ltrim(rtrim(messageSource)) + 
		' (' + convert(varchar(100), messageTime, 100) +')' + 
		char(10) + message
	FROM [dbo].[ASRSysMessages]
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND uid = @iUID
		AND loginTime = @dtLoginTime;

	/* Remove any orphaned messages. */
	DELETE
	FROM [dbo].[ASRSysMessages]
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND uid = @iUID
		AND loginTime = @dtLoginTime;
		
END