CREATE PROCEDURE [dbo].[sp_ASRGetMessages]
AS
BEGIN
	DECLARE @iDBID		integer,
		@iID			integer,
		@dtLoginTime	datetime,
		@sLoginName		varchar(256),
		@iCount			integer,
		@Realspid		integer;

	-- Need to get spid of parent process
	SELECT @Realspid = a.spid
	FROM master..sysprocesses a
	FULL OUTER JOIN master..sysprocesses b
		ON a.hostname = b.hostname
		AND a.hostprocess = b.hostprocess
		AND a.spid <> b.spid
	WHERE b.spid = @@Spid;

	-- If there is no parent spid then use current spid
	IF @Realspid is null SET @Realspid = @@spid;

	-- Get the current user's process information.
	SELECT @iDBID = dbID,
		@dtLoginTime = login_time,
		@sLoginName = loginame
	FROM master..sysprocesses
	WHERE spid = @Realspid;

	-- Return the recordset of messages.
	SELECT 'Message from user ''' + ltrim(rtrim(messageFrom)) + 
		''' using ' + ltrim(rtrim(messageSource)) + 
		' (' + convert(varchar(100), messageTime, 100) +')' + 
		char(10) + message
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime
		AND spid = @Realspid;

	-- Remove any messages that have just been picked up.
	DELETE
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime
		AND spid = @Realspid;

END