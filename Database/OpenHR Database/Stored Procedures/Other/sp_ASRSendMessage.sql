CREATE PROCEDURE [dbo].[sp_ASRSendMessage] 
(
	@psMessage	varchar(MAX),
	@psSPIDS	varchar(MAX)
)
AS
BEGIN
	DECLARE @iDBid		integer,
		@iSPid			integer,
		@iUid			integer,
		@sLoginName		varchar(256),
		@dtLoginTime	datetime, 
		@sCurrentUser	varchar(256),
		@sCurrentApp	varchar(256),
		@Realspid		integer;

		DECLARE @currentDate	datetime = GETDATE();

	CREATE TABLE #tblCurrentUsers				
		(
			hostname varchar(256)
			,loginame varchar(256)
			,program_name varchar(256)
			,hostprocess varchar(20)
			,sid binary(86)
			,login_time datetime
			,spid int
			,uid smallint);
			
	INSERT INTO #tblCurrentUsers
		EXEC spASRGetCurrentUsers;

	--Need to get spid of parent process
	SELECT @Realspid = a.spid
	FROM #tblCurrentUsers a
	FULL OUTER JOIN #tblCurrentUsers b
		ON a.hostname = b.hostname
		AND a.hostprocess = b.hostprocess
		AND a.spid <> b.spid
	WHERE b.spid = @@Spid;

	--If there is no parent spid then use current spid
	IF @Realspid is null SET @Realspid = @@spid;

	/* Get the process information for the current user. */
	SELECT @iDBid = db_id(), 
		@sCurrentUser = loginame,
		@sCurrentApp = program_name
	FROM #tblCurrentUsers
	WHERE spid = @@Spid;

	/* Get a cursor of the other logged in users. */
	DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT spid, loginame, uid, login_time
		FROM #tblCurrentUsers
		WHERE (spid <> @@spid and spid <> @Realspid)
		AND (@psSPIDS = '' OR charindex(' '+convert(varchar,spid)+' ', @psSPIDS)>0);

	OPEN logins_cursor;
	FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
	WHILE (@@fetch_status = 0)
	BEGIN
		/* Create a message record for each user. */
		INSERT INTO ASRSysMessages 
			(loginname, [message], loginTime, [dbid], [uid], spid, messageTime, messageFrom, messageSource) 
			VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, @currentDate, @sCurrentUser, @sCurrentApp);

		FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
	END
	CLOSE logins_cursor;
	DEALLOCATE logins_cursor;

	IF OBJECT_ID('tempdb..#tblCurrentUsers', N'U') IS NOT NULL
		DROP TABLE #tblCurrentUsers;

	-- Message to the Web Server
	DELETE FROM ASRSysMessages WHERE loginname = 'OpenHR Web Server';

	INSERT INTO ASRSysMessages 
		(loginname, [message], loginTime, [dbid], [uid], spid, messageTime, messageFrom, messageSource) 
		VALUES('OpenHR Web Server', @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, @currentDate, @sCurrentUser, @sCurrentApp);

END