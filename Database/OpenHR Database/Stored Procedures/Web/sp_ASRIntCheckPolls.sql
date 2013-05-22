CREATE PROCEDURE [dbo].[sp_ASRIntCheckPolls] AS
BEGIN
	DECLARE	@iSPID				integer,
			@sExecSQL			nvarchar(MAX),
			@iDBID				integer, 
			@iID				integer, 
			@dtLoginTime		datetime, 
			@sLoginName			varchar(256),
			@iCount				integer,
			@UserGroupName		varchar(256),
			@iUserGroupID		integer,
			@sActualUserName	sysname;

	SET NOCOUNT ON;

	IF IS_SRVROLEMEMBER('processadmin') = 1 and @@trancount = 0
	BEGIN
	
		/* Kill any intranet processes that have not been polled for a while. */
		DECLARE hits_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ip.spid,
			ip.dbID,
			ip.loginTime,
			ip.loginName
		FROM ASRSysIntPoll ip
		WHERE (ip.hitTime < dateadd(second, -45, getdate()))
		OR (ip.loginTime > (
			SELECT l.Lock_Time FROM
			(
				SELECT TOP 1 * 
				FROM ASRSysLock
				WHERE Priority < 3
				ORDER BY Lock_Time
			) AS l
			INNER JOIN master..sysprocesses p
			ON p.spid = l.spid
			AND p.dbID = ip.dbid));
			
		OPEN hits_cursor;
		FETCH NEXT FROM hits_cursor INTO @iSPID, @iDBID, @dtLoginTime, @sLoginName;

		WHILE (@@fetch_status = 0)
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM master..sysprocesses
			WHERE spid = @iSPID
				AND dbID = @iDBID
				AND login_time = @dtLoginTime
				AND loginame = @sLoginName
				AND ((program_name = 'OpenHR Intranet') OR (program_name = 'OpenHR Self-service Intranet'))
				AND status = 'sleeping'
				AND (last_batch < dateadd(second, -45, getdate()))
				OR (login_Time > (SELECT TOP 1 Lock_Time FROM ASRSysLock WHERE Priority < 3 ORDER BY Priority, Lock_Time));

			IF @iCount > 0
			BEGIN
				SET @sExecSQL = 'KILL ' + convert(varchar(MAX), @iSPID);
				EXECUTE sp_executeSQL @sExecSQL;

				/* Get the current user's group ID. */
				EXEC [dbo].[spASRIntGetActualUserDetailsForLogin]
					@sLoginName,
					@sActualUserName OUTPUT,
					@UserGroupName OUTPUT,
					@iUserGroupID OUTPUT;
						
				IF @UserGroupName IS null SET @UserGroupName = '<Unknown>';
						
				SET @sExecSQL = 'INSERT INTO AsrSysAuditAccess (DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action)
				                 VALUES (GetDate(), '''+replace(@UserGroupName,'''','''''')+''', '''+replace(rtrim(@sLoginName),'''','''''')+''', LOWER(HOST_NAME()), ''Intranet'', ''Log Out'')';
				EXECUTE sp_executeSQL @sExecSQL;
			END

			DELETE FROM [dbo].[ASRSysIntPoll]
				WHERE spid = @iSPID;

			FETCH NEXT FROM hits_cursor INTO @iSPID, @iDBID, @dtLoginTime, @sLoginName;
		END
		
		CLOSE hits_cursor;
		DEALLOCATE hits_cursor;

		/* Remove any orphaned messages. */
		DECLARE messages_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT id,
			loginName, 
			dbID, 
			loginTime 
		FROM ASRSysMessages;

		OPEN messages_cursor;	
		FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime;
		WHILE (@@fetch_status = 0)
		BEGIN
			SELECT @iCount = COUNT(loginame) 
			FROM master..sysprocesses
			WHERE loginame =  @sLoginName
				AND dbID = @iDBID
				AND login_time = @dtLoginTime;

			IF @iCount = 0
			BEGIN
				DELETE FROM ASRSysMessages 
				WHERE id = @iID;
			END
				
			FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime;
		END
		
		CLOSE messages_cursor;
		DEALLOCATE messages_cursor;

	END

END