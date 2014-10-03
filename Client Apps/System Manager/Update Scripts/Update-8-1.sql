/* --------------------------------------------------- */
/* Update the database from version 8.0 to version 8.1*/
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion int,
	@NVarCommand nvarchar(MAX),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @sSPCode nvarchar(MAX)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON;
SET @DBName = DB_NAME();

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '8.0') and (@sDBVersion <> '8.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion < 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRSendMessage]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRSendMessage];
EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRSendMessage] 
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
		AND (@psSPIDS = '''' OR charindex('' ''+convert(varchar,spid)+'' '', @psSPIDS)>0);

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

	IF OBJECT_ID(''tempdb..#tblCurrentUsers'', N''U'') IS NOT NULL
		DROP TABLE #tblCurrentUsers;

	-- Send message to all the web connections
	MERGE INTO ASRSysMessages AS Target
		USING (SELECT username, loginTime
			FROM ASRSysCurrentLogins) AS SOURCE (LoginName, loginTime)
	ON target.loginName = source.LoginName AND target.loginTime = source.loginTime
	WHEN MATCHED THEN
		UPDATE SET message = @psMessage
	WHEN NOT MATCHED BY TARGET THEN
		INSERT (LoginName, message, loginTime, messageTime, messageFrom, messageSource)
		VALUES (LoginName, @psMessage, loginTime, @currentDate, @sCurrentUser, @sCurrentApp)
	WHEN NOT MATCHED BY SOURCE THEN
		DELETE;

	-- Message to the Web Server
	INSERT INTO ASRSysMessages 
		(loginname, [message], loginTime, [dbid], [uid], spid, messageTime, messageFrom, messageSource) 
		VALUES(''OpenHR Web Server'', @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, @currentDate, @sCurrentUser, @sCurrentApp);

END'


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountInApp]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp];
EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
(
	@piCount integer OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @Mode			smallint;

	SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode'';
	IF @@ROWCOUNT = 0 SET @Mode = 0;

	IF (@Mode = 1 OR @Mode = 2) AND (NOT IS_SRVROLEMEMBER(''sysadmin'') = 1)
	BEGIN
		SELECT @piCount = dbo.[udfASRNetCountCurrentUsersInApp](APP_NAME());
	END
	ELSE
	BEGIN

		SELECT @piCount = COUNT(p.Program_Name)
		FROM     master..sysprocesses p
		JOIN     master..sysdatabases d
		  ON     d.dbid = p.dbid
		WHERE    p.program_name = APP_NAME()
		  AND    d.name = db_name()
		GROUP BY p.program_name;
	END

END'


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersAppName];
EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
(
	@psAppName		varchar(MAX) OUTPUT,
	@psUserName		varchar(MAX)
)
AS
BEGIN

    SELECT TOP 1 @psAppName = rtrim(p.program_name)
    FROM master..sysprocesses p
    WHERE p.program_name LIKE ''OpenHR%''
		AND	p.program_name NOT LIKE ''OpenHR Workflow%''
		AND	p.program_name NOT LIKE ''OpenHR Outlook%''
		AND	p.program_name NOT LIKE ''OpenHR Server.Net%''
		AND	p.program_name NOT LIKE ''OpenHR Intranet Embedding%''
		AND	p.loginame = @psUsername
    GROUP BY p.hostname
           , p.loginame
           , p.program_name
           , p.hostprocess
    ORDER BY p.loginame;

END'


/* ------------------------------------------------------- */
PRINT 'Step - XML Export Improvement'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XSDFileName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XSDFileName nvarchar(255) NULL, PreserveTransformPath bit, PreserveXSDPath bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitXMLNodesFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitXMLNodesFile bit;';

/* --------------------------------------------------------- */
PRINT 'Step - Update ASRSysCrossTab definition for 9-Box Grid'
/* --------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysCrossTab', 'U') AND name = 'CrossTabType') BEGIN
		EXEC sp_executesql N'
			ALTER TABLE ASRSysCrossTab ADD
									   CrossTabType tinyint, 
									   XAxisLabel nvarchar(255) NULL,
									   XAxisSubLabel1 nvarchar(255) NULL,
									   XAxisSubLabel2 nvarchar(255) NULL,
									   XAxisSubLabel3 nvarchar(255) NULL,
									   YAxisLabel nvarchar(255) NULL,
									   YAxisSubLabel1 nvarchar(255) NULL,
									   YAxisSubLabel2 nvarchar(255) NULL,
									   YAxisSubLabel3 nvarchar(255) NULL,
									   Description1 nvarchar(255) NULL,
									   ColorDesc1 INT NULL,
									   Description2 nvarchar(255) NULL,
									   ColorDesc2 INT NULL,
									   Description3 nvarchar(255) NULL,
									   ColorDesc3 INT NULL,
									   Description4 nvarchar(255) NULL,
									   ColorDesc4 INT NULL,
									   Description5 nvarchar(255) NULL,
									   ColorDesc5 INT NULL,
									   Description6 nvarchar(255) NULL,
									   ColorDesc6 INT NULL,
									   Description7 nvarchar(255) NULL,
									   ColorDesc7 INT NULL,
									   Description8 nvarchar(255) NULL,
									   ColorDesc8 INT NULL,
									   Description9 nvarchar(255) NULL,
									   ColorDesc9 INT NULL;
									   ';
        EXEC sp_executesql N'UPDATE ASRSysCrossTab SET CrossTabType = 0'; --'Normal' crosstab
	END

/* ------------------------------------------------------- */

/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '8.1';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.1.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.1.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.1')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
EXEC dbo.spsys_setsystemsetting 'database', 'refreshstoredprocedures', 1;


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF;

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.1 Of OpenHR'