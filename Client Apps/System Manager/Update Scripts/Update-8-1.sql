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
									   XAxisLabel varchar(255) NULL,
									   XAxisSubLabel1 varchar(255) NULL,
									   XAxisSubLabel2 varchar(255) NULL,
									   XAxisSubLabel3 varchar(255) NULL,
									   YAxisLabel varchar(255) NULL,
									   YAxisSubLabel1 varchar(255) NULL,
									   YAxisSubLabel2 varchar(255) NULL,
									   YAxisSubLabel3 varchar(255) NULL,
									   Description1 varchar(255) NULL,
	 									 ColorDesc1 varchar(6) NULL,
										 Description2 varchar(255) NULL,
										 ColorDesc2 varchar(6) NULL,
										 Description3 varchar(255) NULL,
										 ColorDesc3 varchar(6) NULL,
										 Description4 varchar(255) NULL,
										 ColorDesc4 varchar(6) NULL,
										 Description5 varchar(255) NULL,
										 ColorDesc5 varchar(6) NULL,
										 Description6 varchar(255) NULL,
										 ColorDesc6 varchar(6) NULL,
										 Description7 varchar(255) NULL,
										 ColorDesc7 varchar(6) NULL,
										 Description8 varchar(255) NULL,
										 ColorDesc8 varchar(6) NULL,
										 Description9 varchar(255) NULL,
										 ColorDesc9 varchar(6) NULL;
									   ';
        EXEC sp_executesql N'UPDATE ASRSysCrossTab SET CrossTabType = 0'; --'Normal' crosstab
	END
	
	-- Insert the system permissions for 9-Box Grid Reports and new picture too
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 45)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [picture])
			VALUES (45, '9-Box Grid Reports', 10, 'NINEBOXGRID',0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000400100000000000000000000000100000000000000000000800080008000000080800000008000000080800000008000C0C0C000C0DCC000F0CAA60080808000FF00FF00FF000000FFFF000000FF000000FFFF000000FF00FFFFFF00F0FBFF00A4A0A000D4F0FF00B1E2FF008ED4FF006BC6FF0048B8FF0025AAFF0000AAFF000092DC00007AB90000629600004A730000325000D4E3FF00B1C7FF008EABFF006B8FFF004873FF002557FF000055FF000049DC00003DB900003196000025730000195000D4D4FF00B1B1FF008E8EFF006B6BFF004848FF002525FF000000FF000000DC000000B900000096000000730000005000E3D4FF00C7B1FF00AB8EFF008F6BFF007348FF005725FF005500FF004900DC003D00B900310096002500730019005000F0D4FF00E2B1FF00D48EFF00C66BFF00B848FF00AA25FF00AA00FF009200DC007A00B900620096004A00730032005000FFD4FF00FFB1FF00FF8EFF00FF6BFF00FF48FF00FF25FF00FF00FF00DC00DC00B900B900960096007300730050005000FFD4F000FFB1E200FF8ED400FF6BC600FF48B800FF25AA00FF00AA00DC009200B9007A009600620073004A0050003200FFD4E300FFB1C700FF8EAB00FF6B8F00FF487300FF255700FF005500DC004900B9003D00960031007300250050001900FFD4D400FFB1B100FF8E8E00FF6B6B00FF484800FF252500FF000000DC000000B9000000960000007300000050000000FFE3D400FFC7B100FFAB8E00FF8F6B00FF734800FF572500FF550000DC490000B93D0000963100007325000050190000FFF0D400FFE2B100FFD48E00FFC66B00FFB84800FFAA2500FFAA0000DC920000B97A000096620000734A000050320000FFFFD400FFFFB100FFFF8E00FFFF6B00FFFF4800FFFF2500FFFF0000DCDC0000B9B90000969600007373000050500000F0FFD400E2FFB100D4FF8E00C6FF6B00B8FF4800AAFF2500AAFF000092DC00007AB90000629600004A73000032500000E3FFD400C7FFB100ABFF8E008FFF6B0073FF480057FF250055FF000049DC00003DB90000319600002573000019500000D4FFD400B1FFB1008EFF8E006BFF6B0048FF480025FF250000FF000000DC000000B90000009600000073000000500000D4FFE300B1FFC7008EFFAB006BFF8F0048FF730025FF570000FF550000DC490000B93D00009631000073250000501900D4FFF000B1FFE2008EFFD4006BFFC60048FFB80025FFAA0000FFAA0000DC920000B97A000096620000734A0000503200D4FFFF00B1FFFF008EFFFF006BFFFF0048FFFF0025FFFF0000FFFF0000DCDC0000B9B900009696000073730000505000F2F2F200E6E6E600DADADA00CECECE00C2C2C200B6B6B600AAAAAA009E9E9E0092929200868686007A7A7A006E6E6E0062626200565656004A4A4A003E3E3E0032323200262626001A1A1A000E0E0E0011111111111111111111111111111111111111111111111111111111111111111111747474747474747474747474741111117411111174989898748C8C8C741111117411111174989898748C8C8C741111117411111174989898748C8C8C741111117474747474747474747474747411111174989898748C8C8C748282827411111174989898748C8C8C748282827411111174989898748C8C8C748282827411111174747474747474747474747474111111748C8C8C748282827482828274111111748C8C8C748282827482828274111111748C8C8C7482828274828282741111117474747474747474747474807411111111111111111111111111111111110000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF);
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (163,45,'New', 10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (164,45,'Edit', 20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (165,45,'View', 30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (166,45,'Delete', 40, 'DELETE');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (167,45,'Run', 40, 'RUN');		
	END
	UPDATE dbo.[ASRSysPermissionCategories] SET picture = 0x000001000100101000000000000068050000160000002800000010000000200000000100080000000000400100000000000000000000000100000000000000000000800080008000000080800000008000000080800000008000C0C0C000C0DCC000F0CAA60080808000FF00FF00FF000000FFFF000000FF000000FFFF000000FF00FFFFFF00F0FBFF00A4A0A000D4F0FF00B1E2FF008ED4FF006BC6FF0048B8FF0025AAFF0000AAFF000092DC00007AB90000629600004A730000325000D4E3FF00B1C7FF008EABFF006B8FFF004873FF002557FF000055FF000049DC00003DB900003196000025730000195000D4D4FF00B1B1FF008E8EFF006B6BFF004848FF002525FF000000FF000000DC000000B900000096000000730000005000E3D4FF00C7B1FF00AB8EFF008F6BFF007348FF005725FF005500FF004900DC003D00B900310096002500730019005000F0D4FF00E2B1FF00D48EFF00C66BFF00B848FF00AA25FF00AA00FF009200DC007A00B900620096004A00730032005000FFD4FF00FFB1FF00FF8EFF00FF6BFF00FF48FF00FF25FF00FF00FF00DC00DC00B900B900960096007300730050005000FFD4F000FFB1E200FF8ED400FF6BC600FF48B800FF25AA00FF00AA00DC009200B9007A009600620073004A0050003200FFD4E300FFB1C700FF8EAB00FF6B8F00FF487300FF255700FF005500DC004900B9003D00960031007300250050001900FFD4D400FFB1B100FF8E8E00FF6B6B00FF484800FF252500FF000000DC000000B9000000960000007300000050000000FFE3D400FFC7B100FFAB8E00FF8F6B00FF734800FF572500FF550000DC490000B93D0000963100007325000050190000FFF0D400FFE2B100FFD48E00FFC66B00FFB84800FFAA2500FFAA0000DC920000B97A000096620000734A000050320000FFFFD400FFFFB100FFFF8E00FFFF6B00FFFF4800FFFF2500FFFF0000DCDC0000B9B90000969600007373000050500000F0FFD400E2FFB100D4FF8E00C6FF6B00B8FF4800AAFF2500AAFF000092DC00007AB90000629600004A73000032500000E3FFD400C7FFB100ABFF8E008FFF6B0073FF480057FF250055FF000049DC00003DB90000319600002573000019500000D4FFD400B1FFB1008EFF8E006BFF6B0048FF480025FF250000FF000000DC000000B90000009600000073000000500000D4FFE300B1FFC7008EFFAB006BFF8F0048FF730025FF570000FF550000DC490000B93D00009631000073250000501900D4FFF000B1FFE2008EFFD4006BFFC60048FFB80025FFAA0000FFAA0000DC920000B97A000096620000734A0000503200D4FFFF00B1FFFF008EFFFF006BFFFF0048FFFF0025FFFF0000FFFF0000DCDC0000B9B900009696000073730000505000F2F2F200E6E6E600DADADA00CECECE00C2C2C200B6B6B600AAAAAA009E9E9E0092929200868686007A7A7A006E6E6E0062626200565656004A4A4A003E3E3E0032323200262626001A1A1A000E0E0E0011111111111111111111111111111111111111111111111111111111111111111111747474747474747474747474741111117411111174989898748C8C8C741111117411111174989898748C8C8C741111117411111174989898748C8C8C741111117474747474747474747474747411111174989898748C8C8C748282827411111174989898748C8C8C748282827411111174989898748C8C8C748282827411111174747474747474747474747474111111748C8C8C748282827482828274111111748C8C8C748282827482828274111111748C8C8C7482828274828282741111117474747474747474747474807411111111111111111111111111111111110000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF WHERE [categoryID] = 45


/* --------------------------------------------------------- */
PRINT 'Step - Licence Modifications'
/* --------------------------------------------------------- */

	EXEC spsys_setsystemsetting 'taxyear', 'startday', '06-Apr';

	IF OBJECT_ID('ASRSysWarningsLog', N'U') IS NULL	
	BEGIN
		EXECUTE sp_executeSQL N'CREATE TABLE [dbo].[ASRSysWarningsLog](
				[UserName]		varchar(255) NOT NULL,
				[WarningType]	integer  NOT NULL,
				[WarningDate]	datetime  NOT NULL) ON [PRIMARY]';
	END


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRUpdateWarningLog]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRUpdateWarningLog];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.spASRUpdateWarningLog(
		@Username			varchar(255),
		@WarningType		integer,
		@WarningRefreshRate	integer,
		@WarnUser			bit OUTPUT)
	AS
	BEGIN

		DECLARE @Today				datetime = GETDATE(),
				@LastWarningDate	datetime;

		SELECT TOP 1 @LastWarningDate = DATEADD(dd, 0, DATEDIFF(dd, 0, WarningDate)) FROM ASRSysWarningsLog
			WHERE Username = @Username AND WarningType = @WarningType
			ORDER BY WarningDate DESC;

		SET @WarnUser = 0;
		IF @LastWarningDate IS NULL OR DATEDIFF(day, @LastWarningDate, DATEDIFF(dd, 0, @Today)) >= @WarningRefreshRate SET @WarnUser = 1

		IF @WarnUser = 1
			INSERT ASRSysWarningsLog (UserName, WarningType, WarningDate) VALUES (@UserName, @WarningType, @Today);

		RETURN @WarnUser;
	END'


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