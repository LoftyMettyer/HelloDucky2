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


/* ------------------------------------------------------- */
PRINT 'Step - Web Messaging'
/* ------------------------------------------------------- */

-- Increase size of audit access 
ALTER TABLE [ASRSysAuditAccess] ALTER COLUMN [HRProModule] varchar(20);


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysCurrentLogins]') AND xtype in (N'U'))
	DROP TABLE [dbo].[ASRSysCurrentLogins];

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysCurrentSessions]') AND xtype in (N'U'))
	DROP TABLE [dbo].[ASRSysCurrentSessions];

SELECT @NVarCommand = 'CREATE TABLE ASRSysCurrentSessions(
	[IISServer]		nvarchar(255),
	[Username]		nvarchar(128),
	[Hostname]		nvarchar(255),
	[SessionID]		nvarchar(255),
	[loginTime]		datetime,
	[WebArea]	varchar(255))';
EXEC sp_executesql @NVarCommand


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

	-- Message to the Web Server
	DELETE FROM ASRSysMessages WHERE loginname = ''OpenHR Web Server'';

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

		-- Clone existing security based on cross tab permissions
		DELETE FROM ASRSysGroupPermissions WHERE itemid IN (163, 164, 165,166, 167)
		INSERT ASRSysGroupPermissions (itemID, groupName, permitted)
			SELECT 163, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 10
			UNION
			SELECT 164, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 11
			UNION
			SELECT 165, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 62
			UNION
			SELECT 166, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 12
			UNION
			SELECT 167, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid = 13

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
	END';


	-- Removal of DMIS licence option
	EXECUTE sp_executeSQL N'UPDATE ASRSysPermissionItems SET [description] = ''Data Manager Intranet'' WHERE categoryID = 1 AND itemKey = ''INTRANET'''
	EXECUTE sp_executeSQL N'UPDATE ASRSysPermissionItems SET [description] = ''Self-service'' WHERE categoryID = 1 AND itemKey = ''SSINTRANET'''
	EXECUTE sp_executeSQL N'DELETE FROM ASRSysPermissionItems where categoryID = 1 AND itemKey = ''INTRANET_SELFSERVICE'''
	EXECUTE sp_executeSQL N'DELETE FROM ASRSysGroupPermissions WHERE itemid = 100'

	-- Add view current users (DMI) security option
	IF NOT EXISTS(SELECT * FROM dbo.ASRSysPermissionItems WHERE [itemID] = 168)
	BEGIN
		INSERT ASRSysPermissionItems ([itemID], [description], [listOrder], [categoryID], [itemKey])
			VALUES (168,'View Current Users',20, 19,'CURRENTUSERS');

		INSERT ASRSysGroupPermissions (itemID, groupName, permitted)
			SELECT 168, groupName, permitted from ASRSysGroupPermissions where itemid = 78
	END


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASREnableServiceBroker]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASREnableServiceBroker];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASREnableServiceBroker]
	AS
	BEGIN
		DECLARE @sSQL nvarchar(MAX),
			@dbName	nvarchar(255) = DB_NAME(),
			@isBrokerEnabled bit = 0,
			@thisBrokerID uniqueidentifier,
			@uniqueBrokerCount integer;

		-- Is service broker enabled on this database?
		SELECT @isBrokerEnabled = is_broker_enabled, @thisBrokerID = service_broker_guid
			FROM sys.databases
			WHERE name = @dbName;

		-- Is it unique?
		SELECT @uniqueBrokerCount = COUNT(*)
			FROM sys.databases
			WHERE service_broker_guid = @thisBrokerID
			GROUP BY service_broker_guid;

		-- Enable if required
		IF @isBrokerEnabled = 0  OR (@isBrokerEnabled = 1 AND @uniqueBrokerCount > 1)
		BEGIN
			SET @sSQL = ''ALTER DATABASE ['' + @dbName + ''] SET NEW_BROKER WITH ROLLBACK IMMEDIATE;'';
			EXEC sp_executeSQL @sSQL;
		END

	END';


/* --------------------------------------------------------- */
PRINT 'Step - Workflow Log Enhancements'
/* --------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowInstances', 'U') AND name = 'TargetName')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowInstances ADD TargetName nvarchar(255) NULL;';
		EXEC sp_executesql N'UPDATE ASRSysWorkflowInstances SET TargetName = ''<Unidentified>'';';
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_Workflows', 'U') AND name = 'HasTargetIdentifier')
	BEGIN
		EXEC sp_executesql N'DROP VIEW ASRSysWorkflows;';
		EXEC sp_executesql N'ALTER TABLE tbsys_Workflows ADD HasTargetIdentifier bit;';
		EXEC sp_executesql N'CREATE VIEW [dbo].[ASRSysWorkflows]
					WITH SCHEMABINDING
					AS SELECT base.[id], base.[name], base.[description], base.[enabled], base.[initiationtype], base.[basetable], base.[querystring], base.[pictureid], obj.[locked], obj.[lastupdated], obj.[lastupdatedby], base.[HasTargetIdentifier]
						FROM dbo.[tbsys_workflows] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.id AND obj.objecttype = 10
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]';

		EXEC sp_executesql N'CREATE TRIGGER [dbo].[INS_ASRSysWorkflows] ON [dbo].[ASRSysWorkflows]
		INSTEAD OF INSERT
		AS
		BEGIN
	
			SET NOCOUNT ON;
	
			-- Update objects table
			IF NOT EXISTS(SELECT [guid]
				FROM dbo.[tbsys_scriptedobjects] o
				INNER JOIN inserted i ON i.id = o.targetid AND o.objecttype = 10)
			BEGIN
				INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
					SELECT NEWID(), 10, [id], dbo.[udfsys_getownerid](), ''01/01/1900'',1,0, GETDATE()
						FROM inserted;
			END

			-- Update base table								
			INSERT dbo.[tbsys_workflows] ([id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid], [HasTargetIdentifier]) 
				SELECT [id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid], [HasTargetIdentifier] FROM inserted;

		END';

		EXEC sp_executesql N'CREATE TRIGGER [dbo].[DEL_ASRSysWorkflows] ON [dbo].[ASRSysWorkflows]
		INSTEAD OF DELETE
		AS
		BEGIN
			SET NOCOUNT ON;

			DELETE FROM [tbsys_workflows] WHERE id IN (SELECT id FROM deleted);
			DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT id FROM deleted) AND objecttype = 10;

		END';

		EXEC sp_executesql N'UPDATE [ASRSysWorkflows] SET HasTargetIdentifier = 0'; 

		EXEC sp_executesql N'GRANT SELECT, UPDATE, INSERT, DELETE ON ASRSysWorkflows TO ASRSysGroup;';

	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElementItems', 'U') AND name = 'UseAsTargetIdentifier')
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElementItems ADD UseAsTargetIdentifier bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElements', 'U') AND name = 'UseAsTargetIdentifier')
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElements ADD UseAsTargetIdentifier bit;';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetStoredDataActionDetails];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
(
	@piInstanceID		integer,
	@piElementID		integer,
	@psSQL				varchar(MAX)	OUTPUT, 
	@piDataTableID		integer			OUTPUT,
	@psTableName		varchar(255)	OUTPUT,
	@piDataAction		integer			OUTPUT, 
	@piRecordID			integer			OUTPUT,
	@bUseAsTargetIdentifier	bit OUTPUT,
	@pfResult	bit OUTPUT
)
AS
BEGIN
	DECLARE 
		@iPersonnelTableID			integer,
		@iInitiatorID				integer,
		@iDataRecord				integer,
		@sIDColumnName				varchar(MAX),
		@iColumnID					integer, 
		@sColumnName				varchar(MAX), 
		@iColumnDataType			integer, 
		@sColumnList				varchar(MAX),
		@sValueList					varchar(MAX),
		@sValue						varchar(MAX),
		@sRecSelWebFormIdentifier	varchar(MAX),
		@sRecSelIdentifier			varchar(MAX),
		@iTempTableID				integer,
		@iSecondaryDataRecord		integer,
		@sSecondaryRecSelWebFormIdentifier	varchar(MAX),
		@sSecondaryRecSelIdentifier	varchar(MAX),
		@sSecondaryIDColumnName		varchar(MAX),
		@iSecondaryRecordID			integer,
		@iElementType				integer,
		@iWorkflowID				integer,
		@iID						integer,
		@sWFFormIdentifier			varchar(MAX),
		@sWFValueIdentifier			varchar(MAX),
		@iDBColumnID				integer,
		@iDBRecord					integer,
		@sSQL						nvarchar(MAX),
		@sParam						nvarchar(MAX),
		@sDBColumnName				nvarchar(MAX),
		@sDBTableName				nvarchar(MAX),
		@iRecordID					integer,
		@sDBValue					varchar(MAX),
		@iDataType					integer, 
		@iValueType					integer, 
		@iSDColumnID				integer,
		@fValidRecordID				bit,
		@iBaseTableID				integer,
		@iBaseRecordID				integer,
		@iRequiredTableID			integer,
		@iRequiredRecordID			integer,
		@iDataRecordTableID			integer,
		@iSecondaryDataRecordTableID	integer,
		@iParent1TableID			integer,
		@iParent1RecordID			integer,
		@iParent2TableID			integer,
		@iParent2RecordID			integer,
		@iInitParent1TableID		integer,
		@iInitParent1RecordID		integer,
		@iInitParent2TableID		integer,
		@iInitParent2RecordID		integer,
		@iEmailID					integer,
		@iType						integer,
		@fDeletedValue				bit,
		@iTempElementID				integer,
		@iCount						integer,
		@iResultType				integer,
		@sResult					varchar(MAX),
		@fResult					bit,
		@dtResult					datetime,
		@fltResult					float,
		@iCalcID					integer,
		@iSize						integer,
		@iDecimals					integer,
		@iTriggerTableID			integer;
			
	SET @psSQL = '''';
	SET @pfResult = 1;
	SET @piRecordID = 0;

	SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
	FROM ASRSysModuleSetup
	WHERE moduleKey = ''MODULE_PERSONNEL''
		AND parameterKey = ''Param_TablePersonnel'';

	IF @iPersonnelTableID = 0
	BEGIN
		SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
		FROM ASRSysModuleSetup
		WHERE moduleKey = ''MODULE_WORKFLOW''
		AND parameterKey = ''Param_TablePersonnel'';
	END

	SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
		@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
		@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
		@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
		@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
	FROM ASRSysWorkflowInstances
	WHERE ASRSysWorkflowInstances.ID = @piInstanceID;

	SELECT @piDataAction = dataAction,
		@piDataTableID = dataTableID,
		@iDataRecord = dataRecord,
		@sRecSelWebFormIdentifier = recSelWebFormIdentifier,
		@sRecSelIdentifier = recSelIdentifier,
		@iSecondaryDataRecord = secondaryDataRecord,
		@sSecondaryRecSelWebFormIdentifier = secondaryRecSelWebFormIdentifier,
		@sSecondaryRecSelIdentifier = secondaryRecSelIdentifier,
		@iDataRecordTableID = dataRecordTable,
		@iSecondaryDataRecordTableID = secondaryDataRecordTable,
		@iWorkflowID = workflowID,
		@iTriggerTableID = ASRSysWorkflows.baseTable,
		@bUseAsTargetIdentifier = ISNULL(UseAsTargetIdentifier, 0)
	FROM ASRSysWorkflowElements
	INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID
	WHERE ASRSysWorkflowElements.ID = @piElementID;

	SELECT @psTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @piDataTableID;

	IF @iDataRecord = 0 -- 0 = Initiator''s record
	BEGIN
		EXEC [dbo].[spASRWorkflowAscendantRecordID]
			@iPersonnelTableID,
			@iInitiatorID,
			@iInitParent1TableID,
			@iInitParent1RecordID,
			@iInitParent2TableID,
			@iInitParent2RecordID,
			@iDataRecordTableID,
			@piRecordID	OUTPUT;

		IF @piDataTableID = @iDataRecordTableID
		BEGIN
			SET @sIDColumnName = ''ID'';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
		END
	END

	IF @iDataRecord = 4 -- 4 = Triggered record
	BEGIN
		EXEC [dbo].[spASRWorkflowAscendantRecordID]
			@iTriggerTableID,
			@iInitiatorID,
			@iInitParent1TableID,
			@iInitParent1RecordID,
			@iInitParent2TableID,
			@iInitParent2RecordID,
			@iDataRecordTableID,
			@piRecordID	OUTPUT;

		IF @piDataTableID = @iDataRecordTableID
		BEGIN
			SET @sIDColumnName = ''ID'';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
		END
	END

	IF @iDataRecord = 1 -- 1 = Identified record
	BEGIN
		SELECT @iElementType = ASRSysWorkflowElements.type
		FROM ASRSysWorkflowElements
		WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
			AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
		IF @iElementType = 2
		BEGIN
			 -- WebForm
			SELECT @sValue = ISNULL(IV.value, ''0''),
				@iTempTableID = EI.tableID,
				@iParent1TableID = IV.parent1TableID,
				@iParent1RecordID = IV.parent1RecordID,
				@iParent2TableID = IV.parent2TableID,
				@iParent2RecordID = IV.parent2RecordID
			FROM ASRSysWorkflowInstanceValues IV
			INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
			INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
			WHERE IV.instanceID = @piInstanceID
				AND IV.identifier = @sRecSelIdentifier
				AND Es.identifier = @sRecSelWebFormIdentifier
				AND Es.workflowID = @iWorkflowID
				AND IV.elementID = Es.ID;
		END
		ELSE
		BEGIN
			-- StoredData
			SELECT @sValue = ISNULL(IV.value, ''0''),
				@iTempTableID = Es.dataTableID,
				@iParent1TableID = IV.parent1TableID,
				@iParent1RecordID = IV.parent1RecordID,
				@iParent2TableID = IV.parent2TableID,
				@iParent2RecordID = IV.parent2RecordID
			FROM ASRSysWorkflowInstanceValues IV
			INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
				AND IV.identifier = Es.identifier
				AND Es.workflowID = @iWorkflowID
				AND Es.identifier = @sRecSelWebFormIdentifier
			WHERE IV.instanceID = @piInstanceID;
		END

		SET @piRecordID = 
			CASE
				WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
				ELSE 0
			END;
	
		SET @iBaseTableID = @iTempTableID;
		SET @iBaseRecordID = @piRecordID;
		EXEC [dbo].[spASRWorkflowAscendantRecordID]
			@iBaseTableID,
			@iBaseRecordID,
			@iParent1TableID,
			@iParent1RecordID,
			@iParent2TableID,
			@iParent2RecordID,
			@iDataRecordTableID,
			@piRecordID	OUTPUT;

		IF @piDataTableID = @iDataRecordTableID
		BEGIN
			SET @sIDColumnName = ''ID'';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
		END
	END

	SET @fValidRecordID = 1
	IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
	BEGIN
		EXEC [dbo].[spASRWorkflowValidTableRecord]
			@iDataRecordTableID,
			@piRecordID,
			@fValidRecordID	OUTPUT;

		IF @fValidRecordID = 0
		BEGIN
			-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
			EXEC [dbo].[spASRWorkflowActionFailed]
				@piInstanceID, 
				@piElementID, 
				''Stored Data primary record has been deleted or not selected.'';

			SET @psSQL = '''';
			SET @pfResult = 0;
			RETURN;
		END
	END

	IF @piDataAction = 0 -- Insert
	BEGIN
		IF @iSecondaryDataRecord = 0 -- 0 = Initiator''s record
		BEGIN
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iPersonnelTableID,
				@iInitiatorID,
				@iInitParent1TableID,
				@iInitParent1RecordID,
				@iInitParent2TableID,
				@iInitParent2RecordID,
				@iSecondaryDataRecordTableID,
				@iSecondaryRecordID	OUTPUT;

			IF @piDataTableID = @iSecondaryDataRecordTableID
			BEGIN
				SET @sSecondaryIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
			END
		END
		
		IF @iSecondaryDataRecord = 4 -- 4 = Triggered record
		BEGIN
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iTriggerTableID,
				@iInitiatorID,
				@iInitParent1TableID,
				@iInitParent1RecordID,
				@iInitParent2TableID,
				@iInitParent2RecordID,
				@iSecondaryDataRecordTableID,
				@iSecondaryRecordID	OUTPUT;
	
			IF @piDataTableID = @iSecondaryDataRecordTableID
			BEGIN
				SET @sSecondaryIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
			END
		END

		IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
		BEGIN
			SELECT @iElementType = ASRSysWorkflowElements.type
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
				AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)));
	
			IF @iElementType = 2
			BEGIN
				 -- WebForm
				SELECT @sValue = ISNULL(IV.value, ''0''),
					@iTempTableID = EI.tableID,
					@iParent1TableID = IV.parent1TableID,
					@iParent1RecordID = IV.parent1RecordID,
					@iParent2TableID = IV.parent2TableID,
					@iParent2RecordID = IV.parent2RecordID
				FROM ASRSysWorkflowInstanceValues IV
				INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
				INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
				WHERE IV.instanceID = @piInstanceID
					AND IV.identifier = @sSecondaryRecSelIdentifier
					AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
					AND Es.workflowID = @iWorkflowID
					AND IV.elementID = Es.ID;
			END
			ELSE
			BEGIN
				-- StoredData
				SELECT @sValue = ISNULL(IV.value, ''0''),
					@iTempTableID = Es.dataTableID,
					@iParent1TableID = IV.parent1TableID,
					@iParent1RecordID = IV.parent1RecordID,
					@iParent2TableID = IV.parent2TableID,
					@iParent2RecordID = IV.parent2RecordID
				FROM ASRSysWorkflowInstanceValues IV
				INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
					AND IV.identifier = Es.identifier
					AND Es.workflowID = @iWorkflowID
					AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
				WHERE IV.instanceID = @piInstanceID;
			END

			SET @iSecondaryRecordID = 
				CASE
					WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
					ELSE 0
				END;
			
			SET @iBaseTableID = @iTempTableID;
			SET @iBaseRecordID = @iSecondaryRecordID;
			EXEC [dbo].[spASRWorkflowAscendantRecordID]
				@iBaseTableID,
				@iBaseRecordID,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				@iSecondaryDataRecordTableID,
				@iSecondaryRecordID	OUTPUT;

			IF @piDataTableID = @iSecondaryDataRecordTableID
			BEGIN
				SET @sSecondaryIDColumnName = ''ID'';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
			END
		END

		SET @fValidRecordID = 1;
		IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
		BEGIN
			EXEC [dbo].[spASRWorkflowValidTableRecord]
				@iSecondaryDataRecordTableID,
				@iSecondaryRecordID,
				@fValidRecordID	OUTPUT;

			IF @fValidRecordID = 0
			BEGIN
				-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
				EXEC [dbo].[spASRWorkflowActionFailed] 
					@piInstanceID, 
					@piElementID, 
					''Stored Data secondary record has been deleted or not selected.'';

				SET @psSQL = '''';
				SET @pfResult = 0;
				RETURN;
			END
		END

	END

	IF @piDataAction = 0 OR @piDataAction = 1
	BEGIN
		/* INSERT or UPDATE. */
		SET @sColumnList = '''';
		SET @sValueList = '''';

		DECLARE @dbValues TABLE (
			ID integer, 
			wfFormIdentifier varchar(1000),
			wfValueIdentifier varchar(1000),
			dbColumnID int,
			dbRecord int,
			value varchar(MAX));

		INSERT INTO @dbValues (ID, 
			wfFormIdentifier,
			wfValueIdentifier,
			dbColumnID,
			dbRecord,
			value) 
		SELECT EC.ID,
			EC.wfformidentifier,
			EC.wfvalueidentifier,
			EC.dbcolumnid,
			EC.dbrecord, 
			''''
		FROM ASRSysWorkflowElementColumns EC
		WHERE EC.elementID = @piElementID
			AND EC.valueType = 2;
			
		DECLARE dbValuesCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ID,
			wfFormIdentifier,
			wfValueIdentifier,
			dbColumnID,
			dbRecord
		FROM @dbValues;
		OPEN dbValuesCursor;
		FETCH NEXT FROM dbValuesCursor INTO @iID,
			@sWFFormIdentifier,
			@sWFValueIdentifier,
			@iDBColumnID,
			@iDBRecord;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @fDeletedValue = 0;

			SELECT @sDBTableName = tbl.tableName,
				@iRequiredTableID = tbl.tableID, 
				@sDBColumnName = col.columnName,
				@iDataType = col.dataType
			FROM ASRSysColumns col
			INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
			WHERE col.columnID = @iDBColumnID;

			SET @sSQL = ''SELECT @sDBValue = ''
				+ CASE
					WHEN @iDataType = 12 THEN ''''
					WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
					ELSE ''convert(varchar(MAX),''
				END
				+ @sDBTableName + ''.'' + @sDBColumnName
				+ CASE
					WHEN @iDataType = 12 THEN ''''
					WHEN @iDataType = 11 THEN '', 101)''
					ELSE '')''
				END
				+ '' FROM '' + @sDBTableName 
				+ '' WHERE '' + @sDBTableName + ''.ID = '';

			SET @iRecordID = 0;

			IF @iDBRecord = 0
			BEGIN
				-- Initiator''s record
				SET @iRecordID = @iInitiatorID;
				SET @iParent1TableID = @iInitParent1TableID;
				SET @iParent1RecordID = @iInitParent1RecordID;
				SET @iParent2TableID = @iInitParent2TableID;
				SET @iParent2RecordID = @iInitParent2RecordID;
				SET @iBaseTableID = @iPersonnelTableID;
			END			

			IF @iDBRecord = 4
			BEGIN
				-- Trigger record
				SET @iRecordID = @iInitiatorID;
				SET @iParent1TableID = @iInitParent1TableID;
				SET @iParent1RecordID = @iInitParent1RecordID;
				SET @iParent2TableID = @iInitParent2TableID;
				SET @iParent2RecordID = @iInitParent2RecordID;

				SELECT @iBaseTableID = isnull(WF.baseTable, 0)
				FROM ASRSysWorkflows WF
				INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
					AND WFI.ID = @piInstanceID;
			END
			
			IF @iDBRecord = 1
			BEGIN
				-- Identified record
				SELECT @iElementType = ASRSysWorkflowElements.type, 
					@iTempElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)));

				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iBaseTableID = EI.tableID,
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
					INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
					WHERE IV.instanceID = @piInstanceID
						AND IV.identifier = @sWFValueIdentifier
						AND Es.identifier = @sWFFormIdentifier
						AND Es.workflowID = @iWorkflowID
						AND IV.elementID = Es.ID;
				END
				ELSE
				BEGIN
					-- StoredData
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iBaseTableID = isnull(Es.dataTableID, 0),
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
						AND IV.identifier = Es.identifier
						AND Es.workflowID = @iWorkflowID
						AND Es.identifier = @sWFFormIdentifier
					WHERE IV.instanceID = @piInstanceID;
				END

				SET @iRecordID = 
					CASE
						WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
						ELSE 0
					END;
			END

			SET @iBaseRecordID = @iRecordID;

			SET @fValidRecordID = 1;
			
			IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
			BEGIN
				SET @fValidRecordID = 0;

				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iBaseTableID,
					@iBaseRecordID,
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					@iRequiredTableID,
					@iRequiredRecordID	OUTPUT;

				SET @iRecordID = @iRequiredRecordID;

				IF @iRecordID > 0 
				BEGIN
					EXEC [dbo].[spASRWorkflowValidTableRecord]
						@iRequiredTableID,
						@iRecordID,
						@fValidRecordID	OUTPUT;
				END

				IF @fValidRecordID = 0
				BEGIN
					IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
					BEGIN
						SELECT @iCount = COUNT(*)
						FROM ASRSysWorkflowQueueColumns QC
						INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
						WHERE WFQ.instanceID = @piInstanceID
							AND QC.columnID = @iDBColumnID;

						IF @iCount = 1
						BEGIN
							SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
							FROM ASRSysWorkflowQueueColumns QC
							INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
							WHERE WFQ.instanceID = @piInstanceID
								AND QC.columnID = @iDBColumnID;

							SET @fValidRecordID = 1;
							SET @fDeletedValue = 1;
						END
					END
					ELSE
					BEGIN
						IF @iDBRecord = 1
						BEGIN
							SELECT @iCount = COUNT(*)
							FROM ASRSysWorkflowInstanceValues IV
							WHERE IV.instanceID = @piInstanceID
								AND IV.columnID = @iDBColumnID
								AND IV.elementID = @iTempElementID;

							IF @iCount = 1
							BEGIN
								SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '''')))
								FROM ASRSysWorkflowInstanceValues IV
								WHERE IV.instanceID = @piInstanceID
									AND IV.columnID = @iDBColumnID
									AND IV.elementID = @iTempElementID;

								SET @fValidRecordID = 1;
								SET @fDeletedValue = 1;
							END
						END
					END
				END

				IF @fValidRecordID = 0
				BEGIN
					-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
					EXEC [dbo].[spASRWorkflowActionFailed]
						@piInstanceID, 
						@piElementID, 
						''Stored Data column database value record has been deleted or not selected.'';

					SET @psSQL = '''';
					SET @pfResult = 0;
					RETURN;
				END
			END

			IF (@iDataType <> -3)
				AND (@iDataType <> -4)
			BEGIN
				IF @fDeletedValue = 0
				BEGIN
					SET @sSQL = @sSQL + convert(nvarchar(255), @iRecordID);
					SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
					EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
				END

				UPDATE @dbValues
				SET value = @sDBValue
				WHERE ID = @iID;
			END
			
			FETCH NEXT FROM dbValuesCursor INTO @iID,
				@sWFFormIdentifier,
				@sWFValueIdentifier,
				@iDBColumnID,
				@iDBRecord;
		END
		CLOSE dbValuesCursor;
		DEALLOCATE dbValuesCursor;

		DECLARE columnCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT EC.columnID,
			SC.columnName,
			SC.dataType,
			CASE
				WHEN EC.valueType = 0 THEN  -- Fixed Value
					CASE
						WHEN SC.dataType = -7 THEN
							CASE 
								WHEN UPPER(EC.value) = ''TRUE'' THEN ''1''
								ELSE ''0''
							END
						ELSE EC.value
					END
				WHEN EC.valueType = 1 THEN -- Workflow Value
					(SELECT IV.value
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
					INNER JOIN ASRSysWorkflowElements WE2 ON WE.workflowID = WE2.workflowID
					WHERE WE.identifier = EC.WFFormIdentifier
						AND WE2.id = @piElementID
						AND IV.instanceID = @piInstanceID
						AND IV.identifier = EC.WFValueIdentifier)
				ELSE '''' -- Database Value. Handle below to avoid collation conflict.
				END AS [value], 
				EC.valueType, 
				EC.ID,
				EC.calcID,
				isnull(SC.size, 0),
				isnull(SC.decimals, 0)
		FROM ASRSysWorkflowElementColumns EC
		INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
		WHERE EC.elementID = @piElementID
			AND ((SC.dataType <> -3) AND (SC.dataType <> -4));

		OPEN columnCursor;
		FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
			BEGIN
				SELECT @sValue = dbV.value
				FROM @dbValues dbV
				WHERE dbV.ID = @iSDColumnID;
			END

			IF @iValueType = 3 -- Calculated Value
			BEGIN
				EXEC [dbo].[spASRSysWorkflowCalculation]
					@piInstanceID,
					@iCalcID,
					@iResultType OUTPUT,
					@sResult OUTPUT,
					@fResult OUTPUT,
					@dtResult OUTPUT,
					@fltResult OUTPUT, 
					0;

				IF @iColumnDataType = 12 SET @sResult = LEFT(@sResult, @iSize); -- Character
				IF @iColumnDataType = 2 -- Numeric
				BEGIN
					IF @fltResult >= power(10, @iSize - @iDecimals) SET @fltResult = 0;
					IF @fltResult <= (-1 * power(10, @iSize - @iDecimals)) SET @fltResult = 0;
				END

				SET @sValue = 
					CASE
						WHEN @iResultType = 2 THEN ltrim(rtrim(STR(@fltResult, 8000, @iDecimals)))
						WHEN @iResultType = 3 THEN 
							CASE 
								WHEN @fResult = 1 THEN ''1''
								ELSE ''0''
							END
						WHEN (@iResultType = 4) THEN
							CASE 
								WHEN @dtResult is NULL THEN ''NULL''
								ELSE convert(varchar(100), @dtResult, 101)
							END
						ELSE convert(varchar(MAX), @sResult)
					END;
			END

			IF @piDataAction = 0 
			BEGIN
				/* INSERT. */
				SET @sColumnList = @sColumnList
					+ CASE
						WHEN LEN(@sColumnList) > 0 THEN '',''
						ELSE ''''
					END
					+ @sColumnName;

				SET @sValueList = @sValueList
					+ CASE
						WHEN LEN(@sValueList) > 0 THEN '',''
						ELSE ''''
					END
					+ CASE
						WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
						WHEN @iColumnDataType = 11 THEN
							CASE 
								WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
								ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
							END
						WHEN LEN(@sValue) = 0 THEN ''0''
						ELSE isnull(@sValue, 0) -- integer, logic, numeric
					END;
			END
			ELSE
			BEGIN
				/* UPDATE. */
				SET @sColumnList = @sColumnList
					+ CASE
						WHEN LEN(@sColumnList) > 0 THEN '',''
						ELSE ''''
					END
					+ @sColumnName
					+ '' = ''
					+ CASE
						WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
						WHEN @iColumnDataType = 11 THEN
							CASE 
								WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
								ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
							END
						WHEN LEN(@sValue) = 0 THEN ''0''
						ELSE isnull(@sValue, 0) -- integer, logic, numeric
					END;
			END

			DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
			WHERE instanceID = @piInstanceID
				AND elementID = @piElementID
				AND columnID = @iColumnID;

			INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
				(instanceID, elementID, identifier, columnID, value, emailID)
				VALUES (@piInstanceID, @piElementID, '''', @iColumnID, @sValue, 0);

			FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
		END

		CLOSE columnCursor;
		DEALLOCATE columnCursor;

		IF @piDataAction = 0 
		BEGIN
			/* INSERT. */
			IF @iDataRecord <> 3 -- 3 = Unidentified record
			BEGIN
				SET @sColumnList = @sColumnList
					+ CASE
						WHEN LEN(@sColumnList) > 0 THEN '',''
						ELSE ''''
					END
					+ @sIDColumnName;
	
				SET @sValueList = @sValueList
					+ CASE
						WHEN LEN(@sValueList) > 0 THEN '',''
						ELSE ''''
					END
					+ convert(varchar(255), @piRecordID);

				IF @piDataAction = 0 -- Insert
					AND (@iSecondaryDataRecord = 0 -- 0 = Initiator''s record
						OR @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
						OR @iSecondaryDataRecord = 4) -- 4 = Triggered record
				BEGIN
					SET @sColumnList = @sColumnList
						+ CASE
							WHEN LEN(@sColumnList) > 0 THEN '',''
							ELSE ''''
						END
						+ @sSecondaryIDColumnName;
				
					SET @sValueList = @sValueList
						+ CASE
							WHEN LEN(@sValueList) > 0 THEN '',''
							ELSE ''''
						END
						+ convert(varchar(255), @iSecondaryRecordID);
				END
			END
		END

		IF LEN(@sColumnList) > 0
		BEGIN
			IF @piDataAction = 0 
			BEGIN
				/* INSERT. */
				SET @psSQL = ''INSERT INTO '' + @psTableName
					+ '' ('' + @sColumnList + '')''
					+ '' VALUES('' + @sValueList + '')'';
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				/* UPDATE. */
				SET @psSQL = ''UPDATE '' + @psTableName
					+ '' SET '' + @sColumnList
					+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
				SET @pfResult = 1;
			END
		END
	END

	IF @piDataAction = 2
	BEGIN
		/* DELETE. */
		SET @psSQL = ''DELETE FROM '' + @psTableName
			+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
		SET @pfResult = 1;
	END	

	IF (@piDataAction = 0) -- Insert
	BEGIN
		SET @iParent1TableID = isnull(@iDataRecordTableID, 0);
		SET @iParent1RecordID = isnull(@piRecordID, 0);
		SET @iParent2TableID = isnull(@iSecondaryDataRecordTableID, 0);
		SET @iParent2RecordID = isnull(@iSecondaryRecordID, 0);
	END
	ELSE
	BEGIN	-- Update or Delete
		exec [dbo].[spASRGetParentDetails]
			@piDataTableID,
			@piRecordID,
			@iParent1TableID	OUTPUT,
			@iParent1RecordID	OUTPUT,
			@iParent2TableID	OUTPUT,
			@iParent2RecordID	OUTPUT;
	END

	UPDATE ASRSysWorkflowInstanceValues
	SET ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID, 
		ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
		ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID, 
		ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
	WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceValues.elementID = @piElementID
		AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
		AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;

	IF (@piDataAction = 2) -- Delete
	BEGIN
		DECLARE curColumns CURSOR LOCAL FAST_FORWARD FOR 
		SELECT columnID
		FROM [dbo].[udfASRWorkflowColumnsUsed] (@iWorkflowID, @piElementID, 0);

		OPEN curColumns;

		FETCH NEXT FROM curColumns INTO @iDBColumnID;
		WHILE (@@fetch_status = 0)
		BEGIN
			DELETE FROM ASRSysWorkflowInstanceValues
			WHERE instanceID = @piInstanceID
				AND elementID = @piElementID
				AND columnID = @iDBColumnID;

			SELECT @sDBTableName = tbl.tableName,
				@iRequiredTableID = tbl.tableID, 
				@sDBColumnName = col.columnName,
				@iDataType = col.dataType
			FROM ASRSysColumns col
			INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
			WHERE col.columnID = @iDBColumnID;

			SET @sSQL = ''SELECT @sDBValue = ''
				+ CASE
					WHEN @iDataType = 12 THEN ''''
					WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
					ELSE ''convert(varchar(MAX),''
				END
				+ @sDBTableName + ''.'' + @sDBColumnName
				+ CASE
					WHEN @iDataType = 12 THEN ''''
					WHEN @iDataType = 11 THEN '', 101)''
					ELSE '')''
				END
				+ '' FROM '' + @sDBTableName 
				+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);

			SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
			EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;

			INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
				(instanceID, elementID, identifier, columnID, value, emailID)
				VALUES (@piInstanceID, @piElementID, '''', @iDBColumnID, @sDBValue, 0);
					
			FETCH NEXT FROM curColumns INTO @iDBColumnID;
		END
		CLOSE curColumns;
		DEALLOCATE curColumns;

		DECLARE curEmails CURSOR LOCAL FAST_FORWARD FOR 
		SELECT emailID,
			type,
			colExprID
		FROM [dbo].[udfASRWorkflowEmailsUsed] (@iWorkflowID, @piElementID, 0);

		OPEN curEmails;

		FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
		WHILE (@@fetch_status = 0)
		BEGIN
			DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
			WHERE instanceID = @piInstanceID
				AND elementID = @piElementID
				AND emailID = @iEmailID;

			IF @iType = 1 -- Column
			BEGIN
				SELECT @sDBTableName = tbl.tableName,
					@iRequiredTableID = tbl.tableID, 
					@sDBColumnName = col.columnName,
					@iDataType = col.dataType
				FROM [dbo].[ASRSysColumns] col
				INNER JOIN [dbo].[ASRSysTables] tbl ON col.tableID = tbl.tableID
				WHERE col.columnID = @iDBColumnID;

				SET @sSQL = ''SELECT @sDBValue = ''
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
						ELSE ''convert(varchar(MAX),''
					END
					+ @sDBTableName + ''.'' + @sDBColumnName
					+ CASE
						WHEN @iDataType = 12 THEN ''''
						WHEN @iDataType = 11 THEN '', 101)''
						ELSE '')''
					END
					+ '' FROM '' + @sDBTableName 
					+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);

				SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
				EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
			END
			ELSE
			BEGIN
				EXEC [dbo].[spASRSysEmailAddr]
					@sDBValue OUTPUT,
					@iEmailID,
					@piRecordID;
			END

			INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
				(instanceID, elementID, identifier, columnID, value, emailID)
				VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID);
					
			FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
		END
		CLOSE curEmails;
		DEALLOCATE curEmails;
	END
END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(MAX)	OUTPUT,
			@psMessage		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID			integer,
				@iStepID				integer,
				@iElementID				integer,
				@iRecordID				integer,
				@iRecordCount			integer,
				@sTargetName			nvarchar(MAX) = '''',
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@bUseAsTargetIdentifier bit,
				@iTemp					integer,
				@iStartElementID		integer,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms					varchar(MAX),
				@iCount					integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled				bit,
				@fHasTargetIdentifier bit,
				@iElementType			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(255),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sSPName				varchar(255),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fSaveForLater			bit,
				@fResult	bit;
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
		
			DECLARE @succeedingElements table(elementID int);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT @fExternallyInitiated = CASE
					WHEN initiationType = 2 THEN 1
					ELSE 0
				END,
				@fEnabled = [enabled],
				@fHasTargetIdentifier = [HasTargetIdentifier]
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;
		
			IF @fExternallyInitiated = 1
			BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = ''<External>'';
			END
			ELSE
			BEGIN
				SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT,
						@sTargetName OUTPUT;
				END

				IF @fHasTargetIdentifier = 1
					SET @sTargetName = ''<Unidentified>'';
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC [dbo].[spASRWorkflowUsesInitiator]
						@piWorkflowID,
						@fUsesInitiator OUTPUT;
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.'';
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.'';
						END
			
						RETURN
					END	
				END
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel'';
					END
		
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
			END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[TargetName],
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				[pageno])
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@sTargetName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = MAX(id)
			FROM [dbo].[ASRSysWorkflowInstances];
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
		
			INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0);
		
			INSERT INTO [dbo].[ASRSysWorkflowInstanceSteps] (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 1
					ELSE 0
				END,
				0,
				0
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID;
		
			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5;
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
					
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
				OPEN immediateSubmitCursor;
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''spASRWorkflowInsertNewRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL;
		
								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''spASRWorkflowUpdateRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC [dbo].[spASRRecordDescription]
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''spASRWorkflowDeleteRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							EXEC [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT;
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(MAX), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 4,	-- 4 = failed
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
		
								SET @psMessage = @sStoredDataMsg;
								RETURN;
							END
							ELSE
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 8,	-- 8 = failed action
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE [instanceID] = @piInstanceID
									AND [elementID] = @iElementID;
		
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
									SET ASRSysWorkflowInstanceSteps.status = 1
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT,
							0;
					END
		
					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				END
				CLOSE immediateSubmitCursor;
				DEALLOCATE immediateSubmitCursor;
		
				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM [dbo].[ASRSysWorkflowInstanceSteps]
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;
			END						
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE @succeedingSteps table(stepID int)
			
			INSERT INTO @succeedingSteps 
				(stepID) VALUES (-1)
		
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM [dbo].[ASRSysWorkflowInstanceSteps]
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
			OPEN formsCursor;
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
				INSERT INTO @succeedingSteps 
				(stepID) VALUES (@iStepID)
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			END
		
			CLOSE formsCursor;
			DEALLOCATE formsCursor;
		
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET ASRSysWorkflowInstanceSteps.status = 2, 
				userName = @sActualLoginName
			WHERE ASRSysWorkflowInstanceSteps.ID IN (SELECT stepID FROM @succeedingSteps)
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
	(
		@piInstanceID		integer,
		@piElementID		integer,
		@psFormInput1		varchar(MAX),
		@psFormElements		varchar(MAX)	OUTPUT,
		@pfSavedForLater	bit				OUTPUT,
		@piPageNo	integer
	)
	AS
	BEGIN
		DECLARE
			@iIndex1			integer,
			@iIndex2			integer,
			@iID				integer,
			@sID				varchar(MAX),
			@sValue				varchar(MAX),
			@iElementType		integer,
			@iPreviousElementID	integer,
			@iValue				integer,
			@hResult			integer,
			@hTmpResult			integer,
			@sTo				varchar(MAX),
			@sCopyTo			varchar(MAX),
			@sTempTo			varchar(MAX),
			@sMessage			varchar(MAX),
			@sMessage_HypertextLinks	varchar(MAX),
			@sHypertextLinkedSteps		varchar(MAX),
			@iEmailID			integer,
			@iEmailCopyID		integer,
			@iTempEmailID		integer,
			@iEmailLoop			integer,
			@iEmailRecord		integer,
			@iEmailRecordID		integer,
			@sSQL				nvarchar(MAX),
			@iCount				integer,
			@superCursor		cursor,
			@curDelegatedRecords	cursor,
			@fDelegate			bit,
			@fDelegationValid	bit,
			@iDelegateEmailID	integer,
			@iDelegateRecordID	integer,
			@sTemp				varchar(MAX),
			@sDelegateTo		varchar(MAX),
			@sAllDelegateTo		varchar(MAX),
			@iCurrentStepID		int,
			@sDelegatedMessage	varchar(MAX),
			@iTemp				integer, 
			@iPrevElementType	integer,
			@iWorkflowID		integer,
			@sRecSelIdentifier	varchar(MAX),
			@sRecSelWebFormIdentifier	varchar(MAX), 
			@iStepID			int,
			@iElementID			int,
			@sUserName			varchar(MAX),
			@sUserEmail			varchar(MAX), 
			@sValueDescription	varchar(MAX),
			@iTableID			integer,
			@iRecDescID			integer,
			@bUseAsTargetIdentifier bit = 0,
			@sEvalRecDesc		varchar(MAX),
			@sExecString		nvarchar(MAX),
			@sParamDefinition	nvarchar(500),
			@sIdentifier		varchar(MAX),
			@iItemType			integer,
			@iDataAction		integer, 
			@fValidRecordID		bit,
			@iEmailTableID		integer,
			@iEmailType			integer,
			@iBaseTableID		integer,
			@iBaseRecordID		integer,
			@iRequiredRecordID	integer,
			@iParent1TableID	int,
			@iParent1RecordID	int,
			@iParent2TableID	int,
			@iParent2RecordID	int,
			@iTempElementID		integer,
			@iTrueFlowType		integer,
			@iExprID			integer,
			@iResultType		integer,
			@sResult			varchar(MAX),
			@fResult			bit,
			@dtResult			datetime,
			@fltResult			float,
			@sEmailSubject		varchar(200),
			@iTempID			integer,
			@iBehaviour			integer;

		SET @pfSavedForLater = 0;

		SELECT @iCurrentStepID = ID
		FROM ASRSysWorkflowInstanceSteps
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;

		SET @iDelegateEmailID = 0;
		SELECT @sTemp = ISNULL(parameterValue, '''')
		FROM ASRSysModuleSetup
		WHERE moduleKey = ''MODULE_WORKFLOW''
			AND parameterKey = ''Param_DelegateEmail'';
		SET @iDelegateEmailID = convert(integer, @sTemp);

		SET @psFormElements = '''';
				
		-- Get the type of the given element 
		SELECT @iElementType = E.type,
			@iEmailID = E.emailID,
			@iEmailCopyID = isnull(E.emailCCID, 0),
			@iEmailRecord = E.emailRecord, 
			@iWorkflowID = E.workflowID,
			@sRecSelIdentifier = E.RecSelIdentifier, 
			@sRecSelWebFormIdentifier = E.RecSelWebFormIdentifier, 
			@iTableID = E.dataTableID,
			@iDataAction = E.dataAction, 
			@iTrueFlowType = isnull(E.trueFlowType, 0), 
			@iExprID = isnull(E.trueFlowExprID, 0), 
			@sEmailSubject = ISNULL(E.emailSubject, ''''),
			@bUseAsTargetIdentifier = ISNULL(E.UseAsTargetIdentifier, 0)
		FROM ASRSysWorkflowElements E
		WHERE E.ID = @piElementID;

		--------------------------------------------------
		-- Read the submitted webForm/storedData values
		--------------------------------------------------
		IF @iElementType = 5 -- Stored Data element
		BEGIN
			SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
			SET @sValue = LEFT(@psFormInput1, @iIndex1-1);
			SET @sTemp = SUBSTRING(@psFormInput1, @iIndex1+1, LEN(@psFormInput1) - @iIndex1);

			SET @sValueDescription = '''';
			SET @sMessage = ''Successfully '' +
				CASE
					WHEN @iDataAction = 0 THEN ''inserted''
					WHEN @iDataAction = 1 THEN ''updated''
					ELSE ''deleted''
				END + '' record'';

			IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
			BEGIN
				SET @sValueDescription = @sTemp;
			END
			ELSE
			BEGIN
				SET @iTemp = convert(integer, @sValue);
				IF @iTemp > 0 
				BEGIN	
					EXEC [dbo].[spASRRecordDescription] 
						@iTableID,
						@iTemp,
						@sEvalRecDesc OUTPUT
					IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
				END
			END

			IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')'';

			UPDATE ASRSysWorkflowInstanceValues
			SET ASRSysWorkflowInstanceValues.value = @sValue, 
				ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @piElementID
				AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
				AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		END
		ELSE
		BEGIN
			-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
			WHILE (charindex(CHAR(9), @psFormInput1) > 0)
			BEGIN

				SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
				SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1);
				SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''');
				SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1);
				SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2);

				--Get the record description (for RecordSelectors only)
				SET @sValueDescription = '''';

				-- Get the WebForm item type, etc.
				SELECT @sIdentifier = EI.identifier,
					@iItemType = EI.itemType,
					@iTableID = EI.tableID,
					@iBehaviour = EI.behaviour,
					@bUseAsTargetIdentifier = ISNULL(EI.UseAsTargetIdentifier, 0)
				FROM ASRSysWorkflowElementItems EI
				WHERE EI.ID = convert(integer, @sID);

				SET @iParent1TableID = 0;
				SET @iParent1RecordID = 0;
				SET @iParent2TableID = 0;
				SET @iParent2RecordID = 0;

				IF @iItemType = 11 -- Record Selector
				BEGIN
					-- Get the table record description ID. 
					SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
					FROM ASRSysTables 
					WHERE ASRSysTables.tableID = @iTableID;

					SET @iTemp = convert(integer, isnull(@sValue, ''0''));

					-- Get the record description. 
					IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
					BEGIN
						SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(MAX), @iRecDescID) + '' @recDesc OUTPUT, @recID'';
						SET @sParamDefinition = N''@recDesc varchar(MAX) OUTPUT, @recID integer'';
						EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp;
						IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
					END

					IF @bUseAsTargetIdentifier = 1
							UPDATE ASRSysWorkflowInstances SET TargetName = @sEvalRecDesc	WHERE ID = @piInstanceID;

					-- Record the selected record''s parent details.
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iTemp,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
				ELSE
				IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
				BEGIN
					SET @pfSavedForLater = 1;
				END

				IF (@iItemType = 17) -- FileUpload Control
				BEGIN
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.fileUpload_File = 
						CASE 
							WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_File
							ELSE null
						END,
						ASRSysWorkflowInstanceValues.fileUpload_ContentType = 
						CASE 
							WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_ContentType
							ELSE null
						END,
						ASRSysWorkflowInstanceValues.fileUpload_FileName = 
						CASE 
							WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_FileName
							ELSE null
						END
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
				END
				ELSE
				BEGIN
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
						ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
						ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
						ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
						ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
				END
			END

			IF @pfSavedForLater = 1
			BEGIN
				/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 7
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
				
				/* Remember the page number too  */
				UPDATE ASRSysWorkflowInstances
				SET ASRSysWorkflowInstances.pageno = @piPageNo
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID;

				RETURN;
			END
		END
			
		SET @hResult = 0;
		SET @sTo = '''';
		SET @sCopyTo = '''';

		--------------------------------------------------
		-- Process email element
		--------------------------------------------------
		IF @iElementType = 3 -- Email element
		BEGIN
			-- Get the email recipient. 
			SET @iEmailRecordID = 0;
			SET @sSQL = ''spASRSysEmailAddr'';

			IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
			BEGIN
				SET @iEmailLoop = 0
				WHILE @iEmailLoop < 2
				BEGIN
					SET @hTmpResult = 0;
					SET @sTempTo = '''';
					SET @iTempEmailID = 
						CASE 
							WHEN @iEmailLoop = 1 THEN @iEmailCopyID
							ELSE isnull(@iEmailID, 0)
						END;

					IF @iTempEmailID > 0 
					BEGIN
						SET @fValidRecordID = 1;

						SELECT @iEmailTableID = isnull(tableID, 0),
							@iEmailType = isnull(type, 0)
						FROM ASRSysEmailAddress
						WHERE emailID = @iTempEmailID;

						IF @iEmailType = 0 
						BEGIN
							SET @iEmailRecordID = 0;
						END
						ELSE
						BEGIN
							SET @iTempElementID = 0;

							-- Get the record ID required. 
							IF (@iEmailRecord = 0) OR (@iEmailRecord = 4)
							BEGIN
								/* Initiator record. */
								SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID,
									@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
									@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
									@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
									@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
								FROM ASRSysWorkflowInstances
								WHERE ASRSysWorkflowInstances.ID = @piInstanceID;

								SET @iBaseRecordID = @iEmailRecordID;

								IF @iEmailRecord = 4
								BEGIN
									-- Trigger record
									SELECT @iBaseTableID = isnull(WF.baseTable, 0)
									FROM ASRSysWorkflows WF
									INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
										AND WFI.ID = @piInstanceID;
								END
								ELSE
								BEGIN
									-- Initiator''s record
									SELECT @iBaseTableID = convert(integer, ISNULL(parameterValue, ''0''))
									FROM ASRSysModuleSetup
									WHERE moduleKey = ''MODULE_PERSONNEL''
										AND parameterKey = ''Param_TablePersonnel'';

									IF @iBaseTableID = 0
									BEGIN
										SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
										FROM ASRSysModuleSetup
										WHERE moduleKey = ''MODULE_WORKFLOW''
										AND parameterKey = ''Param_TablePersonnel'';
									END
								END
							END

							IF @iEmailRecord = 1
							BEGIN
								SELECT @iPrevElementType = ASRSysWorkflowElements.type,
									@iTempElementID = ASRSysWorkflowElements.ID
								FROM ASRSysWorkflowElements
								WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
									AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));

								IF @iPrevElementType = 2
								BEGIN
									 -- WebForm
									SELECT @sValue = ISNULL(IV.value, ''0''),
										@iBaseTableID = EI.tableID,
										@iParent1TableID = IV.parent1TableID,
										@iParent1RecordID = IV.parent1RecordID,
										@iParent2TableID = IV.parent2TableID,
										@iParent2RecordID = IV.parent2RecordID
									FROM ASRSysWorkflowInstanceValues IV
									INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
									INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
									WHERE IV.instanceID = @piInstanceID
										AND IV.identifier = @sRecSelIdentifier
										AND Es.identifier = @sRecSelWebFormIdentifier
										AND Es.workflowID = @iWorkflowID
										AND IV.elementID = Es.ID;
								END
								ELSE
								BEGIN
									-- StoredData
									SELECT @sValue = ISNULL(IV.value, ''0''),
										@iBaseTableID = isnull(Es.dataTableID, 0),
										@iParent1TableID = IV.parent1TableID,
										@iParent1RecordID = IV.parent1RecordID,
										@iParent2TableID = IV.parent2TableID,
										@iParent2RecordID = IV.parent2RecordID
									FROM ASRSysWorkflowInstanceValues IV
									INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
										AND IV.identifier = Es.identifier
										AND Es.workflowID = @iWorkflowID
										AND Es.identifier = @sRecSelWebFormIdentifier
									WHERE IV.instanceID = @piInstanceID;
								END

								SET @iEmailRecordID = 
									CASE
										WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
										ELSE 0
									END;

								SET @iBaseRecordID = @iEmailRecordID;
							END

							SET @fValidRecordID = 1;
							IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
							BEGIN
								SET @fValidRecordID = 0;

								EXEC [dbo].[spASRWorkflowAscendantRecordID]
									@iBaseTableID,
									@iBaseRecordID,
									@iParent1TableID,
									@iParent1RecordID,
									@iParent2TableID,
									@iParent2RecordID,
									@iEmailTableID,
									@iRequiredRecordID	OUTPUT;

								SET @iEmailRecordID = @iRequiredRecordID;

								IF @iRequiredRecordID > 0 
								BEGIN
									EXEC [dbo].[spASRWorkflowValidTableRecord]
										@iEmailTableID,
										@iEmailRecordID,
										@fValidRecordID	OUTPUT;
								END

								IF @fValidRecordID = 0
								BEGIN
									IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
									BEGIN
										SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
										FROM ASRSysWorkflowQueueColumns QC
										INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
										WHERE WFQ.instanceID = @piInstanceID
											AND QC.emailID = @iTempEmailID;

										IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
									END
									ELSE
									BEGIN
										IF @iEmailRecord = 1
										BEGIN
											SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
											FROM ASRSysWorkflowInstanceValues IV
											WHERE IV.instanceID = @piInstanceID
												AND IV.emailID = @iTempEmailID
												AND IV.elementID = @iTempElementID;

											IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
										END
									END
								END

								IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
								BEGIN
									-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
									EXEC [dbo].[spASRWorkflowActionFailed] 
										@piInstanceID, 
										@piElementID, 
										''Email record has been deleted or not selected.'';
											
									SET @hTmpResult = -1;
								END
							END
						END

						IF @fValidRecordID = 1
						BEGIN
							/* Get the recipient address. */
							IF len(@sTempTo) = 0
							BEGIN
								EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID;
								IF @sTempTo IS null SET @sTempTo = '''';
							END

							IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
							BEGIN
								-- Email step failure if no known recipient.
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC [dbo].[spASRWorkflowActionFailed] 
									@piInstanceID, 
									@piElementID, 
									''No email recipient.'';
										
								SET @hTmpResult = -1;
							END
						END

						IF @iEmailLoop = 1 
						BEGIN
							SET @sCopyTo = @sTempTo;

							IF (rtrim(ltrim(@sCopyTo)) = ''@'')
								OR (charindex('' @ '', @sCopyTo) > 0)
							BEGIN
								SET @sCopyTo = '''';
							END
						END
						ELSE
						BEGIN
							SET @sTo = @sTempTo;
						END
					END
				
					SET @iEmailLoop = @iEmailLoop + 1;

					IF @hTmpResult <> 0 SET @hResult = @hTmpResult;
				END
			END

			IF LEN(rtrim(ltrim(@sTo))) > 0
			BEGIN
				IF (rtrim(ltrim(@sTo)) = ''@'')
					OR (charindex('' @ '', @sTo) > 0)
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.userEmail = @sTo,
						ASRSysWorkflowInstanceSteps.emailCC = @sCopyTo
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;

					EXEC [dbo].[spASRWorkflowActionFailed] 
						@piInstanceID, 
						@piElementID, 
						''Invalid email recipient.'';
				
					SET @hResult = -1;
				END
				ELSE
				BEGIN
					/* Build the email message. */
					EXEC [dbo].[spASRGetWorkflowEmailMessage] 
						@piInstanceID, 
						@piElementID, 
						@sMessage OUTPUT, 
						@sMessage_HypertextLinks OUTPUT, 
						@sHypertextLinkedSteps OUTPUT, 
						@fValidRecordID OUTPUT, 
						@sTo;

					IF @fValidRecordID = 1
					BEGIN
						exec [dbo].[spASRDelegateWorkflowEmail] 
							@sTo,
							@sCopyTo,
							@sMessage,
							@sMessage_HypertextLinks,
							@iCurrentStepID,
							@sEmailSubject;
					END
					ELSE
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed] 
							@piInstanceID, 
							@piElementID, 
							''Email item database value record has been deleted or not selected.'';
								
						SET @hResult = -1;
					END
				END
			END
		END

		--------------------------------------------------
		-- Mark the step as complete
		--------------------------------------------------
		IF @hResult = 0
		BEGIN
			/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 3,
				ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.userEmail = CASE
					WHEN @iElementType = 3 THEN @sTo
					ELSE ASRSysWorkflowInstanceSteps.userEmail
				END,
				ASRSysWorkflowInstanceSteps.emailCC = CASE
					WHEN @iElementType = 3 THEN @sCopyTo
					ELSE ASRSysWorkflowInstanceSteps.emailCC
				END,
				ASRSysWorkflowInstanceSteps.hypertextLinkedSteps = CASE
					WHEN @iElementType = 3 THEN @sHypertextLinkedSteps
					ELSE ASRSysWorkflowInstanceSteps.hypertextLinkedSteps
				END,
				ASRSysWorkflowInstanceSteps.message = CASE
					WHEN @iElementType = 3 THEN @sMessage
					WHEN @iElementType = 5 THEN @sMessage
					ELSE ''''
				END,
				ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
	
			IF @iElementType = 4 -- Decision element
			BEGIN
				IF @iTrueFlowType = 1
				BEGIN
					-- Decision Element flow determined by a calculation
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iExprID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0;

					SET @iValue = convert(integer, @fResult);
				END
				ELSE
				BEGIN
					-- Decision Element flow determined by a button in a preceding web form
					SET @iPrevElementType = 4; -- Decision element
					SET @iPreviousElementID = @piElementID;

					WHILE (@iPrevElementType = 4)
					BEGIN
						SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
							@iPrevElementType = isnull(WE.type, 0)
						FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
						INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
						INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
							AND WIS.instanceID = @piInstanceID;

						SET @iPreviousElementID = @iTempID;
					END
			
					SELECT @sValue = ISNULL(IV.value, ''0'')
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
					WHERE IV.elementID = @iPreviousElementID
						AND IV.instanceid = @piInstanceID
						AND E.ID = @piElementID;

					SET @iValue = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END;
				END
		
				IF @iValue IS null SET @iValue = 0;

				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
	
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT SUCC.id FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, @iValue) SUCC)
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 2
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8
						OR ASRSysWorkflowInstanceSteps.status = 3);
			END
			ELSE
			BEGIN
				IF @iElementType <> 3 -- 3=Email element
				BEGIN
					IF @iElementType = 2 -- WebForm
					BEGIN
						SELECT @sUserName = isnull(WIS.userName, ''''),
							@sUserEmail = isnull(WIS.userEmail, '''')
						FROM ASRSysWorkflowInstanceSteps WIS
						WHERE WIS.instanceID = @piInstanceID
							AND WIS.elementID = @piElementID;
					END;
							
					-- Do not the following bit when the submitted element is an Email element as 
					-- the succeeding elements will already have been actioned.
					DECLARE @succeedingElements TABLE(elementID integer);

					EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
						@piInstanceID, 
						@piElementID, 
						@superCursor OUTPUT,
						'''';

					FETCH NEXT FROM @superCursor INTO @iTemp;
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO @succeedingElements (elementID) VALUES (@iTemp);
					
						FETCH NEXT FROM @superCursor INTO @iTemp;
					END
					CLOSE @superCursor;
					DEALLOCATE @superCursor;

					-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
					IF @iElementType = 2 -- WebForm
					BEGIN
						-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
						DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT ASRSysWorkflowInstanceSteps.ID,
							ASRSysWorkflowInstanceSteps.elementID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT suc.elementID
								FROM @succeedingElements suc)
							AND ASRSysWorkflowElements.type = 2
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 6
								OR ASRSysWorkflowInstanceSteps.status = 8
								OR ASRSysWorkflowInstanceSteps.status = 3);

						OPEN formsCursor;
						FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
						WHILE (@@fetch_status = 0) 
						BEGIN
							SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);

							DELETE FROM ASRSysWorkflowStepDelegation
							WHERE stepID = @iStepID;

							INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
								(SELECT WSD.delegateEmail, @iStepID
								FROM ASRSysWorkflowStepDelegation WSD
								WHERE WSD.stepID = @iCurrentStepID);
						
							-- Change the step status to be 2 (pending user input). 
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 2, 
								ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.completionDateTime = null,
								ASRSysWorkflowInstanceSteps.userName = @sUserName,
								ASRSysWorkflowInstanceSteps.userEmail = @sUserEmail 
							WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3);
						
							FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
						END
						CLOSE formsCursor;
						DEALLOCATE formsCursor;

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.completionDateTime = null
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT suc.elementID
								FROM @succeedingElements suc)
							AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
								(SELECT ASRSysWorkflowElements.ID
								FROM ASRSysWorkflowElements
								WHERE ASRSysWorkflowElements.type = 2)
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 6
								OR ASRSysWorkflowInstanceSteps.status = 8
								OR ASRSysWorkflowInstanceSteps.status = 3);
					END
					ELSE
					BEGIN
						DELETE FROM ASRSysWorkflowStepDelegation
						WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
							FROM ASRSysWorkflowInstanceSteps
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3));
					
						INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
						(SELECT WSD.delegateEmail,
							SuccWIS.ID
						FROM ASRSysWorkflowStepDelegation WSD
						INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
						INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
							AND SuccWIS.elementID IN (SELECT suc.elementID
								FROM @succeedingElements suc)
							AND (SuccWIS.status = 0
								OR SuccWIS.status = 2
								OR SuccWIS.status = 6
								OR SuccWIS.status = 8
								OR SuccWIS.status = 3)
						INNER JOIN ASRSysWorkflowElements SuccWE ON SuccWIS.elementID = SuccWE.ID
							AND SuccWE.type = 2
						WHERE WSD.stepID = @iCurrentStepID);

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.completionDateTime = null,
							ASRSysWorkflowInstanceSteps.userEmail = CASE
								WHEN (SELECT ASRSysWorkflowElements.type 
									FROM ASRSysWorkflowElements 
									WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
								ELSE ASRSysWorkflowInstanceSteps.userEmail
							END
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT suc.elementID
								FROM @succeedingElements suc)
							AND (ASRSysWorkflowInstanceSteps.status = 0
								OR ASRSysWorkflowInstanceSteps.status = 2
								OR ASRSysWorkflowInstanceSteps.status = 6
								OR ASRSysWorkflowInstanceSteps.status = 8
								OR ASRSysWorkflowInstanceSteps.status = 3);
					END
				END
			END
	
			-- Set activated Web Forms to be ''pending'' (to be done by the user) 
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 2
			WHERE ASRSysWorkflowInstanceSteps.id IN (
				SELECT ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2);

			-- Set activated Terminators to be ''completed'' 
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 3,
				ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
			WHERE ASRSysWorkflowInstanceSteps.id IN (
				SELECT ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowElements.type = 1);

			-- Count how many terminators have completed. ie. if the workflow has completed. 
			SELECT @iCount = COUNT(*)
			FROM ASRSysWorkflowInstanceSteps
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.status = 3
				AND ASRSysWorkflowElements.type = 1;
					
			IF @iCount > 0 
			BEGIN
				UPDATE ASRSysWorkflowInstances
				SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
					ASRSysWorkflowInstances.status = 3,
					ASRSysWorkflowInstances.pageno = @piPageNo
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
			
				-- Steps pending action are no longer required.
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
						OR ASRSysWorkflowInstanceSteps.status = 2); -- 2 = Pending User Action
			END

			IF @iElementType = 3 -- Email element
				OR @iElementType = 5 -- Stored Data element
			BEGIN
				exec [dbo].[spASREmailImmediate] ''OpenHR Workflow'';
			END
		END
	END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRMobileInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRMobileInstantiateWorkflow];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRMobileInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@psKeyParameter	varchar(max),			
			@psPWDParameter	varchar(max),			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(MAX)	OUTPUT,
			@psMessage		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID			integer,
				@iStepID				integer,
				@iElementID				integer,
				@iRecordID				integer,
				@iRecordCount			integer,
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@bUseAsTargetIdentifier bit,
				@iTemp					integer,
				@iStartElementID		integer,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms					varchar(MAX),
				@iCount					integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled				bit,
				@iElementType			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(255),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sSPName				varchar(255),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fSaveForLater			bit,
				@fResult	bit;
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
		
			DECLARE @succeedingElements table(elementID int);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT
			-- @fExternallyInitiated = CASE
			--		WHEN initiationType = 2 THEN 1
			--		ELSE 0
			--	END,
				@fEnabled = enabled
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;

			--IF @fExternallyInitiated = 1
			--BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = @psKeyParameter;
			--END
			--ELSE
			--BEGIN
				--SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysMobileGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@psKeyParameter,			
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT;
				END
			
			print @iRecordID;
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC [dbo].[spASRWorkflowUsesInitiator]
						@piWorkflowID,
						@fUsesInitiator OUTPUT;
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.'';
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.'';
						END
			
						RETURN
					END	
				END
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel'';
					END
		
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
			--END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				pageno)
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = MAX(id)
			FROM [dbo].[ASRSysWorkflowInstances];
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
		
			INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0);
		
			INSERT INTO [dbo].[ASRSysWorkflowInstanceSteps] (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 1
					ELSE 0
				END,
				0,
				0
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID;
		
			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5;
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
					
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
				OPEN immediateSubmitCursor;
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult	OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''spASRWorkflowInsertNewRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL;
		
								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''spASRWorkflowUpdateRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC [dbo].[spASRRecordDescription]
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''spASRWorkflowDeleteRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							EXEC [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT;
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(MAX), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 4,	-- 4 = failed
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
		
								SET @psMessage = @sStoredDataMsg;
								RETURN;
							END
							ELSE
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 8,	-- 8 = failed action
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE [instanceID] = @piInstanceID
									AND [elementID] = @iElementID;
		
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
									SET ASRSysWorkflowInstanceSteps.status = 1
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT,
							0;
					END
		
					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				END
				CLOSE immediateSubmitCursor;
				DEALLOCATE immediateSubmitCursor;
		
				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM [dbo].[ASRSysWorkflowInstanceSteps]
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;
			END						
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE @succeedingSteps table(stepID int)
			
			INSERT INTO @succeedingSteps 
				(stepID) VALUES (-1)
		
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM [dbo].[ASRSysWorkflowInstanceSteps]
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
			OPEN formsCursor;
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
				INSERT INTO @succeedingSteps 
				(stepID) VALUES (@iStepID)
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			END
		
			CLOSE formsCursor;
			DEALLOCATE formsCursor;
		
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET ASRSysWorkflowInstanceSteps.status = 2, 
				userName = @sActualLoginName
			WHERE ASRSysWorkflowInstanceSteps.ID IN (SELECT stepID FROM @succeedingSteps)
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
(
	@piInstanceID		integer,
	@piElementID		integer,
	@succeedingElements	cursor varying output,
	@psTo				varchar(MAX)
)
AS
BEGIN
	-- Action any immediate elements (Or, Decision and StoredData elements) and return the IDs of the workflow elements that 
	-- succeed them.
	-- This ignores connection elements.
	DECLARE
		@iTempID				integer,
		@iElementID				integer,
		@iElementType			integer,
		@iFlowCode				integer,
		@bUseAsTargetIdentifier	bit,
		@iTrueFlowType			integer,
		@iExprID				integer,
		@iResultType			integer,
		@sValue					varchar(MAX),
		@sResult				varchar(MAX),
		@fResult				bit,
		@dtResult				datetime,
		@fltResult				float,
		@iValue					integer,
		@iPrecedingElementType	integer, 
		@iPrecedingElementID	integer, 
		@iCount					integer,
		@iStepID				integer,
		@curRecipients			cursor,
		@sEmailAddress			varchar(MAX),
		@fDelegated				bit,
		@sDelegatedTo			varchar(MAX),
		@iSQLVersion			integer,
		@fStoredDataOK			bit, 
		@sStoredDataMsg			varchar(MAX), 
		@sStoredDataSQL			varchar(MAX), 
		@iStoredDataTableID		integer,
		@sStoredDataTableName	varchar(MAX),
		@iStoredDataAction		integer, 
		@iStoredDataRecordID	integer,
		@sStoredDataRecordDesc	varchar(MAX),
		@sStoredDataWebForms	varchar(MAX),
		@sStoredDataSaveForLater bit,
		@sSPName				varchar(MAX),
		@iNewRecordID			integer,
		@sEvalRecDesc			varchar(MAX),
		@iResult				integer,
		@iFailureFlows			integer,
		@fDeadlock				bit,
		@iErrorNumber			integer,
		@iRetryCount			integer,
		@iDEADLOCKERRORNUMBER	integer,
		@iMAXRETRIES			integer,
		@fIsDelegate			bit;

	SET @iDEADLOCKERRORNUMBER = 1205;
	SET @iMAXRETRIES = 5;
					
	SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
					
	DECLARE @elements table
	(
		elementID		integer,
		elementType		integer,
		processed		tinyint default 0,
		trueFlowType	integer,
		trueFlowExprID	integer
	);
					
	INSERT INTO @elements 
		(elementID,
		elementType,
		processed,
		trueFlowType,
		trueFlowExprID)
	SELECT SUCC.id,
		E.type,
		0,
		ISNULL(E.trueFlowType, 0),
		ISNULL(E.trueFlowExprID, 0)
	FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 0) SUCC
	INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID;
		
	SELECT @iCount = COUNT(*)
	FROM @elements
	WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
		AND processed = 0;

	WHILE @iCount > 0
	BEGIN
		UPDATE @elements
		SET processed = 1
		WHERE processed = 0;

		-- Action any succeeding immediate elements (Decision, Or and StoredData elements)
		DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT E.elementID,
			E.elementType,
			E.trueFlowType, 
			E.trueFlowExprID
		FROM @elements E
		WHERE (E.elementType = 4 OR (@iSQLVersion >= 9 AND E.elementType = 5) OR E.elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
			AND E.processed = 1;

		OPEN immediateCursor;
		FETCH NEXT FROM immediateCursor INTO 
			@iElementID, 
			@iElementType, 
			@iTrueFlowType, 
			@iExprID;
		WHILE (@@fetch_status = 0)
		BEGIN
			-- Submit the immediate elements, and get their succeeding elements
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 3,
				ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
				ASRSysWorkflowInstanceSteps.message = '''',
				ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

			SET @iFlowCode = 0;

			IF @iElementType = 4 -- Decision
			BEGIN
				IF @iTrueFlowType = 1
				BEGIN
					-- Decision Element flow determined by a calculation
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iExprID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0;

					SET @iValue = convert(integer, @fResult);
				END
				ELSE
				BEGIN
					-- Decision Element flow determined by a button in a preceding web form
					SET @iPrecedingElementType = 4; -- Decision element
					SET @iPrecedingElementID = @iElementID;

					WHILE (@iPrecedingElementType = 4)
					BEGIN
						SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
							@iPrecedingElementType = isnull(WE.type, 0)
						FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPrecedingElementID) PE
						INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
						INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
							AND WIS.instanceID = @piInstanceID;

						SET @iPrecedingElementID = @iTempID;
					END
					
					SELECT @sValue = ISNULL(IV.value, ''0'')
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
					WHERE IV.elementID = @iPrecedingElementID
					AND IV.instanceid = @piInstanceID
						AND E.ID = @iElementID;

					SET @iValue = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END;
				END
				
				IF @iValue IS null SET @iValue = 0;
				SET @iFlowCode = @iValue;

				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
			END
			ELSE IF @iElementType = 7 -- Or
			BEGIN
				EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @piInstanceID, @iElementID;
			END
			ELSE IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
			BEGIN
				SET @fStoredDataOK = 1;
				SET @sStoredDataMsg = '''';
				SET @sStoredDataRecordDesc = '''';

				EXEC [spASRGetStoredDataActionDetails]
					@piInstanceID,
					@iElementID,
					@sStoredDataSQL			OUTPUT, 
					@iStoredDataTableID		OUTPUT,
					@sStoredDataTableName	OUTPUT,
					@iStoredDataAction		OUTPUT, 
					@iStoredDataRecordID	OUTPUT,
					@bUseAsTargetIdentifier OUTPUT,
					@fResult OUTPUT;

				IF @fResult = 1
				BEGIN
					IF @iStoredDataAction = 0 -- Insert
					BEGIN
						SET @sSPName  = ''sp_ASRInsertNewRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@sStoredDataSQL;

								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE IF @iStoredDataAction = 1 -- Update
					BEGIN
						SET @sSPName  = ''sp_ASRUpdateRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@sStoredDataSQL,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID,
									null;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE IF @iStoredDataAction = 2 -- Delete
					BEGIN
						EXEC spASRRecordDescription
							@iStoredDataTableID,
							@iStoredDataRecordID,
							@sStoredDataRecordDesc OUTPUT;

						SET @sSPName  = ''sp_ASRDeleteRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @iErrorNumber = ERROR_NUMBER();

								IF @iErrorNumber = @iDEADLOCKERRORNUMBER
								BEGIN
									IF @iRetryCount < @iMAXRETRIES
									BEGIN
										SET @iRetryCount = @iRetryCount + 1;
										SET @fDeadlock = 1;
										--Sleep for 5 seconds
										WAITFOR DELAY ''00:00:05'';
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END
								ELSE
								BEGIN
									SET @fStoredDataOK = 0;
									SET @sStoredDataMsg = ERROR_MESSAGE();
								END
							END CATCH
						END
					END
					ELSE
					BEGIN
						SET @fStoredDataOK = 0;
						SET @sStoredDataMsg = ''Unrecognised data action.'';
					END

					IF (@fStoredDataOK = 1)
						AND ((@iStoredDataAction = 0)
							OR (@iStoredDataAction = 1))
					BEGIN

						exec [dbo].[spASRStoredDataFileActions]
							@piInstanceID,
							@iElementID,
							@iStoredDataRecordID;
					END

					IF @fStoredDataOK = 1
					BEGIN
						SET @sStoredDataMsg = ''Successfully '' +
							CASE
								WHEN @iStoredDataAction = 0 THEN ''inserted''
								WHEN @iStoredDataAction = 1 THEN ''updated''
								ELSE ''deleted''
							END + '' record'';

						IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
						BEGIN
							IF @iStoredDataRecordID > 0 
							BEGIN	
								EXEC [dbo].[spASRRecordDescription] 
									@iStoredDataTableID,
									@iStoredDataRecordID,
									@sEvalRecDesc OUTPUT
								IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
							END
						END

						IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';

						UPDATE ASRSysWorkflowInstanceValues
						SET ASRSysWorkflowInstanceValues.value = convert(varchar(255), @iStoredDataRecordID), 
							ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
						WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceValues.elementID = @iElementID
							AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
							AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 3,
							ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

						IF @bUseAsTargetIdentifier = 1
						BEGIN
							EXEC [dbo].[spASRRecordDescription] @iStoredDataTableID, @iStoredDataRecordID, @sEvalRecDesc OUTPUT;
							UPDATE ASRSysWorkflowInstances SET TargetName = @sEvalRecDesc WHERE ID = @piInstanceID;
						END

					END
					ELSE
					BEGIN
						-- Check if the failed element has an outbound flow for failures.
						SELECT @iFailureFlows = COUNT(*)
						FROM ASRSysWorkflowElements Es
						INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
							AND Ls.startOutboundFlowCode = 1
						WHERE Es.ID = @iElementID
							AND Es.type = 5; -- 5 = StoredData

						IF @iFailureFlows = 0
						BEGIN
							UPDATE ASRSysWorkflowInstanceSteps
							SET status = 4,	-- 4 = failed
								message = @sStoredDataMsg,
								failedCount = isnull(failedCount, 0) + 1,
								completionCount = isnull(completionCount, 0) - 1
							WHERE instanceID = @piInstanceID
								AND elementID = @iElementID;

							UPDATE ASRSysWorkflowInstances
							SET status = 2	-- 2 = error
							WHERE ID = @piInstanceID;
						END
						ELSE
						BEGIN
							UPDATE ASRSysWorkflowInstanceSteps
							SET status = 8,	-- 8 = failed action
								message = @sStoredDataMsg,
								failedCount = isnull(failedCount, 0) + 1,
								completionCount = isnull(completionCount, 0) - 1
							WHERE instanceID = @piInstanceID
								AND elementID = @iElementID;

							INSERT INTO @elements 
								(elementID,
								elementType,
								processed,
								trueFlowType,
								trueFlowExprID)
							SELECT SUCC.id,
								E.type,
								0,
								isnull(E.trueFlowType, 0),
								isnull(E.trueFlowExprID, 0)
							FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1) SUCC
							INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
							WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
						END
					END
				END
				ELSE				
				BEGIN
					SET @fStoredDataOK = 0;

					-- Check if the failed element has an outbound flow for failures.
					SELECT @iFailureFlows = COUNT(*)
					FROM ASRSysWorkflowElements Es
					INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
						AND Ls.startOutboundFlowCode = 1
					WHERE Es.ID = @iElementID
						AND Es.type = 5; -- 5 = StoredData

					IF @iFailureFlows = 0
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET completionCount = isnull(completionCount, 0) - 1
						WHERE instanceID = @piInstanceID
							AND elementID = @iElementID;
					END
					ELSE
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET completionCount = isnull(completionCount, 0) - 1
						WHERE instanceID = @piInstanceID
							AND elementID = @iElementID;

						INSERT INTO @elements 
							(elementID,
							elementType,
							processed,
							trueFlowType,
							trueFlowExprID)
						SELECT SUCC.id,
							E.type,
							0,
							isnull(E.trueFlowType, 0),
							isnull(E.trueFlowExprID, 0)
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1) SUCC
						INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
						WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
					END
				END;
			END

			IF (@iElementType <> 5) OR (@fStoredDataOK = 1)
			BEGIN
				-- Get this immediate element''s succeeding elements
				INSERT INTO @elements 
					(elementID,
					elementType,
					processed,
					trueFlowType,
					trueFlowExprID)
				SELECT SUCC.id,
					E.type,
					0,
					isnull(E.trueFlowType, 0),
					isnull(E.trueFlowExprID, 0)
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, @iFlowCode) SUCC
				INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
				WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
			END

			FETCH NEXT FROM immediateCursor INTO 
				@iElementID, 
				@iElementType, 
				@iTrueFlowType, 
				@iExprID;
		END
		CLOSE immediateCursor;
		DEALLOCATE immediateCursor;

		UPDATE @elements
		SET processed = 2
		WHERE processed = 1;

		SELECT @iCount = COUNT(*)
		FROM @elements
		WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
			AND processed = 0;
	END

	SELECT @iCount = COUNT(*)
	FROM @elements
	WHERE elementType = 2; -- 2=WebForm

	IF (@iCount > 0) AND len(ltrim(rtrim(@psTo))) > 0 
	BEGIN
		SELECT @iStepID = ASRSysWorkflowInstanceSteps.ID
		FROM ASRSysWorkflowInstanceSteps
		WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
			AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;

		DECLARE @recipients TABLE (
			emailAddress	varchar(MAX),
			delegated		bit,
			delegatedTo		varchar(MAX),
			isDelegate		bit
		);

		exec [dbo].[spASRGetWorkflowDelegates] 
			@psTo, 
			@iStepID, 
			@curRecipients output;
		FETCH NEXT FROM @curRecipients INTO 
				@sEmailAddress,
				@fDelegated,
				@sDelegatedTo,
				@fIsDelegate;
		WHILE (@@fetch_status = 0)
		BEGIN
			INSERT INTO @recipients
				(emailAddress,
				delegated,
				delegatedTo,
				isDelegate)
			VALUES (
				@sEmailAddress,
				@fDelegated,
				@sDelegatedTo,
				@fIsDelegate
			);
			
			FETCH NEXT FROM @curRecipients INTO 
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate;
		END
		CLOSE @curRecipients;
		DEALLOCATE @curRecipients;

		DELETE FROM ASRSysWorkflowStepDelegation
		WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType = 2) -- 2 = WebForm
				AND (ASRSysWorkflowInstanceSteps.status = 0
					OR ASRSysWorkflowInstanceSteps.status = 2
					OR ASRSysWorkflowInstanceSteps.status = 6
					OR ASRSysWorkflowInstanceSteps.status = 8
					OR ASRSysWorkflowInstanceSteps.status = 3));

		INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
		SELECT DISTINCT RECS.emailAddress, WIS.ID
		FROM @recipients RECS, 
			ASRSysWorkflowInstanceSteps WIS
		WHERE RECS.isDelegate = 1
			AND WIS.instanceID = @piInstanceID
				AND WIS.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType = 2) -- 2 = WebForm
				AND (WIS.status = 0
					OR WIS.status = 2
					OR WIS.status = 6
					OR WIS.status = 8
					OR WIS.status = 3);
	END

	UPDATE ASRSysWorkflowInstanceSteps
	SET ASRSysWorkflowInstanceSteps.status = 1,
		ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
		ASRSysWorkflowInstanceSteps.completionDateTime = null,
		ASRSysWorkflowInstanceSteps.userEmail = CASE
			WHEN (SELECT ASRSysWorkflowElements.type 
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @psTo -- 2 = Web Form element
			ELSE ASRSysWorkflowInstanceSteps.userEmail
		END
	WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceSteps.elementID IN 
			(SELECT E.elementID
			FROM @elements E
			WHERE E.elementType <> 7 -- 7 = Or
				AND (E.elementType <> 5 OR @iSQLVersion <= 8) -- 5 = StoredData
				AND E.elementType <> 4) -- 4 = Decision
		AND (ASRSysWorkflowInstanceSteps.status = 0
			OR ASRSysWorkflowInstanceSteps.status = 2
			OR ASRSysWorkflowInstanceSteps.status = 6
			OR ASRSysWorkflowInstanceSteps.status = 8
			OR ASRSysWorkflowInstanceSteps.status = 3);

	UPDATE ASRSysWorkflowInstanceSteps
	SET ASRSysWorkflowInstanceSteps.status = 2
	WHERE ASRSysWorkflowInstanceSteps.id IN (
		SELECT ASRSysWorkflowInstanceSteps.ID
		FROM ASRSysWorkflowInstanceSteps
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowInstanceSteps.status = 1
			AND ASRSysWorkflowElements.type = 2);

	-- Return the cursor of succeeding elements. 
	SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
		SELECT elementID 
		FROM @elements E
		WHERE E.elementType <> 7 -- 7 = Or
			AND E.elementType <> 4 -- 4 = Decision
			AND (E.elementType <> 5 OR @iSQLVersion <= 8); -- 5 = StoredData

	OPEN @succeedingElements;
END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRInstantiateTriggeredWorkflows]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows];
	EXECUTE sp_executesql N'CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE
				@iQueueID			integer,
				@iWorkflowID		integer,
				@iRecordID			integer,
				@iInstanceID		integer,
				@iStartElementID	integer,
				@iTemp				integer,
				@iBaseTable		integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer,
				@TargetName varchar(MAX);

			DECLARE @succeedingElements table(elementID int)
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID,
				Q.parent1TableID,
				Q.parent1RecordID,
				Q.parent2TableID,
				Q.parent2RecordID,
				WF.baseTable
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID;

				EXEC [dbo].[sp_ASRIntGetRecordDescription] @iBaseTable, @iRecordID, 0, 0, @TargetName OUTPUT;
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, 
					initiatorID, 
					status, 
					userName, 
					parent1TableID,
					parent1RecordID,
					parent2TableID,
					parent2RecordID,
					pageno,
					TargetName)
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					''<Triggered>'',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					0,
					@TargetName)
								
				SELECT @iInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
				
				UPDATE ASRSysWorkflowQueue
				SET instanceID = @iInstanceID
				WHERE queueID = @iQueueID	

				-- Create the Workflow Instance Steps records. 
				-- Set the first steps'' status to be 1 (pending Workflow Engine action). 
				-- Set all subsequent steps'' status to be 0 (on hold). */
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @iWorkflowID
		
				INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
		
				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
				SELECT 
					@iInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 3
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 1
						ELSE 0
					END,
					0,
					0
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
				
				-- Create the Workflow Instance Value records. 
				INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
				SELECT @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElementItems.identifier
				FROM ASRSysWorkflowElementItems 
				INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 2
					AND (ASRSysWorkflowElementItems.itemType = 3 
						OR ASRSysWorkflowElementItems.itemType = 5
						OR ASRSysWorkflowElementItems.itemType = 6
						OR ASRSysWorkflowElementItems.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 13
						OR ASRSysWorkflowElementItems.itemType = 14
						OR ASRSysWorkflowElementItems.itemType = 15
						OR ASRSysWorkflowElementItems.itemType = 17
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 5						
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END';
		



/* --------------------------------------------------------- */
PRINT 'Step - Editable grids Enhancements'
/* --------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysOrderItems', 'U') AND name = 'Editable')
		EXEC sp_executesql N'ALTER TABLE ASRSysOrderItems ADD [Editable] bit NULL';


/* --------------------------------------------------------- */
PRINT 'Step - Cleanup metadata interim build issues'
/* --------------------------------------------------------- */

	EXEC sp_executesql N'UPDATE ASRSysCrossTab SET Selection = 0 WHERE (Selection = 1 AND PicklistID = 0) OR (Selection = 2 AND FilterID = 0);';
	EXEC sp_executesql N'UPDATE ASRSysMailMergeName SET Selection = 0 WHERE (Selection = 1 AND PicklistID = 0) OR (Selection = 2 AND FilterID = 0);';


/* --------------------------------------------------------- */
PRINT 'Step - P&E Core functions'
/* --------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysDurationFromPattern]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsysDurationFromPattern];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsysDurationFromPattern](
		@Absence_In	varchar(5),
		@IndividualDate datetime,
		@SessionType varchar(3),
		@Sunday_Hours_AM numeric(4,2),
		@Monday_Hours_AM numeric(4,2),
		@Tuesday_Hours_AM numeric(4,2),
		@Wednesday_Hours_AM numeric(4,2),
		@Thursday_Hours_AM numeric(4,2),
		@Friday_Hours_AM numeric(4,2),
		@Saturday_Hours_AM numeric(4,2),
		@Sunday_Hours_PM numeric(4,2),
		@Monday_Hours_PM numeric(4,2),
		@Tuesday_Hours_PM numeric(4,2),
		@Wednesday_Hours_PM numeric(4,2),
		@Thursday_Hours_PM numeric(4,2),
		@Friday_Hours_PM numeric(4,2),
		@Saturday_Hours_PM numeric(4,2))
	RETURNS numeric(5,2)
	AS 
	BEGIN

		DECLARE @value numeric(5,2) = 0;

		SET @value = ISNULL(CASE @Absence_In
			WHEN ''Hours'' THEN
				CASE WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = ''AM'' THEN @Sunday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = ''AM'' THEN @Monday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = ''AM'' THEN @Tuesday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = ''AM'' THEN @Wednesday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = ''AM'' THEN @Thursday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = ''AM'' THEN @Friday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = ''AM'' THEN @Saturday_Hours_AM
					WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = ''PM'' THEN @Sunday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = ''PM'' THEN @Monday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = ''PM'' THEN @Tuesday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = ''PM'' THEN @Wednesday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = ''PM'' THEN @Thursday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = ''PM'' THEN @Friday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = ''PM'' THEN @Saturday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 1 AND @SessionType = ''Day'' THEN @Sunday_Hours_AM + @Sunday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 2 AND @SessionType = ''Day'' THEN @Monday_Hours_AM + @Monday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 3 AND @SessionType = ''Day'' THEN @Tuesday_Hours_AM + @Tuesday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 4 AND @SessionType = ''Day'' THEN @Wednesday_Hours_AM + @Wednesday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 5 AND @SessionType = ''Day'' THEN @Thursday_Hours_AM + @Thursday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 6 AND @SessionType = ''Day'' THEN @Friday_Hours_AM + @Friday_Hours_PM
					WHEN DATEPART(dw, @IndividualDate) = 7 AND @SessionType = ''Day'' THEN @Saturday_Hours_AM + @Saturday_Hours_PM
				END
			WHEN ''Days'' THEN
				CASE WHEN @SessionType = ''AM'' OR  @SessionType = ''PM'' THEN 0.5 ELSE 1 END 
			END, 0)

		RETURN @value

	END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysDateRangeToTable]') AND xtype = 'TF')
		DROP FUNCTION [dbo].[udfsysDateRangeToTable];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsysDateRangeToTable]
	(     
		  @Increment              char(1),
		  @StartDate              datetime,
		  @StartSession           char(2),
		  @EndDate                datetime,
		  @EndSession			  char(2)
	)
	RETURNS  
		@SelectedRange	TABLE ([IndividualDate] datetime, [SessionType] char(3))
	AS 
	BEGIN
		SET @StartDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @StartDate));
		SET @EndDate = DATEADD(dd, 0, DATEDIFF(dd, 0, @EndDate));

		WITH cteRange (DateRange) AS (
			SELECT @StartDate
			UNION ALL
			SELECT DATEADD(dd, 0, DATEDIFF(dd, 0, 
					CASE
						WHEN @Increment = ''d'' THEN DATEADD(dd, 1, DateRange)
						WHEN @Increment = ''w'' THEN DATEADD(ww, 1, DateRange)
						WHEN @Increment = ''m'' THEN DATEADD(mm, 1, DateRange)
					END))
			FROM cteRange
			WHERE DateRange <= 
					CASE
						WHEN @Increment = ''d'' THEN DATEADD(dd, -1, @EndDate)
						WHEN @Increment = ''w'' THEN DATEADD(ww, -1, @EndDate)
						WHEN @Increment = ''m'' THEN DATEADD(mm, -1, @EndDate)
					END)         
		INSERT INTO @SelectedRange (IndividualDate, SessionType)
		SELECT DateRange, 
		CASE
			WHEN @StartSession = ''PM'' AND DateRange = @StartDate THEN ''PM''
			WHEN @EndSession = ''AM'' AND DateRange = @EndDate THEN ''AM''
			ELSE ''Day''
		END
		FROM cteRange
		OPTION (MAXRECURSION 3660);
		RETURN;
	END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysPatternFromHours]') AND xtype = 'FN')
		DROP FUNCTION [dbo].[udfsysPatternFromHours];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsysPatternFromHours] (
	@PatternType	varchar(5),
	@Sunday_Hours numeric(4,2),
	@Monday_Hours numeric(4,2),
	@Tuesday_Hours numeric(4,2),
	@Wednesday_Hours numeric(4,2),
	@Thursday_Hours numeric(4,2),
	@Friday_Hours numeric(4,2),
	@Saturday_Hours numeric(4,2))
RETURNS varchar(28)
AS 
BEGIN

	DECLARE @value varchar(28);

	IF @PatternType = ''Days''
		SET @value = CASE WHEN @Sunday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Monday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Tuesday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Wednesday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Thursday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Friday_Hours > 0 THEN ''1'' ELSE ''0'' END +
						CASE WHEN @Saturday_Hours > 0 THEN ''1'' ELSE ''0'' END;
	ELSE
		SET @value = CASE WHEN ISNULL(@Sunday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Sunday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Monday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Monday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Tuesday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Tuesday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Wednesday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Wednesday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Thursday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Thursday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Friday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Friday_Hours), 5), ''.'','''') ELSE ''0000'' END +
						CASE WHEN ISNULL(@Saturday_Hours,0) > 0 THEN REPLACE(RIGHT(''00000'' + CONVERT(varchar(5), @Saturday_Hours), 5), ''.'','''') ELSE ''0000'' END;

	RETURN @value;

END';




/* ------------------------------------------------------- */
PRINT 'Step - Table Triggers'
/* ------------------------------------------------------- */

	EXEC sp_executesql N'DROP VIEW ASRSysTables;';
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_Tables', 'U') AND name = 'InsertTriggerDisabled')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE tbsys_Tables ADD [InsertTriggerDisabled] bit NULL, UpdateTriggerDisabled bit NULL, DeleteTriggerDisabled bit NULL';
		EXEC sp_executesql N'UPDATE [tbsys_Tables] SET InsertTriggerDisabled = 0, UpdateTriggerDisabled = 0, DeleteTriggerDisabled = 0'; 
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('tbsys_Tables', 'U') AND name = 'CopyWhenParentRecordIsCopied')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE tbsys_Tables ADD [CopyWhenParentRecordIsCopied] bit NULL';
		EXEC sp_executesql N'UPDATE [tbsys_Tables] SET [CopyWhenParentRecordIsCopied] = 0'; 
	END

	EXEC sp_executesql N'CREATE VIEW [dbo].[ASRSysTables]
				WITH SCHEMABINDING
				AS SELECT base.[tableid], base.[tabletype], base.[defaultorderid], base.[recorddescexprid], base.[defaultemailid], base.[tablename], base.[manualsummarycolumnbreaks], base.[auditinsert], base.[auditdelete]
						, base.[isremoteview], base.[inserttriggerdisabled], base.[updatetriggerdisabled], base.[deletetriggerdisabled], base.[CopyWhenParentRecordIsCopied]
						, obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
					FROM dbo.[tbsys_tables] base
					INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.tableid AND obj.objecttype = 1
					INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
					WHERE obj.effectivedate <= dt.[date]';

	EXEC sp_executesql N'CREATE TRIGGER [dbo].[INS_ASRSysTables] ON [dbo].[ASRSysTables]
			INSTEAD OF INSERT
			AS
			BEGIN

				SET NOCOUNT ON;

				-- Update objects table
				IF NOT EXISTS(SELECT [guid]
					FROM dbo.[tbsys_scriptedobjects] o
					INNER JOIN inserted i ON i.tableid = o.targetid AND o.objecttype = 1)
				BEGIN
					INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
						SELECT NEWID(), 1, [tableid], dbo.[udfsys_getownerid](), ''01/01/1900'',1,0, GETDATE()
							FROM inserted;
				END

				-- Update base table								
				INSERT dbo.[tbsys_tables] ([TableID], [TableType], [DefaultOrderID], [RecordDescExprID], [DefaultEmailID], [TableName], [ManualSummaryColumnBreaks], [AuditInsert], [AuditDelete], [isremoteview], [inserttriggerdisabled], [updatetriggerdisabled], [deletetriggerdisabled], [CopyWhenParentRecordIsCopied]) 
					SELECT [TableID], [TableType], [DefaultOrderID], [RecordDescExprID], [DefaultEmailID], [TableName], [ManualSummaryColumnBreaks], [AuditInsert], [AuditDelete], [isremoteview], [inserttriggerdisabled], [updatetriggerdisabled], [deletetriggerdisabled], [CopyWhenParentRecordIsCopied] FROM inserted;

			END';

	EXEC sp_executesql N'CREATE TRIGGER [dbo].[DEL_ASRSysTables] ON [dbo].[ASRSysTables]
	INSTEAD OF DELETE
	AS
	BEGIN
		SET NOCOUNT ON;

		DELETE FROM [tbsys_tables] WHERE tableid IN (SELECT tableid FROM deleted);
		DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT tableid FROM deleted) AND objecttype = 1;

	END';

	EXEC sp_executesql N'GRANT SELECT, UPDATE, INSERT, DELETE ON ASRSysTables TO ASRSysGroup;';



	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysTableTriggers]') AND xtype in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE dbo.ASRSysTableTriggers(
			[TriggerID]		integer NOT NULL,
			[TableID]		integer,
			[Name]			nvarchar(255) NOT NULL,
			[Content]		nvarchar(MAX),
			[CodePosition]	integer,
			[IsSystem]		bit NOT NULL
		 CONSTRAINT [PK_ASRSysTableTrigger] PRIMARY KEY CLUSTERED 
			([TriggerID] ASC)) ON [PRIMARY]';
	END


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRCopyChildRecords]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRCopyChildRecords];
	EXECUTE sp_executesql N'CREATE PROCEDURE dbo.spASRCopyChildRecords(
		@iParentTableID integer,
		@iNewRecordID integer,
		@iOriginalRecordID integer)
	WITH EXECUTE AS OWNER
	AS
	BEGIN

		DECLARE @sqlCopyData nvarchar(MAX) = '''';
		DECLARE @childDataColumns TABLE (TableID integer, TableName nvarchar(255), ColumnNames nvarchar(MAX));

		INSERT @childDataColumns (TableID, TableName, ColumnNames)	
			SELECT DISTINCT r.ParentID, t.tablename, d.StringValues
				FROM ASRSysRelations r
				INNER JOIN ASRSysTables t ON t.tableid = r.ChildID
				INNER JOIN ASRSysColumns c ON c.tableid = t.tableid
				CROSS APPLY ( SELECT '', '' + columnname
								FROM ASRSysColumns v2
								WHERE v2.tableid = c.tableid AND v2.datatype <> 4
									FOR XML PATH('''') )  d ( StringValues )
				WHERE r.ParentID = @iParentTableID AND t.CopyWhenParentRecordIsCopied = 1;

		SELECT @sqlCopyData = @sqlCopyData + ''INSERT '' + TableName + ''(ID_'' + CONVERT(varchar(10), TableID) +  ColumnNames + '') SELECT '' 
			+ CONVERT(varchar(10), @iNewRecordID) + ColumnNames 
			+ '' FROM '' + TableName + '' WHERE ID_'' + CONVERT(varchar(10), TableID) + '' = '' + CONVERT(varchar(10), @iOriginalRecordID) + '';'' + CHAR(13)
			FROM @childDataColumns;

		EXECUTE sp_executesql @sqlCopyData;

	END';



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