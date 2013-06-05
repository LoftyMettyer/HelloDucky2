
/* -------------------------------------------------- */
/* Update the database from version 26 to version 27. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @AuditCommand nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */

/* Check if the database version column exists. */

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'databaseVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0
BEGIN
	/* The database version column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [databaseVersion] [int] NULL 
END

/* Check if the refreshStoredProcedures column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'refreshStoredProcedures'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The refreshStoredProcedures column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [refreshStoredProcedures] [bit] NULL 
END

/* Check if the systemManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'systemManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The systemManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SystemManagerVersion] [varchar] (50)NULL 
END

/* Check if the securityManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'securityManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The securityManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [SecurityManagerVersion] [varchar] (50)NULL 
END

/* Check if the DataManagerVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DataManagerVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The DataManagerVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [DataManagerVersion] [varchar] (50)NULL 
END

/* Check if the IntranetVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'IntranetVersion'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN
	/* The IntranetVersion column doesn't exist, so create it. */
	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [IntranetVersion] [varchar] (50)NULL 
END


SET @sCommand = N'SELECT @iDBVersion = databaseVersion
	FROM ASRSysConfig'
SET @sParam = N'@iDBVersion integer OUTPUT'
execute sp_executesql @sCommand, @sParam, @iDBVersion OUTPUT

IF @iDBVersion IS null SET @iDBVersion = 0

/* Exit if the database is not version 25 or 26. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 26) or (@iDBVersion > 27)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

/* ---------------------------- */

PRINT 'Step 1 of 16 - Amending Table Definition Table'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'GrantRead'
	AND sysobjects.name = 'ASRSysTables'

IF @iRecCount > 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysTables] DROP COLUMN [GrantRead]
	ALTER TABLE [dbo].[ASRSysTables] DROP COLUMN [GrantEdit]
	ALTER TABLE [dbo].[ASRSysTables] DROP COLUMN [GrantNew]
	ALTER TABLE [dbo].[ASRSysTables] DROP COLUMN [GrantDelete]

END


/* ---------------------------- */

PRINT 'Step 2 of 16 - Amending View Definition Table'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'GrantRead'
	AND sysobjects.name = 'ASRSysViews'

IF @iRecCount > 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysViews] DROP COLUMN [GrantRead]
	ALTER TABLE [dbo].[ASRSysViews] DROP COLUMN [GrantEdit]
	ALTER TABLE [dbo].[ASRSysViews] DROP COLUMN [GrantNew]
	ALTER TABLE [dbo].[ASRSysViews] DROP COLUMN [GrantDelete]

END


/* ---------------------------- */

PRINT 'Step 3 of 16 - Updating Column Definition Tables'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'UniqueCheckType'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysColumns] ADD [UniqueCheckType] int
END

IF @iRecCount = 0 
BEGIN
	EXEC sp_sqlexec 'UPDATE ASRSysColumns SET uniqueCheckType = 0'
	EXEC sp_sqlexec 'UPDATE ASRSysColumns SET uniqueCheckType = -1 WHERE uniqueCheck = 1'
	EXEC sp_sqlexec 'UPDATE ASRSysColumns SET uniqueCheckType = -2 WHERE childUniqueCheck = 1'
END


/* ---------------------------- */

PRINT 'Step 4 of 16 - Updating Child View Definition Tables'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Type'
	AND sysobjects.name = 'ASRSysChildViews'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysChildViews] ADD [Type] int

END

/* ---------------------------- */

PRINT 'Step 5 of 16 - Amending Mail Merge Definition Tables'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Size'
	AND sysobjects.name = 'ASRSysMailMergeColumns'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysMailMergeColumns] ADD [Size] int NULL 
	ALTER TABLE [dbo].[ASRSysMailMergeColumns] ADD [Decimals] int NULL 
	EXEC sp_sqlexec 'UPDATE ASRSysMailMergeColumns SET size = 0, decimals = 0'

END


/* ---------------------------- */

PRINT 'Step 6 of 16 - Amending Import Definition Tables'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'ImportType'
	AND sysobjects.name = 'ASRSysImportName'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysImportName] ADD [ImportType] int
	EXEC sp_sqlexec 'UPDATE ASRSysImportName SET ImportType = Case CreateNewOnly When 0 Then 0 Else 2 End'
	ALTER TABLE [dbo].[ASRSysImportName] DROP COLUMN [CreateNewOnly]

END


/* ---------------------------- */

PRINT 'Step 7 of 16 - Amending Export Definition Tables'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'CMGExportFileCode'
	AND sysobjects.name = 'ASRSysExportName'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExportName] ADD [CMGExportFileCode] varchar(10) NULL 
	ALTER TABLE [dbo].[ASRSysExportName] ADD [CMGExportUpdateAudit] bit NULL 
	ALTER TABLE [dbo].[ASRSysExportName] ADD [CMGExportRecordID] int NULL 

END


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'CMGColumnCode'
	AND sysobjects.name = 'ASRSysExportDetails'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExportDetails] ADD [CMGColumnCode] varchar(50) NULL 

END


/* ---------------------------- */

PRINT 'Step 8 of 16 - Amending Audit trail'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'CMGExportDate'
	AND sysobjects.name = 'ASRSysAuditTrail'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysAuditTrail] ADD [CMGExportDate] datetime NULL 
	ALTER TABLE [dbo].[ASRSysAuditTrail] ADD [CMGCommitDate] datetime NULL 

END


/* ---------------------------- */

PRINT 'Step 9 of 16 - Amending Expression Components'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'PromptDateType'
	AND sysobjects.name = 'ASRSysExprComponents'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExprComponents] ADD [PromptDateType] int NULL

END


/* ---------------------------- */

PRINT 'Step 10 of 16 - Adding Permission Category'

SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 20

IF @iRecCount = 0 
BEGIN
	SET IDENTITY_INSERT ASRSysPermissionCategories ON

	/* The record doesn't exist, so create it. */
	INSERT INTO ASRSysPermissionCategories
		(categoryID, 
			description, 
			picture, 
			listOrder, 
			categoryKey)
		VALUES(20,
			'CMG Export',
			'',
			10,
			'CMG')

	SET IDENTITY_INSERT ASRSysPermissionCategories OFF

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 20

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000000000000680300001600000028000000100000002000000001001800000000004003000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000808080808080808080808080808080808080808080808080808080000000000000000000000000000000000000FF0000808080FFFFFF00FFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFF808080000000000000000000000000FF0000FF0000FF0000808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFF808080000000000000000000FF0000FF0000FF0000FF0000808080FFFFFF00FFFF800000800000800000808080FFFFFF808080000000000000000000FF0000FF0000FF0000FF0000808080FFFFFF808000FF0000FF0000FF0000800000FFFFFF808080000000000000FF0000FF0000FF0000FF0000FF0000808080FFFFFF808000808080008000FF0000800000FFFFFF808080000000000000FF0000FF0000FF0000008000008000808080FFFFFF808000FFFFFF808080008000800000FFFFFF808080000000000000FF0000FF0000008000008000008000808080FFFFFF00FFFF808000808000808000808080FFFFFF808080000000000000FF0000FF0000008000008000008000808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFF000000000000000000000000000000FF0000FF0000C0C0C0008000008000808080FFFFFF00FFFFFFFFFFFFFFFFFFFFFF808080FFFFFF808080000000000000808080FF0000FF0000FFFFFFC0C0C0808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFF808080808080000000000000000000808080FF0000FFFFFFC0C0C0FFFFFF808080808080808080808080808080808080808080000000000000000000000000000000808080FF0000FF0000C0C0C0FFFFFFC0C0C0008000008000008000008000000000000000000000000000000000000000000000808080808080FF0000FF0000FFFFFFC0C0C0008000000000000000000000000000000000000000000000000000000000000000000000808080808080808080808080808080000000000000000000000000000000000000FC000000F8000000E0000000C0000000800000008000000000000000000000000000000000000000000100008003000080030000C0070000E00F0000F83F000000

END

delete from asrsysPermissionItems where itemid in (85,86,87)
Insert Into asrsysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
values (85,'Recovery',30,20,'CMGRECOVERY')
Insert Into asrsysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
values (86,'Run',10,20,'CMGRUN')
Insert Into asrsysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
values (87,'Commit',20,20,'CMGCOMMIT')


/* ---------------------------- */

PRINT 'Step 11 of 16 - Creating Messaging Table'

if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysMessages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ASRSysMessages] (
	[loginName] [varchar] (256) NULL ,
	[message] [varchar] (200) NULL ,
	[spid] [int] NULL ,
	[dbid] [int] NULL ,
	[uid] [int] NULL ,
	[loginTime] [datetime] NULL ,
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[messageTime] [datetime] NULL ,
	[messageFrom] [varchar] (256) NULL ,
	[messageSource] [varchar] (256) NULL 
) ON [PRIMARY]
END


/* ---------------------------- */

PRINT 'Step 12 of 16 - Creating Messaging Stored Procedures'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRSendMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRSendMessage]

exec('CREATE PROCEDURE sp_ASRSendMessage 
(
	@psMessage	varchar(8000)
)
AS
BEGIN
	DECLARE @iDBid	integer,
		@iSPid		integer,
		@iUid		integer,
		@sLoginName	varchar(256),
		@dtLoginTime	datetime, 
		@sCurrentUser	varchar(256),
		@sCurrentApp	varchar(256)

	/* Get the process information for the current user. */
	SELECT @iDBid = dbid, 
		@sCurrentUser = loginame,
		@sCurrentApp = program_name
	FROM master..sysprocesses
	WHERE spid = @@spid

	/* Get a cursor of the other logged in HR Pro users. */
	DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT spid, loginame, uid, login_time
		FROM master..sysprocesses
		WHERE program_name LIKE ''HR Pro%''
		AND dbid = @iDBid
		AND spid <> @@spid

	OPEN logins_cursor
	FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
	WHILE (@@fetch_status = 0)
	BEGIN
		/* Create a message record for each HR Pro user. */
		INSERT INTO ASRSysMessages 
			(loginname, message, loginTime, dbid, uid, spid, messageTime, messageFrom, messageSource) 
			VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, getdate(), @sCurrentUser, @sCurrentApp)

		FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
	END
	CLOSE logins_cursor
	DEALLOCATE logins_cursor
END')


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetMessages]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetMessages]

exec('CREATE PROCEDURE sp_ASRGetMessages AS
BEGIN
	DECLARE @iDBID		integer,
		@iID		integer,
		@dtLoginTime	datetime,
		@sLoginName	varchar(256),
		@iCount		integer

	/* Get the current user''s process information. */
	SELECT @iDBID = dbID,
		@dtLoginTime = login_time,
		@sLoginName = loginame
	FROM master..sysprocesses
	WHERE spid = @@spid

	/* Return the recordset of messages. */
	SELECT ''Message from user '''''' + ltrim(rtrim(messageFrom)) + 
		'''''' using '' + ltrim(rtrim(messageSource)) + 
		'' ('' + convert(varchar(100), messageTime, 100) +'')'' + 
		char(10) + char(10) + message
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime

	/* Remove any messages that have just been picked up. */
	DELETE
	FROM ASRSysMessages
	WHERE loginName = @sLoginName
		AND spid = @@spid
		AND dbID = @iDBID
		AND loginTime = @dtLoginTime

	/* Remove any orphaned messages. */
	/* NB. This is done via a cursor to avoid any possible collation conflict between ASRSysMessages.loginName and sysprocesses.loginame. */
	DECLARE messages_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT id,
		loginName, 
		dbID, 
		loginTime 
	FROM ASRSysMessages
	OPEN messages_cursor
	FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @iCount = COUNT(*) 
		FROM master..sysprocesses
		WHERE loginame =  @sLoginName
			AND dbID = @iDBID
			AND login_time = @dtLoginTime

		IF @iCount = 0
		BEGIN
			DELETE FROM ASRSysMessages 
			WHERE id = @iID
		END
			
		FETCH NEXT FROM messages_cursor INTO @iID, @sLoginName, @iDBID, @dtLoginTime
	END
	CLOSE messages_cursor 
	DEALLOCATE messages_cursor 

END')

/* ---------------------------- */

PRINT 'Step 13 of 16 - Creating Child View Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRInsertChildView]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRInsertChildView]

exec('CREATE PROCEDURE sp_ASRInsertChildView (
	@plngNewRecordID	int OUTPUT,		/* Output variable to hold the new record ID. */
	@plngTableID		int,			/* ID of the table we''re creating a view for. */
	@piType		integer)			/* 0 = OR inter-table join, 1 = AND inter-table join. */
AS
BEGIN
	DECLARE @lngRecordID	int

	/* Insert a record in the ASRSysChildViews table. */
	INSERT INTO ASRSysChildViews (tableID, type)
	VALUES (@plngTableID, @piType)

	/* Get the ID of the inserted record.*/
	SELECT @lngRecordID = MAX(childViewID) FROM ASRSysChildViews

	/* Return the new record ID. */
	SET @plngNewRecordID = @lngRecordID
END')

/* ---------------------------- */

PRINT 'Step 14 of 16 - Creating Bradford Merge Absences Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_Bradford_MergeAbsences]

exec('CREATE PROCEDURE sp_ASR_Bradford_MergeAbsences
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
)
AS
BEGIN
	declare @sSql as char(8000)

	/* Variables to hold current absence record */
	declare @pdStartDate as datetime
	declare @pdEndDate as datetime
	declare @pcStartSession as char(2)
	declare @pfDuration as float
	declare @piID as integer
	declare @piPersonnelID as integer
	declare @pbContinuous as bit

	/* Variables to hold last absence record */
	declare @pdLastStartDate as datetime
	declare @pcLastStartSession as char(2)
	declare @pfLastDuration as float
	declare @piLastID as integer
	declare @piLastPersonnelID as integer

	/* Open the passed in table */
	set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' FOR UPDATE OF Start_Date, Start_Session, Duration,Included_Days''
	execute(@sSQL)
	open BradfordIndexCursor

	/* Loop through the records in the bradford report table */
	Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
	while @@FETCH_STATUS = 0
	begin

		if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
			begin
				Set @pdLastStartDate = @pdStartDate
				Set @pcLastStartSession = @pcStartSession
				Set @pfLastDuration = @pfDuration
				Set @piLastID = @piID

			end
		else
			begin

				Set @pfLastDuration = @pfLastDuration + @pfDuration

				/* update start date */
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(4), @pfLastDuration) + '', Included_Days = '' + Convert(Char(4), @pfLastDuration) + '' WHERE CURRENT OF BradFordIndexCursor''
				execute(@sSQL)

				/* Delete the previous record from our collection */
				set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = '' + Convert(varchar(10),@piLastId)
				execute(@sSQL)

				Set @piLastID = @piID

			end

		/* Get next absence record */
		Set @piLastPersonnelID = @piPersonnelID
		Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
	end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor

END')

/* ---------------------------- */


PRINT 'Step 15 of 16 - Creating System Settings'

delete from asrsyssystemsettings
where [Section] = 'support'

insert asrsyssystemsettings([Section],[SettingKey],[SettingValue])
values('support','telephone no','01582 714820')

insert asrsyssystemsettings([Section],[SettingKey],[SettingValue])
values('support','fax','01582 714814')

insert asrsyssystemsettings([Section],[SettingKey],[SettingValue])
values('support','email','helpdesk@asr.co.uk')

insert asrsyssystemsettings([Section],[SettingKey],[SettingValue])
values('support','webpage','http://www.asr.co.uk')



delete from asrsyssystemsettings
where [section] = 'email'

insert into asrsyssystemsettings([Section],[SettingKey],[SettingValue])
select 'email', 'date format', isnull(EmailDateFormat,103) from asrsysconfig

insert into asrsyssystemsettings([Section],[SettingKey],[SettingValue])
select 'email', 'attachment path', isnull(EmailAttachmentsPath,'') from asrsysconfig



delete from asrsyssystemsettings
where [section] = 'password'

insert into asrsyssystemsettings([Section],[SettingKey],[SettingValue])
select 'password', 'minimum length', isnull(MinimumPasswordLength,'') from asrsysconfig

insert into asrsyssystemsettings([Section],[SettingKey],[SettingValue])
select 'password', 'change frequency', isnull(ChangePasswordFrequency,'') from asrsysconfig

insert into asrsyssystemsettings([Section],[SettingKey],[SettingValue])
select 'password', 'change period', isnull(ChangePasswordPeriod,'') from asrsysconfig



/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 16 of 16 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 27,
	systemManagerVersion = '1.25',
	securityManagerVersion = '1.25',
	dataManagerVersion = '1.25'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.25')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
UPDATE ASRSysConfig SET refreshstoredprocedures = 1

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 27 Has Converted Your HR Pro Database To Use V1.25 Of HR Pro'
