/* -------------------------------------------------- */
/* Update the database from version 19 to version 20. */
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

/* Exit if the database is not version 19 or 20. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 19) or (@iDBVersion > 20)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ---------------------------- */
/* Amend sp_ASRSystemPermission */
/* ---------------------------- */

PRINT 'Step 1 of 12 - Amending sp_ASRSystemPermission'

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRSystemPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_ASRSystemPermission]

EXEC('CREATE PROCEDURE sp_ASRSystemPermission
(
	@pfPermissionGranted 	bit OUTPUT,
	@psCategoryKey	varchar(50),
	@psPermissionKey	varchar(50)
)
AS
BEGIN
	/* Return 1 if the given permission is granted to the current user, 0 if it is not.	*/
	DECLARE @fGranted bit


	/* MH20010222 - This needs to be System_User and not Current_User ! */
	SELECT @fGranted = sysAdmin FROM master..syslogins WHERE name = SYSTEM_USER

	IF @fGranted = 0
	BEGIN

		SELECT @fGranted = ASRSysGroupPermissions.permitted
		FROM ASRSysGroupPermissions
			INNER JOIN ASRSysPermissionItems 
				ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
			INNER JOIN ASRSysPermissionCategories
				ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID,
		sysusers a
			INNER JOIN sysusers b 
				ON a.uid = b.gid
		WHERE b.name = CURRENT_USER
			AND ASRSysPermissionItems.itemKey = @psPermissionKey
			AND ASRSysGroupPermissions.groupName = a.name
			AND ASRSysPermissionCategories.categoryKey = @psCategoryKey

	END


	IF @fGranted IS NULL
	BEGIN
		SET @fGranted = 0
	END

	SET @pfPermissionGranted = @fGranted
END')


/* ------------------- */
/* Delete the sys jobs */
/* ------------------- */

PRINT 'Step 2 of 12 - Deleting unused sys jobs'

select @DBName = name 
from master..sysdatabases
where dbid = (select dbid 
       from master..sysprocesses 
       where spid = @@spid)

SET @Command = 'IF EXISTS(SELECT * from msdb..sysjobs where name = ''job_ASRRefreshDateColumns_' + @DBName + ''') BEGIN EXEC sp_sqlexec ''msdb..sp_delete_job @job_name=''''job_ASRRefreshDateColumns_' + @DBName + '''''''' + ' END'
exec sp_sqlexec @Command

SET @Command = 'IF EXISTS(SELECT * from msdb..sysjobs where name = ''job_ASRDiaryProcessing_' + @DBName + ''') BEGIN EXEC sp_sqlexec ''msdb..sp_delete_job @job_name=''''job_ASRDiaryProcessing_' + @DBName + '''''''' + ' END'
exec sp_sqlexec @Command

SET @Command = 'IF EXISTS(SELECT * from msdb..sysjobs where name = ''job_ASREmailProcessing_' + @DBName + ''') BEGIN EXEC sp_sqlexec ''msdb..sp_delete_job @job_name=''''job_ASREmailProcessing_' + @DBName + '''''''' + ' END'
exec sp_sqlexec @Command

/* ---------------------- */
/* Amend sp_AsrDiaryPurge */
/* ---------------------- */

PRINT 'Step 3 of 12 - Amending sp_AsrDiaryPurge'

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_AsrDiaryPurge]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_AsrDiaryPurge]

EXEC('CREATE PROCEDURE sp_AsrDiaryPurge AS

BEGIN

	DECLARE @PurgeDate datetime
	DECLARE @sSQL nvarchar(1000)


    DECLARE @unit char(1),
            @period int,
            @today datetime

    /* Note can''t use sp_ASRPurgeDate as the diary dates include the time !!! */

    select @today = getdate()

    /* Get purge period details */
    select @unit = unit, @period = (period * -1)
    from asrsyspurgeperiods where purgekey =  ''DIARYSYS''

    /* calculate purge date */
    SELECT @purgedate = CASE @unit
        WHEN ''D'' THEN dateadd(dd,@period,@today)
        WHEN ''W'' THEN dateadd(ww,@period,@today)
        WHEN ''M'' THEN dateadd(mm,@period,@today)
        WHEN ''Y'' THEN dateadd(yy,@period,@today)
    END

    SELECT @sSQL = ''DELETE FROM ASRSysDiaryEvents WHERE EventDate < '''''' + convert(varchar,@PurgeDate,101) + '''''' AND ColumnID > 0''

    EXEC sp_executesql @sSQL


    /* Get purge period details */
    select @unit = unit, @period = (period * -1)
    from asrsyspurgeperiods where purgekey =  ''DIARYMAN''

    /* calculate purge date */
    SELECT @purgedate = CASE @unit
        WHEN ''D'' THEN dateadd(dd,@period,@today)
        WHEN ''W'' THEN dateadd(ww,@period,@today)
        WHEN ''M'' THEN dateadd(mm,@period,@today)
        WHEN ''Y'' THEN dateadd(yy,@period,@today)
    END

    SELECT @sSQL = ''DELETE FROM ASRSysDiaryEvents WHERE EventDate < '''''' + convert(varchar,@PurgeDate,101) + '' '' + convert(varchar,@PurgeDate,108) + '''''' AND ColumnID = 0''

    EXEC sp_executesql @sSQL

END')

/* ------------------------------------------------- */
/* Resetting AsrSysColumns size for Photo/Ole fields */
/* ------------------------------------------------- */

PRINT 'Step 4 of 12 - Resetting Size/Decimals In ASRSysColumns Where Unused'

UPDATE ASRSysColumns SET Size = 0 WHERE datatype <> 2 AND datatype <> 12
UPDATE ASRSysColumns SET Decimals = 0 WHERE datatype <> 2

/* -------------------------- */
/* Amending Name Column Sizes */
/* -------------------------- */

PRINT 'Step 5 of 12 - Amending Name Column Size in AsrSysColumns'

EXEC('ALTER TABLE ASRSysColumns ADD TempColumnName VARCHAR(128) NULL')
EXEC('UPDATE ASRSysColumns SET TempColumnName = columnName')
EXEC('ALTER TABLE ASRSysColumns DROP COLUMN columnName')
EXEC('ALTER TABLE ASRSysColumns ADD ColumnName VARCHAR(128) NULL')
EXEC('UPDATE ASRSysColumns SET ColumnName = TempColumnName')
EXEC('ALTER TABLE ASRSysColumns DROP Column TempColumnName')

PRINT 'Step 6 of 12 - Amending Name Column Size in AsrSysTables'

EXEC('ALTER TABLE ASRSysTables ADD TempTableName VARCHAR(128) NULL')
EXEC('UPDATE ASRSysTables SET TempTableName = TableName')
EXEC('ALTER TABLE ASRSysTables DROP COLUMN TableName')
EXEC('ALTER TABLE ASRSysTables ADD TableName VARCHAR(128) NULL')
EXEC('UPDATE ASRSysTables SET TableName = TempTableName')
EXEC('ALTER TABLE ASRSysTables DROP Column TempTableName')

/* -------------------------- */
/* Amending Name Column Sizes */
/* -------------------------- */

PRINT 'Step 7 of 12 - Creating New UserSettings Table'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE name = 'ASRSysUserSettings'

IF @iRecCount = 0 
BEGIN

	CREATE TABLE [dbo].[ASRSysUserSettings] (
		[UserName] [varchar] (50) NULL ,
		[Section] [varchar] (50) NULL ,
		[SettingKey] [varchar] (50) NULL ,
		[SettingValue] [varchar] (50) NULL 
	) ON [PRIMARY]
END


/* ------------------------------------ */
/* Grant/Deny Permissions for Audit Log */
/* ------------------------------------ */

PRINT 'Step 8 of 12 - Denying permissions on System Audit Tables'

DECLARE group_cursor CURSOR FOR 
select groupname
from asrsysgrouppermissions
where asrsysgrouppermissions.itemid = 3
and permitted = 0
and groupname in (select name from sysusers)
order by groupname

OPEN group_cursor
  
FETCH NEXT FROM group_cursor 
INTO @GroupName
  
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @AuditCommand = 'DENY SELECT ON ASRSysAuditAccess TO [' + @GroupName + ']'
	EXEC sp_executesql @AuditCommand

	SET @AuditCommand = 'DENY SELECT ON ASRSysAuditGroup TO [' + @GroupName + ']'
	EXEC sp_executesql @AuditCommand

	SET @AuditCommand = 'DENY SELECT ON ASRSysAuditPermissions TO [' + @GroupName + ']'
	EXEC sp_executesql @AuditCommand

	SET @AuditCommand = 'DENY SELECT ON ASRSysAuditTrail TO [' + @GroupName + ']'
	EXEC sp_executesql @AuditCommand

	FETCH NEXT FROM group_cursor 
	INTO @GroupName
END
  
CLOSE group_cursor

DEALLOCATE group_cursor

/* --------- */
/* SP Change */
/* --------- */

PRINT 'Step 9 of 12 - Updating sp_ASRFn_DaysBetweenTwoDates Procedure'

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_DaysBetweenTwoDates]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_ASRFn_DaysBetweenTwoDates]

EXEC('CREATE PROCEDURE sp_ASRFn_DaysBetweenTwoDates 
(
	@piResult	integer OUTPUT,
	@pdtDate1 	datetime,
	@pdtDate2 	datetime
)
AS
BEGIN
	SET @pdtDate1 = convert(datetime, convert(varchar(20), @pdtDate1, 101))
	SET @pdtDate2 = convert(datetime, convert(varchar(20), @pdtDate2, 101))

	/* Get the total number of days difference. */
	SET @piResult = dateDiff(dd, @pdtDate1, @pdtDate2)+1
END')

/* ----------------------------------------------------------- */
/* Email SP Change                                             */
/* ----------------------------------------------------------- */

PRINT 'Step 10 of 12 - Updating Email Stored Procedure'

IF EXISTS (select * from sysobjects where id = object_id(N'[dbo].[sp_ASREmailBatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[sp_ASREmailBatch]

EXEC('CREATE PROCEDURE sp_ASREmailBatch  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue datetime,
		@RecDescID int,
		@RecDesc nvarchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int

/*		@username varchar(50),
*/

	/* Clear Servers Inbox */
	/* Doing this just before sending messages means that any failure return messages 

will */
	/* stay in the servers inbox until this sp is run again - could be useful for 

support ? */
	DECLARE @message_id varchar(255)

	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	WHILE not @message_ID is null
	BEGIN
		EXEC master.dbo.xp_deletemail @message_id
		SET @message_id = null
		EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	END


	/* Purge email queue */
	EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue''

	/* Loop through all entries which are to be sent */
	DECLARE emailqueue_cursor
	CURSOR LOCAL FAST_FORWARD FOR 
		SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue
		FROM ASRSysEmailQueue
		WHERE DateSent IS Null And DateDue <= GetDate()
		ORDER BY DateDue
	OPEN emailqueue_cursor
	FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue

	WHILE (@@fetch_status = 0)
	BEGIN

		SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
					 (SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))

		SET @RecDesc = ''''
		SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @RecDesc OUTPUT, @Recordid
		END


		SELECT @sSQL = ''sp_ASREmailSend_'' + convert(varchar,@LinkID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			SELECT @emailDate = getDate()
		             EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''

			IF @hResult = 0
			BEGIN
				UPDATE ASRSysEmailQueue SET DateSent = @emailDate
				WHERE QueueID = @QueueID
			END
		END

		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue
	END
	CLOSE emailqueue_cursor
	DEALLOCATE emailqueue_cursor

END')


/* ----------------------------------------------------------- */
/* Create DefaultDisplayWidth Col                              */
/* ----------------------------------------------------------- */

PRINT 'Step 11 of 12 - Updating Email Stored Procedure'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DefaultDisplayWidth'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysColumns]
		ADD [DefaultDisplayWidth] [int] NULL 

	CONSTRAINT [DF_ASRSysColumns_DefaultDisplayWidth] DEFAULT (1)

END

EXEC ('UPDATE ASRSysColumns SET DefaultDisplayWidth = 1')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 12 of 12 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 20,
	systemManagerVersion = '1.1.18',
	securityManagerVersion = '1.1.18',
	dataManagerVersion = '1.1.18'

/*,
	intranetversion = '0.0.6'
*/


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */

UPDATE ASRSysConfig SET refreshstoredprocedures = 1

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 20 Has Converted Your HR Pro Database To Use V1.1.18 Of HR Pro'
