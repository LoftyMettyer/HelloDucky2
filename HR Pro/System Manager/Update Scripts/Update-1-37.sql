
/* -------------------------------------------------- */
/* Update the database from version 35 to version 36. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@sDBVersion varchar(10),
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
        @GroupName varchar(8000),
        @NVarCommand nvarchar(4000),
	@sColumnDataType varchar(8000),
	@iDateFormat varchar(255)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ----------------------------------------------------- */
/* Get the database version from the ASRSysConfig table. */
/* ----------------------------------------------------- */
SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

if @sDBVersion = ''
BEGIN
  SELECT @sDBVersion = SystemManagerVersion FROM ASRSysConfig
END


/* Exit if the database is not version 36 or 37. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.36') and (@sDBVersion <> '1.37')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 9 - Removing Obsolete Stored Procedures'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_BradfordStraddleDays]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_BradfordStraddleDays]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCopyBatchJob]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCopyBatchJob]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCopyDataTransfer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCopyDataTransfer]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCopyGlobal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCopyGlobal]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCopyPicklist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCopyPicklist]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetGlobalFunction]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetGlobalFunction]


/* ---------------------------- */

PRINT 'Step 2 of 9 - Removing Obsolete Columns'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Diaryoffset'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [Diaryoffset]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Diaryperiod'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [Diaryperiod]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Diaryremind'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [Diaryremind]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DiaryComment'
	AND sysobjects.name = 'ASRSysColumns'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysColumns] DROP COLUMN [DiaryComment]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'viewAlternativeName'
	AND sysobjects.name = 'ASRSysViews'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysViews] DROP COLUMN [viewAlternativeName]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'childtable'
	AND sysobjects.name = 'ASRSysCustomReportsName'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysCustomReportsName] DROP COLUMN [childtable]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'childfilter'
	AND sysobjects.name = 'ASRSysCustomReportsName'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysCustomReportsName] DROP COLUMN [childfilter]


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'childmaxrecords'
	AND sysobjects.name = 'ASRSysCustomReportsName'

IF @iRecCount > 0 
	ALTER TABLE [dbo].[ASRSysCustomReportsName] DROP COLUMN [childmaxrecords]


/* ---------------------------- */

PRINT 'Step 3 of 9 - Amending Purge Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRPurgeRecords]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRPurgeRecords]

EXEC('CREATE PROCEDURE dbo.sp_ASRPurgeRecords
(
	@PurgeKey varchar(8000),
	@TableName varchar(8000),
	@DateColumn varchar(8000)
)
AS
BEGIN

	/* EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue'' */

	DECLARE @PurgeDate datetime
	DECLARE @sSQL nvarchar(1000)

	EXEC sp_ASRPurgeDate @PurgeDate OUTPUT, @PurgeKey

	SELECT @sSQL = ''DELETE FROM '' + @TableName + '' WHERE '' + @DateColumn + '' < '''''' + convert(varchar,@PurgeDate,101) + ''''''''
	EXEC sp_executesql @sSQL

END')


/* ---------------------------- */

PRINT 'Step 4 of 9 - Amending Rounding Functions'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_RoundDownToNearestWholeNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_RoundDownToNearestWholeNumber]

EXEC('CREATE PROCEDURE sp_ASRFn_RoundDownToNearestWholeNumber 
(
	@piResult 	integer OUTPUT,	
	@pdblNumber 	float
)
AS
BEGIN
	SET @piResult = round(@pdblNumber, 0, 1)
END')


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_RoundToNearestNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_RoundToNearestNumber]

EXEC('CREATE PROCEDURE sp_ASRFn_RoundToNearestNumber
(
	@pfReturn 		float OUTPUT,
	@pfNumberToRound 	float,
	@pfNearestNumber	float
)
AS
BEGIN

	declare @pfRemainder as float

	/* Calculate the remainder. Cannot use the % because it only works on integers and not floats. */
	set @pfReturn = 0
	set @pfRemainder = @pfNumberToRound - (floor(@pfNumberToRound / @pfNearestNumber) * @pfNearestNumber)

	/* Formula for rounding to the nearest specified number */
	if ((@pfNumberToRound < 0) AND (@pfRemainder <= (@pfNearestNumber / 2.0)))
		OR ((@pfNumberToRound >= 0) AND (@pfRemainder < (@pfNearestNumber / 2.0))) set @pfReturn = @pfNumberToRound - @pfRemainder
		else set @pfReturn = @pfNumberToRound + @pfNearestNumber - @pfRemainder

END')


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_RoundUpToNearestWholeNumber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_RoundUpToNearestWholeNumber]

EXEC('CREATE PROCEDURE sp_ASRFn_RoundUpToNearestWholeNumber 
(
	@piResult 	integer OUTPUT,	
	@pdblNumber 	float
)
AS
BEGIN
	SET @piResult = CASE WHEN @pdblNumber < 0 THEN floor(@pdblNumber)
		ELSE ceiling(@pdblNumber)
		END
END')


/* ---------------------------- */

PRINT 'Step 5 of 9 - Amending Permissions Stored Procedures'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAllTablePermissions]

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAllTablePermissionsForGroup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAllTablePermissionsForGroup]

EXEC('CREATE PROCEDURE [dbo].sp_ASRAllTablePermissions 
AS
BEGIN
	/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
	DECLARE @iUserGroupID	int

	/* Initialise local variables. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = CURRENT_USER

	SELECT sysobjects.name, sysprotects.action
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
		AND sysprotects.protectType <> 206
		AND sysprotects.action <> 193
		AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
	UNION
	SELECT sysobjects.name, 193
	FROM syscolumns
	INNER JOIN sysprotects ON (syscolumns.id = sysprotects.id
		AND sysprotects.action = 193 
		AND sysprotects.uid = @iUserGroupID
		AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
		OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE syscolumns.name = ''timestamp''
		AND ((sysprotects.protectType = 205) 
		OR (sysprotects.protectType = 204))
	ORDER BY sysobjects.name
END')

EXEC('CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]
(
	@psGroupName sysname
)
AS
BEGIN
	/* Return parameters showing what permissions the current user has on all of the HR Pro tables. */
	DECLARE @iUserGroupID	int

	/* Initialise local variables. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = @psGroupName

	SELECT sysobjects.name, sysprotects.action
	FROM sysprotects 
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	WHERE sysprotects.uid = @iUserGroupID
		AND sysprotects.protectType <> 206
		AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
	ORDER BY sysobjects.name
END')


/* ---------------------------- */
PRINT 'Step 6 of 9 - Amending Event Log Purge Trigger'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeEventLog]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[INS_ASRSysPurgeEventLog]

EXEC('CREATE TRIGGER INS_ASRSysPurgeEventLog ON ASRSysEventLog 
FOR INSERT 
AS 
DECLARE @intFrequency int, 
	@strPeriod char(2) 

SELECT @intFrequency = Frequency 
FROM ASRSysEventLogPurge 

SELECT @strPeriod = Period 
FROM ASRSysEventLogPurge 

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL) 
BEGIN 
	IF @strPeriod = ''dd'' 
	BEGIN 
		DELETE FROM ASRSysEventLog 
		WHERE [DateTime] < DATEADD(dd,-@intfrequency,getdate()) 
	END 

	IF @strPeriod = ''wk''
	BEGIN 
		DELETE FROM ASRSysEventLog 
		WHERE [DateTime] < DATEADD(wk,-@intfrequency,getdate()) 
	END 

	IF @strPeriod = ''mm'' 
	BEGIN 
		DELETE FROM ASRSysEventLog 
		WHERE [DateTime] < DATEADD(mm,-@intfrequency,getdate()) 
	END 

	IF @strPeriod = ''yy'' 
	BEGIN 
		DELETE FROM ASRSysEventLog 
		WHERE [DateTime] < DATEADD(yy,-@intfrequency,getdate()) 
	END 

	DELETE FROM ASRSysEventLogDetails 
	WHERE [EventLogID] NOT IN (SELECT ID FROM AsrSysEventLog) 
END')

/* ---------------------------- */

PRINT 'Step 7 of 9 - Updating CMG Stored Procedure'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_GetCMGFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_GetCMGFields]

EXEC('CREATE Procedure sp_ASR_GetCMGFields
(
	@Result datetime OUTPUT,
	@ColumnIDs varchar(8000),
	@RecordID int
)

AS
BEGIN

	DECLARE @sSQL		varchar(8000)

	SET @sSQL = ''SELECT  ASRSysColumns.DataType,  ASRSysAuditTrail.newvalue, ASRSysAuditTrail.DateTimeStamp, ASRSysAuditTrail.DateTimeStamp,ASRSysAuditTrail.ColumnID
	FROM ASRSysAuditTrail
	INNER JOIN ASRSysColumns ON ASRSysAuditTrail.ColumnID = ASRSysColumns.ColumnID
	WHERE ASRSysAuditTrail.ColumnID IN
		('' + @ColumnIDs + '') And ASRSysAuditTrail.RecordID =  Convert(Int,'' + Convert(char(10),@RecordID) + '') 
		And ASRSysAuditTrail.CMGCommitDate IS null Order By ASRSysAuditTrail.ColumnID, ASRSysAuditTrail.DateTimeStamp Desc''

	EXECUTE (@sSQL)
END')



/* ------------------------------------------------------------- */
/*

--This is no longer required as it is done during the save process.


PRINT 'Step 8 of 9 - Grant permission to database objects'

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)

DECLARE curNonDBOGroups CURSOR LOCAL FAST_FORWARD FOR 
SELECT name 
FROM sysusers
INNER JOIN ASRSysGroupPermissions nonSysMgrs ON (sysusers.name = nonSysMgrs.groupName)
INNER JOIN ASRSysPermissionItems nonSysMgrPerms ON nonSysMgrs.itemID = nonSysMgrPerms.itemID
            AND nonSysMgrPerms.categoryID = 1
            AND nonSysMgrPerms.itemKey = 'SYSTEMMANAGER'
            AND nonSysMgrs.permitted = 0
INNER JOIN ASRSysGroupPermissions nonSecMgrs ON (sysusers.name = nonSecMgrs.groupName)
INNER JOIN ASRSysPermissionItems nonSecMgrPerms ON nonSecMgrs.itemID = nonSecMgrPerms.itemID
            AND nonSecMgrPerms.categoryID = 1
            AND nonSecMgrPerms.itemKey = 'SECURITYMANAGER'
            AND nonSecMgrs.permitted = 0
WHERE sysusers.gid = sysusers.uid
            AND sysusers.uid > 0

OPEN curNonDBOGroups
FETCH NEXT FROM curNonDBOGroups INTO @sGroup
WHILE (@@fetch_status = 0)
BEGIN
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR 
            SELECT sysobjects.name, sysobjects.xtype
            FROM sysobjects
            INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
            WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
            OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%')))
                        AND (sysusers.name = 'dbo')

            OPEN curObjects
            FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
            WHILE (@@fetch_status = 0)
            BEGIN
                        IF rtrim(@sObjectType) = 'P'
                        BEGIN
                                    SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [' + @sGroup + ']'
                                    EXEC(@sSQL)
                        END
                        ELSE
                        BEGIN
SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [' + @sGroup + ']'
                        EXEC(@sSQL)
                        END

                        FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
            END

            CLOSE curObjects
            DEALLOCATE curObjects

            FETCH NEXT FROM curNonDBOGroups INTO @sGroup
END

CLOSE curNonDBOGroups
DEALLOCATE curNonDBOGroups
*/

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Set the flag to refresh the stored procedures               */
/* ----------------------------------------------------------- */

PRINT 'Step 9 of 9 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.37')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '1.11.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.37')

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', '1')

/* ------------------------------------------- */
/* Grant permission to email stored procedures */
/* ------------------------------------------- */
SELECT @NVarCommand = 'USE master
GRANT ALL ON master..xp_StartMail TO public
GRANT ALL ON master..xp_SendMail TO public'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE '+@DBName
EXEC sp_executesql @NVarCommand

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
--delete from asrsyssystemsettings
--where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
--insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
--values('database', 'refreshstoredprocedures', 1)


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.37 Of HR Pro'
