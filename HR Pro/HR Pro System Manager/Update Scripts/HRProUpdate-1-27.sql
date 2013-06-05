
/* -------------------------------------------------- */
/* Update the database from version 26 to version 27. */
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
        @AuditCommand nvarchar(4000)

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


/* Exit if the database is not version 26 or 27. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.26') and (@sDBVersion <> '1.27')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 4 - Amending Event Log Table'

ALTER TABLE ASRSysEventLog ALTER COLUMN name varchar(100)

/* ---------------------------- */

PRINT 'Step 2 of 4 - Removing obsolete tables'

SELECT @iRecCount = count(sysobjects.id) FROM sysobjects WHERE name = 'ASRSysSecurityDescriptions'
if @iRecCount > 0
BEGIN
  DROP TABLE ASRSysSecurityDescriptions
END

SELECT @iRecCount = count(sysobjects.id) FROM sysobjects WHERE name = 'ASRSysSecurityGroupPrivileges'
if @iRecCount > 0
BEGIN
  DROP TABLE ASRSysSecurityGroupPrivileges
END

SELECT @iRecCount = count(sysobjects.id) FROM sysobjects WHERE name = 'ASRSysSecurityPrivileges'
if @iRecCount > 0
BEGIN
  DROP TABLE ASRSysSecurityPrivileges
END


/* ---------------------------- */

PRINT 'Step 3 of 4 - Amending Export Details Definition Table'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Decimals'
	AND sysobjects.name = 'ASRSysExportDetails'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExportDetails] ADD [Decimals] int NULL 
	EXEC sp_sqlexec 'UPDATE ASRSysExportDetails SET decimals = 0'

END


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 4 of 4 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.27')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.27')

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
/*
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)
*/

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.27 Of HR Pro'
