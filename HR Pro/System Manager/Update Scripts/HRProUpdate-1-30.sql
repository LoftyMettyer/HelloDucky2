
/* -------------------------------------------------- */
/* Update the database from version 29 to version 30. */
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
        @NVarCommand nvarchar(4000)

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


/* Exit if the database is not version 28 or 29. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.29') and (@sDBVersion <> '1.30')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END



/* ---------------------------- */

PRINT 'Step 1 of 3 - Updating Audit Trail.'

-- Converts the tablename & columnname into a columnID. This is needed for CMG and audit functions to work correctly.
UPDATE ASRSysAuditTrail SET ColumnID = c.ColumnID FROM ASRsysAuditTrail a, ASRsysTables t, ASRsysColumns c
 WHERE t.TableName = a.TableName AND t.TableID = c.TableID AND c.ColumnName = a.ColumnName


/* ---------------------------- */

PRINT 'Step 2 of 3 - Updating CMG Stored Procedure.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_GetCMGFields]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_GetCMGFields]


SELECT @NVarCommand = 'CREATE Procedure sp_ASR_GetCMGFields
(
	@Result datetime OUTPUT,
	@ColumnIDs varchar(8000),
	@RecordID int
)

As

Begin

	declare @sSQL		varchar(8000)

	set @sSQL = ''Select NewValue, DateTimeStamp, ColumnID From ASRSysAuditTrail Where ColumnID in ('' + @ColumnIDs + '') And RecordID =  Convert(Int,'' + Convert(char(10),@RecordID) + '') And CMGCommitDate IS null Order By ColumnID, DateTimeStamp Desc''

	execute (@sSQL)
End'

exec sp_executesql @NVarCommand



/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 3 of 3 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.30')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.30')

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
-- Refresh Not Required for v30...
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.30 Of HR Pro'
