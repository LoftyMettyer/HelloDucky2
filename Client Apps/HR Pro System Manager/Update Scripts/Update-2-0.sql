
/* ---------------------------------------------------- */
/* Update the database from version 1.37 to version 2.9 */
/* ----------------------------------------------------	*/

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
	@iDateFormat varchar(255),
	@sSQLVersion nvarchar(20)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 1.37 or 2.9. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.37') 
and (@sDBVersion <> '2.0') 
and (@sDBVersion <> '2.1') 
and (@sDBVersion <> '2.2') 
and (@sDBVersion <> '2.3') 
and (@sDBVersion <> '2.4') 
and (@sDBVersion <> '2.5') 
and (@sDBVersion <> '2.6') 
and (@sDBVersion <> '2.7') 
and (@sDBVersion <> '2.8') 
and (@sDBVersion <> '2.9')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */

PRINT 'Step 1 of 1 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.9')

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.9 Of HR Pro'