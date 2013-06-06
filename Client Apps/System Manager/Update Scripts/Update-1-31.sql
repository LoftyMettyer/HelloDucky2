
/* -------------------------------------------------- */
/* Update the database from version 30 to version 31. */
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


/* Exit if the database is not version 30 or 31. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.30') and (@sDBVersion <> '1.31')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END



/* ---------------------------- */

PRINT 'Step 1 of 5 - Updating System Settings.'

ALTER TABLE ASRSYSSystemSettings ALTER COLUMN SettingValue varchar(200)


/* ---------------------------- */

PRINT 'Step 2 of 5 - Updating Audit Trail Stored Procedure.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetAuditTrail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetAuditTrail]


SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRGetAuditTrail (
	@piAuditType	int,
	@psOrder 	varchar(200))
AS
BEGIN
	DECLARE @sSQL varchar(2000)

	IF @piAuditType = 1
	BEGIN
/*		SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
			ASRSysAuditTrail.dateTimeStamp AS [Date / Time], 
			ASRSysTables.tableName AS [Table], 
			ASRSysColumns.columnName AS [Column], 
			ASRSysAuditTrail.oldValue AS [Old Value], 
			ASRSysAuditTrail.newValue AS [New Value], 
			ASRSysAuditTrail.recordDesc AS [Record Description],
			ASRSysColumns.columnID AS [ColumnID],
			ASRSysAuditTrail.id
			FROM ASRSysAuditTrail 
			INNER JOIN ASRSysTables ON ASRSysAuditTrail.TableID = ASRSysTables.TableID 
			INNER JOIN ASRSysColumns ON ASRSysAuditTrail.ColumnID = ASRSysColumns.columnID ''
*/
		SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
			ASRSysAuditTrail.dateTimeStamp AS [Date / Time], 
			ASRSysAuditTrail.tableName AS [Table], 
			ASRSysAuditTrail.columnName AS [Column], 
			ASRSysAuditTrail.oldValue AS [Old Value], 
			ASRSysAuditTrail.newValue AS [New Value], 
			ASRSysAuditTrail.recordDesc AS [Record Description],
			ASRSysAuditTrail.id
			FROM ASRSysAuditTrail ''

		IF LEN(@psOrder) >0
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END
	ELSE IF @piAuditType = 2
	BEGIN
		SET @sSQL =  ''SELECT userName AS [User], 
			dateTimeStamp AS [Date / Time],
			groupName AS [User Group],
			viewTableName AS [View / Table],
			columnName AS [Column], 
			action AS [Action],
			permission AS [Permission], 
			id
			FROM ASRSysAuditPermissions ''

		IF LEN(@psOrder) > 0
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END	
	END
	ELSE IF @piAuditType = 3
	BEGIN
		SET @sSQL = ''SELECT userName AS [User],
    			dateTimeStamp AS [Date / Time],
			groupName AS [User Group], 
			userLogin AS [User Login],
			[Action], 
			id
			FROM ASRSysAuditGroup ''

		IF LEN(@psOrder) > 0 
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END
	ELSE IF @piAuditType = 4
	BEGIN
		SET @sSQL = ''SELECT DateTimeStamp AS [Date / Time],
    			UserGroup AS [User Group],
			UserName AS [User], 
			ComputerName AS [Computer Name],
			HRProModule AS [HR Pro Module],
			Action AS [Action], 
			id
			FROM ASRSysAuditAccess ''

		IF LEN(@psOrder) > 0 
		BEGIN
			EXEC (@sSQL + @psOrder)
		END
		ELSE
		BEGIN
			EXEC (@sSQL)
		END
	END

END'

exec sp_executesql @NVarCommand


/* ---------------------------- */

PRINT 'Step 3 of 5 - Updating Audit Trail Purge Stored Procedure.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAuditLogPurge]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAuditLogPurge]


SELECT @NVarCommand = 'CREATE PROCEDURE [sp_ASRAuditLogPurge] AS

DECLARE @intFrequency int,
                  @strPeriod char(2)


SET @strPeriod = null
SET @intFrequency = null

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Users''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Users''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)

BEGIN

  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END

SET @strPeriod = null
SET @intFrequency = null

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Permissions''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Permissions''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)

BEGIN
  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END

SET @strPeriod = null
SET @intFrequency = null

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Data''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Data''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
BEGIN

  IF @strPeriod = ''dd''
  BEGIN
     DELETE FROM AsrSysAuditTrail  WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
   DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END

END

SET @strPeriod = null
SET @intFrequency = null

SELECT @intFrequency = Frequency
FROM AsrSysAuditCleardown
WHERE Type = ''Access''

SELECT @strPeriod = Period
FROM AsrSysAuditCleardown
WHERE Type = ''Access''

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
BEGIN

  IF @strPeriod = ''dd''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = ''wk''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = ''mm''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = ''yy''
  BEGIN
    DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate())
  END
END'

exec sp_executesql @NVarCommand



/* ---------------------------- */

PRINT 'Step 4 of 5 - Updating Overnight Stored Procedure.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_IsOvernightProcess]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_IsOvernightProcess]


SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRFn_IsOvernightProcess
(
    @result integer OUTPUT
)
AS
BEGIN
	DECLARE
		@iCount		integer,
		@sTempExecString	nvarchar(4000),
		@sTempParamDefinition	nvarchar(500),
		@sValue		varchar(8000)

	/* Check if the ''ASRSysSystemSettings'' table exists. */
	SELECT @iCount = count(*)
	FROM sysobjects 
	WHERE name = ''ASRSysSystemSettings''
		
	IF @iCount = 1
	BEGIN
		/* The ASRSysSystemSettings table exists. See if the required records exists in it. */
		SET @sTempExecString = ''SELECT @sValue = settingValue'' +
			'' FROM ASRSysSystemSettings'' +
			'' WHERE section = ''''database'''''' +
			'' AND settingKey = ''''updatingdatedependantcolumns''''''
		SET @sTempParamDefinition = N''@sValue varchar(8000) OUTPUT''
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT
	
		IF NOT @sValue IS NULL
		BEGIN
			SET @result = convert(bit, @sValue)
		END
	END
	ELSE
	BEGIN
		SELECT @result = UpdatingDateDependentColumns FROM ASRSysConfig
	END
END'

exec sp_executesql @NVarCommand


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 5 of 5 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.31')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.31')

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
-- Refresh Not Required for v31...
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.31 Of HR Pro'
