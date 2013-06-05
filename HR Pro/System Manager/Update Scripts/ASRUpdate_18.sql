/* -------------------------------------------------- */
/* Update the database from version 17 to version 18. */
/* -------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@iDBVersion integer,
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16)


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

/* Exit if the database is not version 17 or 18. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 17) or (@iDBVersion > 18)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ------------------------------ */
/* Create sp_ASRCrossTabsRecDescs */
/* ------------------------------ */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRCrossTabsRecDescs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRCrossTabsRecDescs]

EXEC('CREATE PROCEDURE sp_ASRCrossTabsRecDescs
(@tablename varchar(8000), @recordDescid int)
AS BEGIN

Declare @sSQL nvarchar(4000)

IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = ''sp_ASRExpr_'' + convert(varchar,@RecordDescID))
BEGIN
set @sSQL = ''
	declare @tableid int
	declare @recordid int
	declare @recorddesc varchar(8000)

	DECLARE table_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ID FROM ''+ convert(nvarchar(4000), @tablename) +''

	OPEN table_cursor
	FETCH NEXT FROM table_cursor INTO @recordid

	WHILE (@@fetch_status = 0)
	BEGIN
		exec sp_ASRExpr_'' + convert(nvarchar(4000),@RecordDescID) + '' @RecordDesc OUTPUT, @Recordid
		UPDATE ''+ convert(nvarchar(4000), @tablename) +'' SET RecDesc = @recordDesc WHERE id = @Recordid
		FETCH NEXT FROM table_cursor INTO @recordid
	END
	CLOSE table_cursor
	DEALLOCATE table_cursor''
EXEC sp_executesql @ssql

END

END')

/* ----------------------- */
/* Create sp_ASRIntNewUser */
/* ----------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntNewUser]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntNewUser]

EXEC('CREATE PROCEDURE sp_ASRIntNewUser (
	@psUserName	sysname)
AS
BEGIN
	/* Create an HR Pro user associated with the given SQL login. 
	Put the new user in the current user''s role.
	Return 1 if everything is done okay, else 0. */
	DECLARE @hResult 	integer,
		@sRoleName	sysname

	/* Create a user in the HR Pro database for the given login. */
	EXEC @hResult = sp_grantdbaccess @psUsername, @psUserName
	IF @hResult <> 0 GOTO Done

	/* Determine the current user''s role. */
	SELECT @sRoleName =  a.name
	FROM sysusers a
	INNER JOIN sysusers b 
		ON a.uid = b.gid
	WHERE b.name = CURRENT_USER

	/* Put the new user in the same role as the current user. */
	EXEC @hResult = sp_addrolemember @sRoleName, @psUserName
	IF @hResult <> 0 GOTO Err

	/* Make the new user a dbo. */
	EXEC @hResult = sp_addrolemember ''db_owner'', @psUserName
	IF @hResult <> 0 GOTO Err

	/* Jump over the error handling code. */
	GOTO Done

Err:
	/* Remove the user from the HR Pro database if it was added okay, but not assigned to a role. */
	EXEC sp_revokedbaccess @psUsername

Done:
	RETURN (@hResult)

END')

/* ---------------------------------- */
/* Create sp_ASRIntGetFindWindowTitle */
/* ---------------------------------- */

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIntGetFindWindowTitle]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIntGetFindWindowTitle]

EXEC('CREATE PROCEDURE sp_ASRIntGetFindWindowTitle (@psTitle varchar(100) OUTPUT, @plngScreenID int)
AS
BEGIN
	/* Return the OUTPUT variable @psTitle with the find window title for the given screen. 	*/
	DECLARE @sScreenName	sysname
	
	/* Get the screen name. */
	IF @plngScreenID > 0
	BEGIN
		/* Find title is just the table name. */
		SELECT @psTitle = name
		FROM ASRSysScreens
		WHERE screenID = @plngScreenID

		IF @psTitle IS NULL 
		BEGIN
			SET @psTitle = ''<unknown screen>''
		END
	END
	ELSE
	BEGIN
		SET @psTitle = ''<unknown screen>''
	END	
END')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

UPDATE ASRSysConfig
SET databaseVersion = 18,
	systemManagerVersion = '1.1.16',
	securityManagerVersion = '1.1.16',
	dataManagerVersion = '1.1.16',
	intranetversion = '0.0.5'

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 18 Has Converted Your HR Pro Database To Use V1.1.16 Of HR Pro'
