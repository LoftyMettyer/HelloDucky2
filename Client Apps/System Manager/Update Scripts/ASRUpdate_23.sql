/* -------------------------------------------------- */
/* Update the database from version 22 to version 23. */
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

/* Exit if the database is not version 22 or 23. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 22) or (@iDBVersion > 23)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ---------------------------- */
/* Create System Settings Table */
/* ---------------------------- */

PRINT 'Step 1 of 10 - Creating System Settings Table'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects
WHERE name = 'ASRSysSystemSettings'

if @iRecCount = 0
BEGIN
	CREATE TABLE [dbo].[ASRSysSystemSettings] (
		[Section] [varchar] (50) NULL ,
		[SettingKey] [varchar] (50) NULL ,
		[SettingValue] [varchar] (50) NULL 
	) ON [PRIMARY]
END

/* ----------------------------- */
/* Create New System Permissions */
/* ----------------------------- */

PRINT 'Step 2 of 10 - Creating New System Permissions'

delete from asrsyspermissionitems where categoryID = 1

insert into asrsyspermissionitems (itemID,description,listOrder,categoryID,itemKey)
values (1,'System Manager',10,1,'SYSTEMMANAGER')

insert into asrsyspermissionitems (itemID,description,listOrder,categoryID,itemKey)
values (3,'Security Manager',20,1,'SECURITYMANAGER')

insert into asrsyspermissionitems (itemID,description,listOrder,categoryID,itemKey)
values (2,'Data Manager',30,1,'DATAMANAGER')

insert into asrsyspermissionitems (itemID,description,listOrder,categoryID,itemKey)
values (4,'Intranet',40,1,'INTRANET')

/* ---------------------------- */
/* Grant New System Permissions */
/* ---------------------------- */

PRINT 'Step 3 of 10 - Granting Data Manager / Intranet (where applicable) Access'

DELETE FROM ASRSysGroupPermissions WHERE itemID = 2 or itemID = 4

DECLARE @SQL varchar(8000)

DECLARE HRProCursor CURSOR
FOR select distinct groupName
from asrsysgrouppermissions
order by groupname

OPEN HRProCursor
FETCH NEXT FROM HRProCursor INTO @GroupName
WHILE @@FETCH_STATUS = 0
BEGIN
	SELECT @SQL = 'INSERT ASRSysGroupPermissions(itemID, groupName, permitted) VALUES(2,'''+@GroupName+''',1)'
	EXECUTE sp_sqlexec @SQL

	SELECT @SQL = 'INSERT ASRSysGroupPermissions(itemID, groupName, permitted) VALUES(4,'''+@GroupName+''',1)'
	EXECUTE sp_sqlexec @SQL

	FETCH NEXT FROM HRProCursor INTO @GroupName
END

CLOSE HRProCursor
DEALLOCATE HRProCursor

/* ---------------------------- */
/* Drop Unused Stored Procedure */
/* ---------------------------- */

PRINT 'Step 4 of 10 - Deleting Unused Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRWhereTransferDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRWhereTransferDetails]

/* -------------------------------------- */
/* Create Bradford Index Stored Procedure */
/* -------------------------------------- */

PRINT 'Step 5 of 10 - Creating Bradford Index Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_BradfordStraddleDays]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_BradfordStraddleDays]

/* --------------------------------- */
/* Create ServerDir Stored Procedure */
/* --------------------------------- */

PRINT 'Step 6 of 10 - Creating Server Directory Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRServerDir]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRServerDir]

exec('CREATE PROCEDURE sp_ASRServerDir(@iMode integer, @sFolder varchar(256)) AS
BEGIN

/* iMode = 1   Return list of available DRIVES (No input necessary)          */
/* iMode = 2   Return list of FOLDERS in specified folder                           */
/* iMode = 3   Return list of FILES in specified folder                                  */


  DECLARE @objectToken integer
  DECLARE @hResult int
  DECLARE @hResult2 int
  DECLARE @pserrormessage varchar(255)
  DECLARE @NumOutput integer
  DECLARE @Index integer
  DECLARE @Output varchar(255)


  /* Create Server DLL object */
  EXEC @hResult = sp_OACreate ''vbpHRProServer.cSrvDir'', @objectToken OUTPUT
  IF @hResult <> 0
  BEGIN
    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
    SET @pserrormessage = ''HR Pro Server.dll not found''
    RAISERROR (@pserrormessage,1,1)
    EXEC sp_OADestroy @objectToken
    RETURN 1
  END


  /* Populate an array within the DLL */
  IF @iMode =1
    EXEC  @hResult = sp_OAMethod @objectToken, ''GetListOfDrives'', @hResult2 OUTPUT
  ELSE
    IF @iMode = 2
      EXEC  @hResult = sp_OAMethod @objectToken, ''GetListOfFolders'', @hResult2 OUTPUT, @sFolder 
    ELSE
        EXEC  @hResult = sp_OAMethod @objectToken, ''GetListOfFiles'', @hResult2 OUTPUT, @sFolder 

  IF @hResult <> 0 
  BEGIN
    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
    SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
    RAISERROR (@pserrormessage,2,1)
    EXEC sp_OADestroy @objectToken
    RETURN 2
  END


  /* Check the UBound of the Array */
  EXEC @hResult = sp_OAGetProperty @objectToken, ''ResultCount'', @NumOutput OUTPUT
  IF @hResult <> 0 
  BEGIN
    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
    SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
    RAISERROR (@pserrormessage,3,1)
    EXEC sp_OADestroy @objectToken
    RETURN 3
  END


  /* Create a temporary table to hold our resultset. */
  CREATE TABLE #ServerOutput (foldername varchar(256))

  SET @Index = 0
  WHILE @Index <= @NumOutput
  BEGIN

    /* Loop though array elements one at a time and put them into a temporary table  */
    EXEC @hResult = sp_OAGetProperty @objectToken, ''Result'', @Output OUTPUT, @Index
    IF @hResult <> 0 
    BEGIN
      EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
      SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
      RAISERROR (@pserrormessage,4,1)
      EXEC sp_OADestroy @objectToken
      RETURN 4
    END

    INSERT #ServerOutput VALUES(@output)
      
    SET @Index = @Index + 1
  END

  EXEC sp_OADestroy @objectToken
  SELECT * FROM #ServerOutput

END')

/* ---------------------------------------- */
/* Create ServerFileExists Stored Procedure */
/* ---------------------------------------- */

PRINT 'Step 7 of 10 - Creating Server File Exists Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRServerFileExists]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRServerFileExists]

EXEC('CREATE PROCEDURE sp_ASRServerFileExists(@sInput varchar(256)) AS
BEGIN

/* Returns 0 if specified FILE EXISTS (otherwises returns 1)   */


  DECLARE @objectToken integer
  DECLARE @hResult int
  DECLARE @hResult2 int
  DECLARE @pserrormessage varchar(255)


  /* Create Server DLL object */
  EXEC @hResult = sp_OACreate ''vbpHRProServer.cSrvDir'', @objectToken OUTPUT
  IF @hResult <> 0
  BEGIN
    EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
    SET @pserrormessage = ''HR Pro Server.dll not found''
    RAISERROR (@pserrormessage,1,1)
    EXEC sp_OADestroy @objectToken
    RETURN 1
  END


  /* Populate an array within the DLL */
  EXEC  @hResult = sp_OAMethod @objectToken, ''CheckIfFileExists'', @hResult2 OUTPUT, @sInput 
  EXEC sp_OADestroy @objectToken

  RETURN @hResult2

END')

/* -------------------------- */
/* Add Intranet Item Category */
/* -------------------------- */

PRINT 'Step 8 of 10 - Adding Intranet Permission Category'

SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 19

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
		VALUES(19,
			'Intranet',
			'',
			10,
			'INTRANET')

	SET IDENTITY_INSERT ASRSysPermissionCategories OFF

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 19

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000000000000680300001600000028000000100000002000000001001800000000004003000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000808080808080808080808080808080808080808080808080808080000000000000000000000000000000000000FF0000808080FFFFFF00FFFFFFFFFFFFFFFFFFFFFF00FFFFFFFFFF808080000000000000000000000000FF0000FF0000FF0000808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFFFFFFFFFFFFFF808080000000000000000000FF0000FF0000FF0000FF0000808080FFFFFF00FFFF800000800000800000808080FFFFFF808080000000000000000000FF0000FF0000FF0000FF0000808080FFFFFF808000FF0000FF0000FF0000800000FFFFFF808080000000000000FF0000FF0000FF0000FF0000FF0000808080FFFFFF808000808080008000FF0000800000FFFFFF808080000000000000FF0000FF0000FF0000008000008000808080FFFFFF808000FFFFFF808080008000800000FFFFFF808080000000000000FF0000FF0000008000008000008000808080FFFFFF00FFFF808000808000808000808080FFFFFF808080000000000000FF0000FF0000008000008000008000808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFF000000000000000000000000000000FF0000FF0000C0C0C0008000008000808080FFFFFF00FFFFFFFFFFFFFFFFFFFFFF808080FFFFFF808080000000000000808080FF0000FF0000FFFFFFC0C0C0808080FFFFFFFFFFFFFFFFFF00FFFFFFFFFF808080808080000000000000000000808080FF0000FFFFFFC0C0C0FFFFFF808080808080808080808080808080808080808080000000000000000000000000000000808080FF0000FF0000C0C0C0FFFFFFC0C0C0008000008000008000008000000000000000000000000000000000000000000000808080808080FF0000FF0000FFFFFFC0C0C0008000000000000000000000000000000000000000000000000000000000000000000000808080808080808080808080808080000000000000000000000000000000000000FC000000F8000000E0000000C0000000800000008000000000000000000000000000000000000000000100008003000080030000C0070000E00F0000F83F000000
END


/* ---------------------- */
/* Add Intranet User Item */
/* -----------------------*/

PRINT 'Step 9 of 10 - Adding Intranet Permission Item'

SELECT @iRecCount = count(*)
FROM ASRSysPermissionItems
WHERE itemID = 82

IF @iRecCount = 0 
BEGIN

	/* The record doesn't exist, so create it. */
	INSERT INTO ASRSysPermissionItems
		(itemID, 
			description, 
			listOrder, 
			categoryID,
			itemKey)
		VALUES(82,
			'New User',
			10,
			19,
			'NEW USER')

END


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 10 of 10 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 23,
	systemManagerVersion = '1.1.21',
	securityManagerVersion = '1.1.21',
	dataManagerVersion = '1.1.21'

/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */

UPDATE ASRSysConfig SET refreshstoredprocedures = 1

/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------------------------------- */
/* Show user message about granting permission */
/* ------------------------------------------- */
print '*****************************************************'
print '* Note : You must manually GRANT EXECUTE permission *'
print '* to DBO for SP_OAGETPROPERTY                       *'
print '*                                                   *'
print '* PLEASE CONTACT YOUR IT ADMINISTRATOR OR THE ASR   *'
print '* HELPDESK FOR HELP ON PERFORMING THIS TASK         *'
print '*****************************************************'

/* --------------------- */
/* INFORM USER ABOUT DLL */
/* --------------------- */

print '*****************************************************'
print '* Note : You must now copy the following file to    *'
print '* the WINDOWS/SYSTEM32 directory, then register it  *'
print '* using REGSRVR32.EXE                               *'
print '*                                                   *'
print '* PLEASE CONTACT YOUR IT ADMINISTRATOR OR THE ASR   *'
print '* HELPDESK FOR HELP ON PERFORMING THIS TASK         *'
print '*****************************************************'

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script 23 Has Converted Your HR Pro Database To Use V1.1.21 Of HR Pro'
