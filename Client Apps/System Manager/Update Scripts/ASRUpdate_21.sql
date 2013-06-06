/* -------------------------------------------------- */
/* Update the database from version 20 to version 21. */
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

/* Exit if the database is not version 20 or 21. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 20) or (@iDBVersion > 21)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ----------------------------- */
/* Dropping ASRSysGlobalItemType */
/* ----------------------------- */

PRINT 'Step 1 of 15 - Dropping ASRSysGlobalItemType'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE sysobjects.name = 'ASRSysGlobalItemType'
	AND sysobjects.xtype = 'U'

IF @iRecCount = 1 
BEGIN
	DROP TABLE ASRSysGlobalItemType
END


/* -------------------------------- */
/* Add Default Display Width Column */
/* -------------------------------- */

PRINT 'Step 2 of 15 - Adding Default Display Width Column'

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


/* ------------------------------ */
/* Updating Default Display Width */
/* ------------------------------ */

PRINT 'Step 3 of 15 - Updating Default Display Widths'


/* VARCHARS */
exec('update asrsyscolumns
set defaultdisplaywidth = asrsyscolumns.[size]
where (datatype = 12 and (columntype = 0 or columntype = 1))')

/* DATES */
exec('update asrsyscolumns
set defaultdisplaywidth = 10
where (datatype = 11)')

/* NUMERICS */
exec('update asrsyscolumns
set defaultdisplaywidth = asrsyscolumns.[size]
where (datatype = 2)')

/* INTEGERS */
exec('update asrsyscolumns
set defaultdisplaywidth = 10
where (datatype = 4 and columntype <> 3)')

exec('UPDATE ASRSysColumns SET [Size] = 10 WHERE (datatype = 4 and columntype <> 3)')

/* LOGICS */
exec('update asrsyscolumns
set defaultdisplaywidth = 1
where (datatype = -7)')

/* OLE */
exec('update asrsyscolumns
set defaultdisplaywidth = 255
where (datatype = -4)')

/* PHOTO */
exec('update asrsyscolumns
set defaultdisplaywidth = 255
where (datatype = -3)')

/* WPATTERN */
exec('update asrsyscolumns
set defaultdisplaywidth = 14
where (datatype = -1)')

exec('UPDATE ASRSysColumns SET [Size] = 14 WHERE datatype = -1')

/* ----------------------- */
/* Adding FilterID Column  */
/* ----------------------- */

PRINT 'Step 4 of 15 - Adding FilterID column to ASRSysExprComponents'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'FilterID'
	AND sysobjects.name = 'ASRSysExprComponents'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExprComponents]
		ADD [FilterID] [int] NULL 

END


/* -------------------------- */
/* Updating Email Links Table */
/* -------------------------- */

PRINT 'Step 5 of 15 - Updating Email Links Table'

UPDATE ASRSysEmailLinks SET IncUserName = 0 WHERE (IncUserName = 1 AND Immediate = 0)

/* ------------------------------------ */
/* Adding Email Attachments Path Column */
/* ------------------------------------ */

PRINT 'Step 6 of 15 - Adding Email Attachments Path Column'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'EmailAttachmentsPath'
	AND sysobjects.name = 'ASRSysConfig'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysConfig]
		ADD [EmailAttachmentsPath] varchar(255) NULL 

END

/* --------------------------- */
/* Adding Expanded Node Column */
/* --------------------------- */

PRINT 'Step 7 of 15 - Expanded Node Column'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'ExpandedNode'
	AND sysobjects.name = 'ASRSysExprComponents'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExprComponents]
		ADD [ExpandedNode] [bit] NULL 

END

/* --------------------------- */
/* Adding Expanded Node Column */
/* --------------------------- */

PRINT 'Step 8 of 15 - Expanded Node Column'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'ExpandedNode'
	AND sysobjects.name = 'ASRSysExpressions'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExpressions]
		ADD [ExpandedNode] [bit] NULL 

END

/* ---------------------------- */
/* Adding View In Colour Column */
/* ---------------------------- */

PRINT 'Step 9 of 15 - View In Colour Column'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'ViewInColour'
	AND sysobjects.name = 'ASRSysExpressions'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysExpressions]
		ADD [ViewInColour] [bit] NULL 

END


/* -------------------------------------- */
/* TableName Changes For ASRSysAuditTrail */
/* -------------------------------------- */

PRINT 'Step 10 of 15 - Tablename Changes For ASRSysAuditTrail'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Tablename'
	AND sysobjects.name = 'ASRSysAuditTrail'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysAuditTrail]
		ADD [Tablename] varchar(200) NULL 

END

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'TableID'
	AND sysobjects.name = 'ASRSysAuditTrail'

IF @iRecCount = 1
BEGIN

	EXEC('UPDATE AsrSysAuditTrail
	SET Tablename = (SELECT AsrSysTables.Tablename FROM AsrSysTables
        WHERE AsrSysTables.TableID = AsrSysAuditTrail.TableID)')

	ALTER TABLE [dbo].[ASRSysAuditTrail]
		DROP COLUMN [TableID]

END


/* --------------------------------------- */
/* ColumnName Changes For ASRSysAuditTrail */
/* --------------------------------------- */

PRINT 'Step 11 of 15 - Columnname Changes For ASRSysAuditTrail'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Columnname'
	AND sysobjects.name = 'ASRSysAuditTrail'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysAuditTrail]
		ADD [Columnname] varchar(200) NULL 

END

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'ColumnID'
	AND sysobjects.name = 'ASRSysAuditTrail'

IF @iRecCount = 1
BEGIN

	EXEC('UPDATE AsrSysAuditTrail
	SET Columnname = (SELECT AsrSysColumns.Columnname FROM AsrSysColumns
        WHERE AsrSysColumns.ColumnID = AsrSysAuditTrail.ColumnID)')

	ALTER TABLE [dbo].[ASRSysAuditTrail]
		DROP COLUMN [ColumnID]

END


/* ------------------ */
/* Update sp_ASRAudit */
/* ------------------ */

PRINT 'Step 12 of 15 - Updating sp_ASRAudit'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAudit]') 
and OBJECTPROPERTY(id,N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAudit]

EXEC('CREATE PROCEDURE sp_ASRAudit (
	@piColumnID int,
	@piRecordID int,
	@psRecordDesc varchar(255),
	@psOldValue varchar(255),
	@psNewValue varchar(255))
AS
BEGIN	

	DECLARE @sTableName varchar(8000)
	DECLARE @sColumnName varchar(8000)

	/* Get the table name for the given column. */
	SELECT @sTableName = tablename 
	FROM asrsystables, asrsyscolumns
	WHERE asrsystables.tableid = asrsyscolumns.tableid
	AND asrsyscolumns.columnid = @piColumnID

	/* Get the column name for the given column. */
	SELECT @sColumnName = columnname
	FROM asrsyscolumns
	WHERE asrsyscolumns.columnid = @piColumnID

	IF @sTableName IS NULL SELECT @sTableName = ''<Unknown>''

	/* Insert a record into the Audit Trail table. */
	INSERT INTO ASRSysAuditTrail 
		(userName, dateTimeStamp, tablename, recordID, recordDesc, columnname, oldValue, newValue)
	VALUES 
		(user, getDate(), @sTableName, @piRecordID, @psRecordDesc, @sColumnName, @psOldValue, @psNewValue)


/*	DECLARE @iTableID int

	Get the table ID for the given column. 
	SELECT @iTableID = tableID 
	FROM ASRSysColumns
	WHERE columnID = @piColumnID

	IF @iTableID IS NULL SELECT @iTableID = 0

	 Insert a record into the Audit Trail table. 
	INSERT INTO ASRSysAuditTrail 
		(userName, dateTimeStamp, tableID, recordID, recordDesc, columnID, oldValue, newValue)
	VALUES 
		(user, getDate(), @iTableID, @piRecordID, @psRecordDesc, @piColumnID, @psOldValue, @psNewValue)
*/

END')

/* -------------------------- */
/* Update sp_ASRGetAuditTrail */
/* -------------------------- */

PRINT 'Step 13 of 15 - Updating sp_ASRGetAuditTrail'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGetAuditTrail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGetAuditTrail]

EXEC('CREATE PROCEDURE sp_ASRGetAuditTrail (
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

END')

/* --------------------- */
/* Add DateFormat Column */
/* --------------------- */

PRINT 'Step 14 of 15 - Adding DateFormat Column'

/* Check if the IntranetVersion column exists. */
SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'DateFormat'
	AND sysobjects.name = 'ASRSysImportName'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysImportName]
		ADD [DateFormat] [varchar] (3)NULL 
END

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 15 of 15 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 21,
	systemManagerVersion = '1.1.19',
	securityManagerVersion = '1.1.19',
	dataManagerVersion = '1.1.19'

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
PRINT 'Update Script 21 Has Converted Your HR Pro Database To Use V1.1.19 Of HR Pro'
