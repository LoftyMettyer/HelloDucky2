
/* -------------------------------------------------- */
/* Update the database from version 32 to version 33. */
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


/* Exit if the database is not version 32 or 33. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.32') and (@sDBVersion <> '1.33')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 11 - Creating SSP Table.'


if not exists (select * from sysobjects
where id = object_id(N'[dbo].[ASRSysSSPRunning]') and OBJECTPROPERTY(id, N'IsTable') = 1)
BEGIN
  CREATE TABLE [dbo].[ASRSysSSPRunning] (
	[personnelRecordID] [int] NOT NULL ,
	[sspRunning] [bit] NOT NULL 
  ) ON [PRIMARY]
END


/* ---------------------------- */

PRINT 'Step 2 of 11 - Creating Utility Triggers.'

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysCustomReportsName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysCustomReportsName]

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysDataTransferName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysDataTransferName]

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysExportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysExportName]

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysGlobalFunctions]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysGlobalFunctions]

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysMailMergeName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysMailMergeName]

if exists (select * from sysobjects
where id = object_id(N'[dbo].[DEL_ASRSysImportName]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[DEL_ASRSysImportName]


--DEL_ASRSysCustomReportsName
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysCustomReportsName ON ASRSysCustomReportsName
FOR DELETE AS
DELETE FROM ASRSysCustomReportsDetails WHERE CustomReportID IN (SELECT ID FROM Deleted)'

exec sp_executesql @NVarCommand


--DEL_ASRSysDataTransferName 
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysDataTransferName ON ASRSysDataTransferName
FOR DELETE AS
DELETE FROM ASRSysDataTransferColumns WHERE DataTransferID IN (SELECT DataTransferID FROM Deleted)'

exec sp_executesql @NVarCommand


--DEL_ASRSysExportName
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysExportName ON ASRSysExportName
FOR DELETE AS
DELETE FROM ASRSysExportDetails WHERE ExportID IN (SELECT ID FROM Deleted)'

exec sp_executesql @NVarCommand


--DEL_ASRSysGlobalFunctions
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysGlobalFunctions ON ASRSysGlobalFunctions
FOR DELETE AS
DELETE FROM ASRSysGlobalItems WHERE FunctionID IN (SELECT FunctionID FROM Deleted)'

exec sp_executesql @NVarCommand


--DEL_ASRSysMailMergeName
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysMailMergeName ON ASRSysMailMergeName
FOR DELETE AS
DELETE FROM ASRSysMailMergeColumns WHERE MailMergeID IN (SELECT MailMergeID FROM Deleted)'

exec sp_executesql @NVarCommand


--DEL_ASRSysImportName
SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysImportName ON ASRSysImportName
FOR DELETE AS
DELETE FROM ASRSysImportDetails WHERE ImportID IN (SELECT ID FROM Deleted)'

exec sp_executesql @NVarCommand



/* ---------------------------- */

PRINT 'Step 3 of 11 - Updating Audit Trail Stored Procedure.'

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

PRINT 'Step 4 of 11 - Amending Cross Tab Definition Table'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysCrossTab')
and name = 'PrintFilterHeader'

if @iRecCount = 0
BEGIN
  ALTER TABLE ASRSysCrossTab ADD PrintFilterHeader bit null

  SELECT @NVarCommand = 'UPDATE ASRSysCrossTab SET PrintFilterHeader = 0'
  EXEC sp_executesql @NVarCommand
END



/* ---------------------------- */

PRINT 'Step 5 of 11 - Updating Modulus Operator Stored Procedure.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASROp_Modulus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASROp_Modulus]


SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASROp_Modulus 
(
	@pdblResult	float OUTPUT,
	@pdblFirst	float,
	@pdblSecond 	float
)
AS
BEGIN
	IF @pdblSecond = 0 
	BEGIN
		SET @pdblResult = 0
	END
	ELSE
	BEGIN
		SET @pdblResult = @pdblFirst - (CAST((@pdblFirst / @pdblSecond) AS INT) * @pdblSecond)
	END
END'

exec sp_executesql @NVarCommand

/* ---------------------------- */

PRINT 'Step 6 of 11 - Amend Export Definition Table.'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'AppendToFile'
	AND sysobjects.name = 'ASRSysExportName'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE ASRSysExportName ADD AppendToFile bit
	ALTER TABLE ASRSysExportName ADD ForceHeader bit
	ALTER TABLE ASRSysExportName ADD OmitHeader bit

	SELECT @NVarCommand = 'UPDATE ASRSysExportName SET AppendToFile = 0, ForceHeader = 0, OmitHeader = 0'
	EXEC sp_executesql @NVarCommand
END


/* ---------------------------- */

PRINT 'Step 7 of 11 - Amending Diary Table.'

ALTER TABLE ASRSysDiaryEvents ALTER COLUMN EventNotes varchar(7000)

/* ---------------------------- */

PRINT 'Step 8 of 11 - Adding Multiple Child Table Reporting Functionality.'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE name = 'ASRSysCustomReportsChildDetails'

IF @iRecCount = 0 
BEGIN
	CREATE TABLE [dbo].[ASRSysCustomReportsChildDetails]
		(
		  [ID] [int] IDENTITY(1, 1) NOT NULL
		, [CustomReportID] [int] NOT NULL
		, [ChildTable] [int] NOT NULL
		, [ChildFilter] [int] NULL
		, [ChildMaxRecords] [int] NULL
		) ON [PRIMARY]

	SELECT @NVarCommand = 'DELETE FROM ASRSysCustomReportsChildDetails
		INSERT INTO ASRSysCustomReportsChildDetails (CustomReportID, ChildTable, ChildFilter, ChildMaxRecords)
			(
			SELECT ID, ChildTable, ChildFilter, ChildMaxRecords
			FROM ASRSysCustomReportsName
			WHERE ChildTable > 0 
			)'

	exec sp_executesql @NVarCommand
END

/* ---------------------------- */

PRINT 'Step 9 of 11 - Adding Repetition Functionality to Custom Reports.'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'Repetition'
	AND sysobjects.name = 'ASRSysCustomReportsDetails'

IF @iRecCount = 0
BEGIN

	SELECT @NVarCommand = 'ALTER TABLE ASRSysCustomReportsDetails ADD Repetition int NULL'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'UPDATE ASRSysCustomReportsDetails SET Repetition = NULL'
	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to -1 where the 'Column' is a child column in the report. 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = -1
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''C'')
						AND (SELECT TableID
			     			     FROM ASRSysColumns
			     			     WHERE ColumnID = B.ColExprID) 
							IN (SELECT ChildTable 
				   			    FROM ASRSysCustomReportsChildDetails 
				    			    WHERE CustomReportID = A.ID)'
	
	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to -1 where the 'Calculation' is a child column(calculation) in the report. 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = -1
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''E'') 
						AND (SELECT TableID 
			     			     FROM ASRSysExpressions 
			     			     WHERE ExprID = B.ColExprID) 
							IN (SELECT ChildTable 
				    			    FROM ASRSysCustomReportsChildDetails 
				    			    WHERE CustomReportID = A.ID)'
	
	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to 1 where the 'Column' is a base or parent column in the report
	-- AND 'Suppress Repeated Values' is OFF for the column. 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = 1
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''C'') 
						AND (B.Srv = 0) 
						AND (SELECT TableID 
						     FROM ASRSysColumns 
						     WHERE ColumnID = B.ColExprID) 
							NOT IN (SELECT ChildTable 
							    	FROM ASRSysCustomReportsChildDetails 
							   	WHERE CustomReportID = A.ID)'

	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to 1 where the 'Calculation' is a base or parent column(calculation) in the report
	-- AND 'Suppress Repeated Values' is OFF for the column(calculation). 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = 1
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''E'') 
						AND (B.Srv = 0) 
						AND (SELECT TableID 
			     			     FROM ASRSysExpressions 
			     			     WHERE ExprID = B.ColExprID) 
							NOT IN (SELECT ChildTable 
				    				FROM ASRSysCustomReportsChildDetails 
	 	    						WHERE CustomReportID = A.ID)'
	
	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to 0 where the 'Column' is a base or parent column in the report
	-- AND 'Suppress Repeated Values' is ON for the column. 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = 0
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''C'') 
						AND (B.Srv = 1) 
						AND (SELECT TableID 
			     			     FROM ASRSysColumns 
			     			     WHERE ColumnID = B.ColExprID) 
							NOT IN (SELECT ChildTable 
				    				FROM ASRSysCustomReportsChildDetails 
				    				WHERE CustomReportID = A.ID)'
	
	EXEC sp_executesql @NVarCommand

	-- Set the Repetition value to 0 where the 'Calculation' is a base or parent column(calculation) in the report
	-- AND 'Suppress Repeated Values' is ON for the column(calculation). 
	SELECT @NVarCommand = 'UPDATE 	ASRSysCustomReportsDetails
				SET 	ASRSysCustomReportsDetails.Repetition = 0
				FROM    ASRSysCustomReportsName A
						INNER JOIN ASRSysCustomReportsDetails B
						ON A.ID = B.CustomReportID
				WHERE 	(B.Type = ''E'') 
						AND (B.Srv = 1) 
						AND (SELECT TableID 
						     FROM ASRSysExpressions 
						     WHERE ExprID = B.ColExprID) 	
							NOT IN (SELECT ChildTable 
							    	FROM ASRSysCustomReportsChildDetails 
							    	WHERE CustomReportID = A.ID)'

	EXEC sp_executesql @NVarCommand

END

/* ---------------------------- */

PRINT 'Step 10 of 11 - Updating Keywords.'

/* Was previously two entries for 'External' so delete both and add one */

DELETE FROM ASRSysKeywords Where Keyword = 'External'
DELETE FROM ASRSysKeywords Where Keyword = 'Function'
DELETE FROM ASRSysKeywords Where Keyword = 'Openxml'
INSERT INTO ASRSysKeywords ([Provider],[Keyword]) VALUES ('Microsoft SQL Server','External')
INSERT INTO ASRSysKeywords ([Provider],[Keyword]) VALUES ('Microsoft SQL Server','Function')
INSERT INTO ASRSysKeywords ([Provider],[Keyword]) VALUES ('Microsoft SQL Server','Openxml')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 11 of 11 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.33')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '1.7.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.33')

update asrsyssystemsettings
set [SettingKey] = 'hrpro@hrpro.co.uk'
where [SettingKey] = 'hrpro@hrpro.com'


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
--Required for SSP stuff
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.33 Of HR Pro'
