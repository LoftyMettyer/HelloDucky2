/* -------------------------------------------------- */
/* Update the database from version 25 to version 26. */
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

/* Exit if the database is not version 25 or 26. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 25) or (@iDBVersion > 26)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

/* ---------------------------- */

PRINT 'Step 1 of 9 - Amending ASRSysTables'


SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'GrantRead'
	AND sysobjects.name = 'ASRSysTables'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysTables] ADD [GrantRead] bit NULL 
	ALTER TABLE [dbo].[ASRSysTables] ADD [GrantEdit] bit NULL 
	ALTER TABLE [dbo].[ASRSysTables] ADD [GrantNew] bit NULL 
	ALTER TABLE [dbo].[ASRSysTables] ADD [GrantDelete] bit NULL 

END


/* ---------------------------- */

PRINT 'Step 2 of 9 - Amending ASRSysViews'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'GrantRead'
	AND sysobjects.name = 'ASRSysViews'

IF @iRecCount = 0 
BEGIN

	ALTER TABLE [dbo].[ASRSysViews] ADD [GrantRead] bit NULL 
	ALTER TABLE [dbo].[ASRSysViews] ADD [GrantEdit] bit NULL 
	ALTER TABLE [dbo].[ASRSysViews] ADD [GrantNew] bit NULL 
	ALTER TABLE [dbo].[ASRSysViews] ADD [GrantDelete] bit NULL 

END


/* ---------------------------- */

PRINT 'Step 3 of 9 - Amending Diary Links Table'

SELECT @iRecCount = count(syscolumns.id)
FROM syscolumns
INNER JOIN sysobjects
	ON syscolumns.id = sysobjects.id
WHERE syscolumns.name = 'CheckLeavingDate'
	AND sysobjects.name = 'ASRSysDiaryLinks'

IF @iRecCount = 0 
BEGIN
	ALTER TABLE [dbo].[ASRSysDiaryLinks] ADD [CheckLeavingDate] bit NULL

END

EXEC sp_sqlexec 'UPDATE ASRSysDiaryLinks SET CheckLeavingDate = 1 WHERE CheckLeavingDate is null'


/* ---------------------------- */

PRINT 'Step 4 of 9 - Creating Absence Breakdown Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_AbsenceBreakdown_Run]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_AbsenceBreakdown_Run]

exec('CREATE PROCEDURE sp_ASR_AbsenceBreakdown_Run
(
  @pdReportStart      datetime,
  @pdReportEnd    datetime,
  @pcReportTableName  char(30)
) 
AS 
begin
  declare @pdStartDate as datetime
  declare @pdEndDate as datetime
  declare @pcStartSession as char(2)
  declare @pcEndSession as char(2)
  declare @pcType as char(50)
  declare @pcRecordDescription as char(100)

  declare @pfDuration as float
  declare @pdblSun as float
  declare @pdblMon as float
  declare @pdblTue as float
  declare @pdblWed as float
  declare @pdblThu as float
  declare @pdblFri as float
  declare @pdblSat as float

  declare @sSQL as char(8000)
  declare @piParentID as integer
  declare @piID as integer
  declare @pbProcessed as bit

  declare @pdTempStartDate as datetime
  declare @pdTempEndDate as datetime
  declare @pcTempStartSession as char(2)
  declare @pcTempEndSession as char(2)

  /* Alter the structure of the temporary table so it can hold the text for the days */
  Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ALTER COLUMN Hor NVARCHAR(10)''
  execute(@sSQL)
  Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ADD Processed BIT''
  execute(@sSQL)

  /* Load the values from the temporary cursor */
  Set @sSQL = ''DECLARE AbsenceBreakdownCursor CURSOR FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM '' + @pcReportTableName
  execute(@sSQL)
  open AbsenceBreakdownCursor

  /* Loop through the records in the absence breakdown report table */
  Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed
  while @@FETCH_STATUS = 0
    begin

      Set @pdblSun = 0
      Set @pdblMon = 0
      Set @pdblTue = 0
      Set @pdblWed = 0
      Set @pdblThu = 0
      Set @pdblFri = 0
      Set @pdblSat = 0

      /* If blank leaving date set it to todays date */
      if @pdEndDate = Null set @pdEndDate = getdate()

      /* The absence should only calculate for absence within the reporting period */
      set @pdTempStartDate = @pdStartDate
      set @pcTempStartSession = @pcStartSession
      set @pdTempEndDate = @pdEndDate
      set @pcTempEndSession = @pcEndSession

      if @pdStartDate <  @pdReportStart
        begin
          set @pdTempStartDate = @pdReportStart
          set @pcTempStartSession = ''AM''
        end
      if @pdEndDate >  @pdReportEnd
        begin
          set @pdTempEndDate = @pdReportEnd
          set @pcTempEndSession = ''PM''
        end

      /* Calculate the days this absence takes up */
      execute sp_ASR_AbsenceBreakdown_Calculate @pfDuration OUTPUT, @pdblMon OUTPUT, @pdblTue OUTPUT, @pdblWed OUTPUT, @pdblThu OUTPUT, @pdblFri OUTPUT, @pdblSat OUTPUT, @pdblSun OUTPUT, @pdTempStartDate, @pcTempStartSession, @pdTempEndDate, @pcTempEndSession, @piParentID

      /* Strip out dodgy characters */
      set @pcRecordDescription = replace(@pcRecordDescription,'''''''','''')
      set @pcType = replace(@pcType,'''''''','''')

      /* Add Mondays records */
      if @pdblMon > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 0) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblMon) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',1,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add Tuesday records */
      if @pdblTue > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 1) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblTue) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',2,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add Wednesdays records */
      if @pdblWed > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 2) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblWed) +  '','''''' + convert(varchar(20),@pdStartDate) +  '''''',3,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Thursdays were found */
      if @pdblThu > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 3) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblThu) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',4,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Fridays were found */
      if @pdblFri > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 4) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblFri) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',5,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Saturdays were found */
      if @pdblSat > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 5) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblSat) + '',''''''+ convert(varchar(20),@pdStartDate) + '''''',6,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Sundays were found */
      if @pdblSun > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 6) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pdblSun) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',7,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Calculate total duraton of absence */
      set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun

      if @pfDuration > 0
        begin
          /* Write records for totals and count */
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Total'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)

          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Count'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(4),1) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',10,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Process next record */
      Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed

    end

  /* Delete this record from our collection as it''s now been processed */
  set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Processed IS NULL''
  execute(@sSQL)

  /* Tidy up */
  close AbsenceBreakdownCursor
  deallocate AbsenceBreakdownCursor

END')


/* ---------------------------- */

PRINT 'Step 5 of 9 - Deleting Obsolete Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[Sp_ASRFn_AbsenceDurationOLD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sp_ASRFn_AbsenceDurationOLD]


/* ---------------------------- */

PRINT 'Step 6 of 9 - Adding system permissions'

delete from ASRSysPermissionItems
where (categoryID = 1 or itemid in (83,84))

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (1, 'System Manager (Read/Write)', 10, 1, 'SYSTEMMANAGER')

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (83, 'System Manager (Read Only)', 20, 1, 'SYSTEMMANAGERRO')

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (3, 'Security Manager (Read/Write)', 30, 1, 'SECURITYMANAGER')

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (84, 'Security Manager (Read Only)', 40, 1, 'SECURITYMANAGERRO')

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (2, 'Data Manager', 50, 1, 'DATAMANAGER')

insert ASRSysPermissionItems (itemID, description, listOrder, categoryID, itemKey)
values (4, 'Intranet', 60, 1, 'INTRANET')


delete from ASRSysGroupPermissions where itemid in (83,84)

insert ASRSysGroupPermissions (itemid, groupName, permitted)
select 83, groupName, permitted from ASRSysGroupPermissions where itemid = 1

insert ASRSysGroupPermissions (itemid, groupName, permitted)
select 84, groupName, permitted from ASRSysGroupPermissions where itemid = 3

/* ---------------------------- */

PRINT 'Step 7 of 9 - Creating Lock Table'

SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects
WHERE name = 'ASRSysLock'

if @iRecCount = 0
BEGIN
	CREATE TABLE [dbo].[ASRSysLock] (
		[Priority] [int] NULL ,
		[Description] [varchar] (50) NULL ,
		[Username] [varchar] (50) NULL ,
		[Hostname] [varchar] (50) NULL ,
		[Lock_Time] [datetime] NULL ,
		[Login_Time] [datetime] NULL ,
		[SPID] [int] NULL 
	) ON [PRIMARY]
END

/* ---------------------------- */

PRINT 'Step 8 of 9 - Creating Locking Stored Procedures'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockCheck]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockCheck]

exec('CREATE PROCEDURE sp_ASRLockCheck AS
BEGIN

/*
	DELETE FROM ASRSysLock WHERE Priority <> 2 AND Priority in
	(SELECT Priority FROM ASRSysLock
	 left outer join master..sysprocesses syspro 
	 ON asrsyslock.spid = syspro.spid and asrsyslock.login_time = syspro.login_time
	 WHERE syspro.spid IS null)
*/

	SELECT ASRSysLock.* FROM ASRSysLock
	left outer join master..sysprocesses syspro 
	ON asrsyslock.spid = syspro.spid and asrsyslock.login_time = syspro.login_time
	WHERE Priority = 2 or syspro.spid IS not null
	ORDER BY Priority

END')


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockDelete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockDelete]

exec('CREATE Procedure sp_ASRLockDelete (@LockType int)
AS
BEGIN
	DELETE FROM ASRSysLock WHERE Priority = @LockType
END')


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRLockWrite]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRLockWrite]

exec('CREATE Procedure sp_ASRLockWrite (@LockType int)
AS
BEGIN

	DECLARE @LockDesc varchar(50)
	DECLARE @OrigTranCount int

	SELECT @LockDesc = case @LockType
	WHEN 1 THEN ''Saving''
	WHEN 2 THEN ''Manual''
	WHEN 3 THEN ''Read Write''
	ELSE ''''
	END

	IF @LockDesc <> ''''
	BEGIN

		SET @OrigTranCount = @@trancount
		IF @OrigTranCount = 0 BEGIN TRANSACTION

		DELETE FROM ASRSysLock WHERE Priority = @LockType

		INSERT ASRSysLock (Priority, Description, Username, Hostname, Lock_Time, Login_Time, SPID)
		SELECT @LockType, @LockDesc, system_user, host_name(), getdate(), Login_Time, @@spid FROM master..sysprocesses WHERE spid = @@spid

		IF @OrigTranCount = 0 COMMIT TRANSACTION

	END

END')


/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 9 of 9 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 26,
	systemManagerVersion = '1.1.24',
	securityManagerVersion = '1.1.24',
	dataManagerVersion = '1.1.24'

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
PRINT 'Update Script 26 Has Converted Your HR Pro Database To Use V1.1.24 Of HR Pro'
