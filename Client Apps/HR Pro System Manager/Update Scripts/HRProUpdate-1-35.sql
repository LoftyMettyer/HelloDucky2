
/* -------------------------------------------------- */
/* Update the database from version 34 to version 35. */
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
	@sColumnDataType varchar(8000)

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


/* Exit if the database is not version 34 or 35. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.34') and (@sDBVersion <> '1.35')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 8 - Updating Child View Structure'

if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysChildViewParents2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN

SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysChildViewParents2] (
	[ChildViewID] [int] NOT NULL ,
	[ParentType] [char] (10) NOT NULL ,
	[ParentID] [int] NOT NULL ,
	[ParentTableID] [int] NULL 
) ON [PRIMARY]'
EXEC sp_executesql @NVarCommand


SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysChildViews2] (
	[childViewID] [int] IDENTITY (1, 1) NOT NULL ,
	[tableID] [int] NOT NULL ,
	[type] [int] NULL ,
	[role] [varchar] (256) NOT NULL 
) ON [PRIMARY]'
EXEC sp_executesql @NVarCommand

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRInsertChildView2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRInsertChildView2]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRInsertChildView2 (
	@plngNewRecordID	int OUTPUT,		/* Output variable to hold the new record ID. */
	@plngTableID		int,			/* ID of the table we''re creating a view for. */
	@piType		integer,			/* 0 = OR inter-table join, 1 = AND inter-table join. */
	@psRole		varchar(256))		/* Role name. */
AS
BEGIN
	DECLARE @lngRecordID	int,
		@iCount		int

	SELECT @lngRecordID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @plngTableID
	AND role = @psRole

	IF @lngRecordID IS null
	BEGIN
		/* Insert a record in the ASRSysChildViews table. */
		INSERT INTO ASRSysChildViews2 (tableID, type, role)
		VALUES (@plngTableID, @piType, @psRole)

		/* Get the ID of the inserted record.*/
		SELECT @lngRecordID = MAX(childViewID) FROM ASRSysChildViews2
	END
	ELSE
	BEGIN
		UPDATE ASRSysChildViews2 
		SET type = @piType
		WHERE tableID = @plngTableID
		AND role = @psRole	
	END

	/* Return the new record ID. */
	SET @plngNewRecordID = @lngRecordID
END'
EXEC sp_executesql @NVarCommand


if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRIsSysSecMgr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRIsSysSecMgr]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRIsSysSecMgr (
	@psGroupName		sysname,
	@pfSysSecMgr		bit	OUTPUT
)
AS
BEGIN
	DECLARE @iUserGroupID integer

	/* Get the current user''s group ID. */
	SELECT @iUserGroupID = sysusers.gid
	FROM sysusers
	WHERE sysusers.name = @psGroupName

	SELECT @pfSysSecMgr = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
	FROM ASRSysGroupPermissions
	INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
	INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
	INNER JOIN sysusers ON ASRSysGroupPermissions.groupName = sysusers.name
	WHERE sysusers.uid = @iUserGroupID
		AND (ASRSysPermissionItems.itemKey = ''SYSTEMMANAGER'' OR ASRSysPermissionItems.itemKey = ''SECURITYMANAGER'')
		AND ASRSysGroupPermissions.permitted = 1
		AND ASRSysPermissionCategories.categorykey = ''MODULEACCESS''
END'
EXEC sp_executesql @NVarCommand

END

/* ---------------------------- */

PRINT 'Step 2 of 8 - Updating Absence Breakdown Stored Procedure'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_AbsenceBreakdown_Run]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_AbsenceBreakdown_Run]


EXEC ('CREATE PROCEDURE sp_ASR_AbsenceBreakdown_Run
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
  Set @sSQL = ''DECLARE AbsenceBreakdownCursor CURSOR STATIC FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM '' + @pcReportTableName
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

PRINT 'Step 3 of 8 - Adding Import Date Separator functionality'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysImportName')
and name = 'DateSeparator'

if @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE ASRSysImportName ADD DateSeparator varchar(6) null'
	EXEC sp_executesql @NVarCommand
	
	SELECT @NVarCommand = 'UPDATE ASRSysImportName SET DateSeparator = ''/'''
	EXEC sp_executesql @NVarCommand
END 


/* ---------------------------- */

PRINT 'Step 4 of 8 - Set up overnight job parameters'

SELECT @iRecCount = COUNT(settingvalue)
FROM asrsyssystemsettings
WHERE [section] = 'overnight'

if @iRecCount = 0
BEGIN
	INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	VALUES('overnight','interval','1')

	INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	VALUES('overnight','time','30000')

	INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	VALUES('overnight','type','4')
END


/* ---------------------------- */

PRINT 'Step 5 of 8 - Updating Email Settings'

update asrsyssystemsettings
set [SettingValue] = 'hrpro@hrpro.co.uk'
where lower([SettingValue]) = 'hrpro@hrpro.com'

update asrsyssystemsettings
set [SettingValue] = ''
where lower([SettingValue]) = 'support@yourcompany.com'

/* ---------------------------- */

PRINT 'Step 6 of 8 - Altering the ASRSysGlobalItems.Value column type.'

SELECT @sColumnDataType = LOWER([DATA_TYPE])
FROM [INFORMATION_SCHEMA].[COLUMNS] 
WHERE LOWER([COLUMN_NAME]) = 'value' AND LOWER([TABLE_NAME]) = 'asrsysglobalitems'

if @sColumnDataType != 'varchar'
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE ASRSysGlobalItems ALTER COLUMN [Value] varchar(255) NULL'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'UPDATE ASRSysGlobalItems SET [Value] = RTRIM([Value])'
	EXEC sp_executesql @NVarCommand
END

/* ---------------------------- */

PRINT 'Step 7 of 8 - Removing redundant procedures.'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRDenyAll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRDenyAll]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRGrantAll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRGrantAll]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRInsertChildView]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRInsertChildView]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRPermittedChildView]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRPermittedChildView]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRQuickTableColumnPermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRQuickTableColumnPermissions]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRTableColumnPermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRTableColumnPermissions]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRAbsenceBreakdown]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRAbsenceBreakdown]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRColumnReadPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRColumnReadPermission]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRTablePermissions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRTablePermissions]

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRUserColumnPermission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRUserColumnPermission]

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 8 of 8 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.35')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '1.8.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.35')

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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.35 Of HR Pro'
