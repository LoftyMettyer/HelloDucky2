
/* -------------------------------------------------- */
/* Update the database from version 35 to version 36. */
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
	@sColumnDataType varchar(8000),
	@iDateFormat varchar(255)

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


/* Exit if the database is not version 35 or 36. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '1.35') and (@sDBVersion <> '1.36')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ---------------------------- */

PRINT 'Step 1 of 7 - Removing Obsolete Email Stored Procedures'

DECLARE @SPName varchar(8000)
DECLARE @SQL varchar(8000)

DECLARE HRProCursor CURSOR
FOR select name from sysobjects where name like 'sp_asremail%' or name like 'spasremail%' order by name

set nocount on

OPEN HRProCursor
FETCH NEXT FROM HRProCursor INTO @SPName
WHILE @@FETCH_STATUS = 0
BEGIN
	SELECT @SQL = 'DROP PROCEDURE ' + @SPName
	--PRINT @SQL
	EXECUTE sp_sqlexec @SQL
	FETCH NEXT FROM HRProCursor INTO @SPName
END

CLOSE HRProCursor
DEALLOCATE HRProCursor

set nocount off


/* ---------------------------- */

PRINT 'Step 2 of 7 - Updating Email Queue'

SELECT @iRecCount = count(id) FROM syscolumns
where id = (select id from sysobjects where name = 'ASRSysEmailQueue')
and name = 'Immediate'

if @iRecCount = 0
BEGIN

  exec sp_rename 'asrsysemailqueue.ColumnValue','OLDColumnValue'
  alter table asrsysemailqueue add ColumnValue varchar(255)
  alter table asrsysemailqueue add [Immediate] bit

  select @iDateFormat = '103'  --default
  select @iDateFormat = SettingValue
  from   asrsyssystemsettings
  where  [Section] = 'email' and [SettingKey] = 'date format'

  select @NVarCommand = 'update asrsysemailqueue set ColumnValue = convert(varchar(255),OldColumnValue,' + @iDateFormat + ')'
  exec sp_sqlexec @NVarCommand

  alter table asrsysemailqueue drop column OLDColumnValue

END

/* ---------------------------- */

PRINT 'Step 3 of 7 - Adding New Email Stored Procedures'

EXEC('CREATE PROCEDURE spASREmailImmediate(@Username varchar(255))  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue varchar(8000),
		@RecDescID int,
		@RecDesc nvarchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int,
		@blnEnabled int


	/* Loop through all entries which are to be sent */
	DECLARE emailqueue_cursor
	CURSOR LOCAL FAST_FORWARD FOR 
		SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue
		FROM ASRSysEmailQueue
		WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
		And (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
		ORDER BY DateDue

	OPEN emailqueue_cursor
	FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue

	WHILE (@@fetch_status = 0)
	BEGIN

		SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
					 (SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))

		SET @RecDesc = ''''
		SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @RecDesc OUTPUT, @Recordid
		END


		SELECT @sSQL = ''spASRSysEmailSend_'' + convert(varchar,@LinkID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			SELECT @emailDate = getDate()
		             EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, ''''

			IF @hResult = 0
			BEGIN
				UPDATE ASRSysEmailQueue SET DateSent = @emailDate
				WHERE QueueID = @QueueID
			END
		END

		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue
	END
	CLOSE emailqueue_cursor
	DEALLOCATE emailqueue_cursor

END
')

EXEC('CREATE PROCEDURE spASREmailBatch  AS
BEGIN

	DECLARE @QueueID int,
		@LinkID int,
		@RecordID int,
		@ColumnID int,
		@ColumnValue datetime,
		@RecDescID int,
		@RecDesc nvarchar(4000),
		@sSQL nvarchar(4000),
		@EmailDate datetime,
		@hResult int,
		@blnEnabled int

	/* Clear Servers Inbox */
	/* Doing this just before sending messages means that any failure return messages will */
	/* stay in the servers inbox until this sp is run again - could be useful for support ? */

	SELECT @blnEnabled = SettingValue FROM ASRSysSystemSettings
	WHERE [Section] = ''email'' and [SettingKey] = ''overnight enabled''

	IF @blnEnabled = 0
	BEGIN
		RETURN
	END


	DECLARE @message_id varchar(255)

	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	WHILE not @message_ID is null
	BEGIN
		EXEC master.dbo.xp_deletemail @message_id
		SET @message_id = null
		EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	END


	/* Purge email queue */
	EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue''


	/* Send all emails waiting to be sent regardless of username */
	EXEC spASREmailImmediate ''''

END')


EXEC('CREATE PROCEDURE spASREmailRebuild
AS
BEGIN	
	/* Refresh all calculated columns in the database. */
	DECLARE @sTableName 	varchar(255),
		@iTableID		int,
		@sSQL			varchar(8000),
		@sColumnName		varchar(255)

	
	/* Get a cursor of the tables in the database. */
	DECLARE curTables CURSOR FOR
		SELECT tableName, tableID
		FROM ASRSysTables
	OPEN curTables

	DELETE FROM AsrSysEmailQueue WHERE DateSent Is Null AND [Immediate] = 0

	/* Loop through the tables in the database. */
	FETCH NEXT FROM curTables INTO @sTableName, @iTableID
	WHILE @@fetch_status <> -1
	BEGIN
		/* Get a cursor of the records in the current table.  */
		/* Call the diary trigger for that table and record  */
		SET @sSQL = ''DECLARE @iCurrentID	int,
			@sSQL		varchar(8000)
			DECLARE curRecords CURSOR FOR
				SELECT id
				FROM '' + @sTableName + ''
			OPEN curRecords

			FETCH NEXT FROM curRecords INTO @iCurrentID
			WHILE @@fetch_status <> -1
			BEGIN
				IF EXISTS (SELECT * FROM sysobjects
				WHERE id = object_id(''''spASREmailRebuild_'' + 

LTrim(Str(@iTableID)) + '''''') AND sysstat & 0xf = 4)
				BEGIN
					SET @sSQL = ''''EXEC spASREmailRebuild_'' + 

LTrim(Str(@iTableID)) + '' '''' + convert(varchar(100), @iCurrentID) + ''''''''
					EXEC (@sSQL)
				END
				FETCH NEXT FROM curRecords INTO @iCurrentID
			END
			CLOSE curRecords
			DEALLOCATE curRecords''

		 EXEC (@sSQL) 

		/* Move onto the next table in the database. */ 
		FETCH NEXT FROM curTables INTO @sTableName, @iTableID
	END

	CLOSE curTables
	DEALLOCATE curTables

	EXEC spASREmailImmediate ''''

END')


EXEC('CREATE PROCEDURE spASREmailQueue AS
BEGIN

Declare @sSQL varchar(8000)
declare @queueid int
declare @recordid int
declare @recorddescid int
declare @recorddesc varchar(8000)


DECLARE emailqueue_cursor CURSOR LOCAL FAST_FORWARD FOR 
SELECT QueueID, RecordID, ASRSysTables.RecordDescExprID as RecDescID 
FROM ASRSysEmailQueue
JOIN ASRSysEmailLinks ON ASRSysEmailQueue.LinkID = ASRSysEmailLinks.LinkID
JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysEmailLinks.ColumnID
JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID

OPEN emailqueue_cursor
FETCH NEXT FROM emailqueue_cursor INTO @queueid, @recordid, @recorddescid

WHILE (@@fetch_status = 0)
BEGIN

	SET @RecordDesc = ''''
	SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecordDescID)
	IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
	BEGIN
		EXEC @sSQL @RecordDesc OUTPUT, @Recordid
	END

	UPDATE ASRSysEmailQueue SET RecordDesc = @recordDesc WHERE queueid = @queueid
	FETCH NEXT FROM emailqueue_cursor INTO @queueid, @recordid, @recorddescid
END
CLOSE emailqueue_cursor
DEALLOCATE emailqueue_cursor

END')

/* ---------------------------- */

PRINT 'Step 4 of 7 - Updating Purge Stored Procedures'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRPurgeRecords]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRPurgeRecords]

EXEC('CREATE PROCEDURE dbo.sp_ASRPurgeRecords
(
	@PurgeKey varchar(8000),
	@TableName varchar(8000),
	@DateColumn varchar(8000)
)
AS
BEGIN

	DECLARE @PurgeDate datetime
	DECLARE @sSQL nvarchar(1000)

	EXEC sp_ASRPurgeDate @PurgeDate OUTPUT, @PurgeKey
	SELECT @sSQL = ''DELETE FROM '' + @TableName + '' WHERE Datediff(dd, '' + @DateColumn + '', '''''' + convert(varchar,@PurgeDate,101) + '''''') >= 0''
	EXEC sp_executesql @sSQL

END')


/* ---------------------------- */

PRINT 'Step 5 of 7 - Updating Standard Report procedures'


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_Bradford_MergeAbsences]

EXEC('CREATE PROCEDURE sp_ASR_Bradford_MergeAbsences
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
)
AS
BEGIN
	declare @sSql as char(8000)

	/* Variables to hold current absence record */
	declare @pdStartDate as datetime
	declare @pdEndDate as datetime
	declare @pcStartSession as char(2)
	declare @pfDuration as float
	declare @piID as integer
	declare @piPersonnelID as integer
	declare @pbContinuous as bit

	/* Variables to hold last absence record */
	declare @pdLastStartDate as datetime
	declare @pcLastStartSession as char(2)
	declare @pfLastDuration as float
	declare @piLastID as integer
	declare @piLastPersonnelID as integer

	/* Open the passed in table */
	set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' FOR UPDATE OF Start_Date, Start_Session, Duration,Included_Days''
	execute(@sSQL)
	open BradfordIndexCursor

	/* Loop through the records in the bradford report table */
	Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
	while @@FETCH_STATUS = 0
	begin

		if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
			begin
				Set @pdLastStartDate = @pdStartDate
				Set @pcLastStartSession = @pcStartSession
				Set @pfLastDuration = @pfDuration
				Set @piLastID = @piID

			end
		else
			begin

				Set @pfLastDuration = @pfLastDuration + @pfDuration

				/* update start date */
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(10), @pfLastDuration) + '', Included_Days = '' + Convert(Char(10), @pfLastDuration) + '' WHERE CURRENT OF BradFordIndexCursor''
				execute(@sSQL)

				/* Delete the previous record from our collection */
				set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = '' + Convert(varchar(10),@piLastId)
				execute(@sSQL)

				Set @piLastID = @piID

			end

		/* Get next absence record */
		Set @piLastPersonnelID = @piPersonnelID
		Fetch Next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
	end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor
END')

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_CalculateDurations]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_Bradford_CalculateDurations]

EXEC('CREATE PROCEDURE sp_ASR_Bradford_CalculateDurations
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
)
AS
BEGIN

	declare @pdStartDate as datetime
	declare @pdEndDate as datetime
	declare @pcStartSession as char(2)
	declare @pcEndSession as char(2)

	declare @pfDuration as float
	declare @piID as integer
	declare @pfIncludedAmount as float
	declare @sSQL as char(8000)
	declare @pbIncludedRecalculate as bit
	declare @pbTodaysDate as datetime

	/* Set an end date in the event of a blank one */
	set @pbTodaysDate = getDate()

	/* Open the passed in table */
	set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, End_Date, End_Session, Personnel_ID,Duration FROM '' + @pcReportTableName + '' FOR UPDATE OF Included_Days, Duration''
	execute(@sSQL)
	open BradfordIndexCursor

	/* Loop through the records in the bradford report table */
	fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
	while @@fetch_status = 0
	begin
		/* Calculate start and end dates */
		Set @pbIncludedRecalculate = 0

		/* If empty end date fire off the absence duration calc with system date */
		if isdate(@pdEndDate) = 0
			begin
				execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pbTodaysDate, ''PM'', @piID
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Duration = '' + convert(char(10), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
				execute(@sSQL)
				set @pdEndDate = @pbTodaysDate
				set @pbIncludedRecalculate = 1
			end

		/* Start date is before reporting period */
		if @pdStartDate < @pdReportStart
			begin
				set @pdStartDate = @pdReportStart
				set @pcStartSession = ''AM''
				set @pbIncludedRecalculate = 1
			end

		/* End date is outside the reporting period */
		if @pdEndDate > @pdReportEnd
			begin
				set @pdEndDate = @pdReportEnd
				set @pcEndSession = ''PM''
				set @pbIncludedRecalculate = 1
			end

		/* If outside of report period, recalculate */
		if @pbIncludedRecalculate = 1
			begin
				execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Included_Days = '' + convert(char(10), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
				execute(@sSQL)
			end

		/* Get next record */
		fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
	end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor

END')


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
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 0) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblMon) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',1,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add Tuesday records */
      if @pdblTue > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 1) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblTue) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',2,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add Wednesdays records */
      if @pdblWed > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 2) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblWed) +  '','''''' + convert(varchar(20),@pdStartDate) +  '''''',3,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Thursdays were found */
      if @pdblThu > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 3) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblThu) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',4,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Fridays were found */
      if @pdblFri > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 4) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblFri) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',5,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Saturdays were found */
      if @pdblSat > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 5) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSat) + '',''''''+ convert(varchar(20),@pdStartDate) + '''''',6,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Add new records depending on how many Sundays were found */
      if @pdblSun > 0
        begin
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 6) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSun) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',7,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)
        end

      /* Calculate total duraton of absence */
      set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun

      if @pfDuration > 0
        begin
          /* Write records for totals and count */
          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Total'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
          execute(@sSQL)

          set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Count'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),1) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',10,1,'''''' + convert(varchar(20),@pdEndDate) +'''''')''
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

PRINT 'Step 6 of 7 - Altering Unique SQL Server Object Checking Functionality. Renaming ASRSysTempObjects table to ASRSysSQLObjects.'


SELECT @iRecCount = count(sysobjects.id)
FROM sysobjects 
WHERE lower(name) = 'asrsystempobjects'

IF @iRecCount = 1 
BEGIN
	SELECT @NVarCommand = 'EXEC sp_rename ''ASRSysTempObjects'', ''ASRSysSQLObjects'''
	EXEC sp_executesql @NVarCommand
END

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRUniqueObjectName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRUniqueObjectName]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRUniqueObjectName(
	  				  @psUniqueObjectName sysname OUTPUT
					, @Prefix sysname
					, @Type int)
			AS
			BEGIN
				DECLARE   @NewObj 		as sysname
		 			, @Count 		as int
					, @sUserName		as sysname
					, @sCommandString	nvarchar(4000)	
			 		, @sParamDefinition	nvarchar(500)

				SELECT @sUserName = SYSTEM_USER
				SELECT @Count = 1
				SELECT @NewObj = @Prefix + CONVERT(varchar(100),@Count)

				WHILE (EXISTS (SELECT * FROM sysobjects WHERE id = object_id(@NewObj) AND sysstat & 0xf = @Type))
					OR (EXISTS (SELECT * FROM ASRSysSQLObjects WHERE Name = @NewObj AND Type = @Type))
					BEGIN
						SELECT @Count = @Count + 1
	    					SELECT @NewObj = @Prefix + CONVERT(varchar(10),@Count)
	  				END

				INSERT INTO ASRSysSQLObjects (Name, Type, DateCreated, Owner) VALUES (@NewObj, @Type, GETDATE(), @sUserName)

				SET @sCommandString = ''SELECT @psUniqueObjectName = '''''' + @NewObj + ''''''''

				SET @sParamDefinition = N''@psUniqueObjectName sysname output''

				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psUniqueObjectName output

			END'

EXEC sp_executesql @NVarCommand

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRDropUniqueObject]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRDropUniqueObject]

SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASRDropUniqueObject(
	  			@psUniqueObjectName sysname,
	  			@Type integer)
			AS
			BEGIN
	
				DECLARE 	@sCommandString 	nvarchar(4000),
						@sCurrentUser		varchar(4000),
						@sDBName		varchar(4000)
						
				SET @sCurrentUser = SUSER_SNAME()
				
				SELECT @sDBName = master..sysdatabases.name 
  				FROM master..sysdatabases
	  				INNER JOIN master..sysprocesses ON master..sysdatabases.dbid = master..sysprocesses.dbid
				WHERE master..sysprocesses.spid = @@spid

				IF (EXISTS (SELECT * FROM sysobjects WHERE name = @psUniqueObjectName))
					BEGIN
						IF @Type = 3 
						BEGIN
							SET @sCommandString = ''DROP TABLE ['' + @sCurrentUser + ''].['' + @psUniqueObjectName + '']''
						END

						IF @Type = 4
						BEGIN
							SET @sCommandString = ''DROP PROCEDURE ['' + @sCurrentUser + ''].['' + @psUniqueObjectName + '']''
						END 

						EXECUTE sp_executesql @sCommandString
	  				END
	
				IF (EXISTS (SELECT * FROM ASRSysSQLObjects WHERE Name = @psUniqueObjectName AND Type = @Type AND Owner = @sCurrentUser))
					BEGIN
						SET @sCommandString = ''DELETE FROM ASRSysSQLObjects WHERE Name = '''''' + @psUniqueObjectName 
												                        + '''''' AND Type = '' + convert(varchar(8000), @Type)
													           + '' AND Owner = '''''' + @sCurrentUser + ''''''''

						EXECUTE sp_executesql @sCommandString
	  				END
			END'

EXEC sp_executesql @NVarCommand

/* ---------------------------- */

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 7 of 7 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '1.36')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '1.9.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v1.36')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v1.36 Of HR Pro'
