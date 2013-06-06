/* -------------------------------------------------- */
/* Update the database from version 23 to version 24. */
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

/* Exit if the database is not version 23 or 24. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@iDBVersion < 23) or (@iDBVersion > 24)
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ---------------------------- */

PRINT 'Step 1 of 6 - Amended Service Years SP'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ServiceYears]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ServiceYears]

EXEC('CREATE PROCEDURE sp_ASRFn_ServiceYears 
(
	@piResult		integer OUTPUT,
	@pdtFirstDate 		datetime,
	@pdtSecondDate	datetime
)
AS
BEGIN

	DECLARE @pdtTempDate	datetime

	/* If start date is in the future then return zero */
	IF datediff(d,@pdtFirstDate,getdate()) < 1 or @pdtFirstDate IS null
		SET @piResult = 0
	ELSE
	BEGIN
		IF datediff(d,@pdtSecondDate,getdate()) < 1 or @pdtSecondDate IS null
			/* If leaving date is in the future or blank then calculate from todays date minus start date */
			SET @pdtTempDate = getdate()
		ELSE
			/* If leaving date is in past then calculate from leaving date minus start date */
			SET @pdtTempDate = @pdtSecondDate

		EXEC sp_ASRFn_WholeYearsBetweenTwoDates @piResult OUTPUT, @pdtFirstDate, @pdtTempDate
		IF @piResult < 0 SET @piResult = 0

	END

END')

/* ---------------------------- */

PRINT 'Step 2 of 6 - Amended Service Months SP'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASRFn_ServiceMonths]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASRFn_ServiceMonths]

EXEC('CREATE PROCEDURE sp_ASRFn_ServiceMonths 
(
	@piResult		integer OUTPUT,
	@pdtFirstDate 		datetime,
	@pdtSecondDate	datetime
)
AS
BEGIN

	DECLARE @dtTempDate	datetime

	/* If start date is in the future then return zero */
	IF datediff(d,@pdtFirstDate,getdate()) < 1 or @pdtFirstDate IS null
		SET @piResult = 0
	ELSE
	BEGIN
		IF datediff(d,@pdtSecondDate,getdate()) < 1 or @pdtSecondDate IS null
			/* If leaving date is in the future or blank then calculate from todays date minus start date */
			SET @dtTempDate = getdate()
		ELSE
			/* If leaving date is in past then calculate from leaving date minus start date */
			SET @dtTempDate = @pdtSecondDate

		EXEC sp_ASRFn_WholeMonthsBetweenTwoDates @piResult OUTPUT, @pdtFirstDate, @dtTempDate

		/* NOTE % 12 means divide by 12 and return the remainder */
		/* Remove any whole years from the result */
		SET @piResult = @piResult % 12
		IF @piResult < 0 SET @piResult = 0

	END

END')

/* ---------------------------- */

PRINT 'Step 3 of 6 - Created Bradford Calculate Durations SP'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_CalculateDurations]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
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
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Duration = '' + convert(char(4), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
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
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Included_Days = '' + convert(char(4), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
				execute(@sSQL)
			end

		/* Get next record */
		fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
	end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor

END')

/* ---------------------------- */

PRINT 'Step 4 of 6 - Created Bradford Merge Absences SP'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
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
	set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' FOR UPDATE OF Start_Date, Start_Session, Duration''
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
				/* update start date */
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(4), @pfDuration + @pfLastDuration) + '' WHERE CURRENT OF BradFordIndexCursor''
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

/* ---------------------------- */

PRINT 'Step 5 of 6 - Created Bradford Delete Absences SP'

if exists (select * from sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_DeleteAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ASR_Bradford_DeleteAbsences]

EXEC('CREATE PROCEDURE sp_ASR_Bradford_DeleteAbsences
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pbOmitBeforeStart	bit,
	@pbOmitAfterEnd	bit,
	@pcReportTableName	char(30)
)
AS
BEGIN

	declare @piID as integer
	declare @pdStartDate as datetime
	declare @pdEndDate as datetime
	declare @pbDeleteThisAbsence as bit
	declare @sSQL as char(8000)

	set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Absence_ID, Start_Date, End_Date FROM '' + @pcReportTableName
	execute(@sSQL)
	open BradfordIndexCursor

	Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate
	while @@FETCH_STATUS = 0
		begin
			set @pbDeleteThisAbsence = 0
			if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1
			if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1

			if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1
			if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1

			if @pbDeleteThisAbsence = 1
				begin
					set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = Convert(Int,'' + Convert(char(10),@piId) + '')''
					execute(@sSQL)
				end

			Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate
		end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor

END')

/* ----------------------------------------------------------- */
/* Update the database version flag in the ASRSysConfig table. */
/* Dont Set the flag to refresh the stored procedures          */
/* ----------------------------------------------------------- */

PRINT 'Step 6 of 6 - Updating Versions'

UPDATE ASRSysConfig
SET databaseVersion = 24,
	systemManagerVersion = '1.1.22',
	securityManagerVersion = '1.1.22',
	dataManagerVersion = '1.1.22'

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
PRINT 'Update Script 24 Has Converted Your HR Pro Database To Use V1.1.22 Of HR Pro'
