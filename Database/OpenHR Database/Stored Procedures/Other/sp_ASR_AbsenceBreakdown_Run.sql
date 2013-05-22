CREATE PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run]
(
	@pdReportStart      datetime,
	@pdReportEnd		datetime,
	@pcReportTableName  char(30)
) 
AS 
BEGIN

	SET NOCOUNT ON;

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

	declare @sSQL as varchar(MAX)
	declare @piParentID as integer
	declare @piID as integer
	declare @pbProcessed as bit

	declare @pdTempStartDate as datetime
	declare @pdTempEndDate as datetime
	declare @pcTempStartSession as char(2)
	declare @pcTempEndSession as char(2)
	declare @sTempEndDate as varchar(50)

	declare @pfCount as float
	declare @psVer as char(80)

	/* Alter the structure of the temporary table so it can hold the text for the days */
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ALTER COLUMN Hor NVARCHAR(10)'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ADD Processed BIT'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ADD DisplayOrder INT'
	execute(@sSQL)
	Set @sSQL = 'ALTER TABLE ' + @pcReportTableName + ' ALTER COLUMN Value decimal(10,5)'
	execute(@sSQL)

	/* Load the values from the temporary cursor */
	Set @sSQL = 'DECLARE AbsenceBreakdownCursor CURSOR STATIC FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM ' + @pcReportTableName
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

		/* The absence should only calculate for absence within the reporting period */
		set @pdTempStartDate = @pdStartDate
		set @pcTempStartSession = @pcStartSession
		set @pdTempEndDate = @pdEndDate
		set @pcTempEndSession = @pcEndSession

		--/* If blank leaving date set it to todays date */
		if @pdTempEndDate is Null set @pdTempEndDate = getdate()

		if @pdStartDate <  @pdReportStart
			begin
			set @pdTempStartDate = @pdReportStart
			set @pcTempStartSession = 'AM'
			end
		if @pdTempEndDate >  @pdReportEnd
			begin
			set @pdTempEndDate = @pdReportEnd
			set @pcTempEndSession = 'PM'
			end

		set @sTempEndDate = case when @pdEndDate is null then 'null' else '''' + convert(varchar(40),@pdEndDate) + '''' end

		/* Calculate the days this absence takes up */
		execute sp_ASR_AbsenceBreakdown_Calculate @pfDuration OUTPUT, @pdblMon OUTPUT, @pdblTue OUTPUT, @pdblWed OUTPUT, @pdblThu OUTPUT, @pdblFri OUTPUT, @pdblSat OUTPUT, @pdblSun OUTPUT, @pdTempStartDate, @pcTempStartSession, @pdTempEndDate, @pcTempEndSession, @piParentID

		/* Strip out dodgy characters */
		set @pcRecordDescription = replace(@pcRecordDescription,'''','')
		set @pcType = replace(@pcType,'''','')

		/* Add Mondays records */
		if @pdblMon > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 0) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblMon) + ',''' + convert(varchar(20),@pdStartDate) + ''',1,1,' + @sTempEndDate + ',1)'
			execute(@sSQL)
			end

		/* Add Tuesday records */
		if @pdblTue > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 1) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblTue) +  ',''' + convert(varchar(20),@pdStartDate) + ''',2,1,' + @sTempEndDate +',2)'
			execute(@sSQL)
			end

		/* Add Wednesdays records */
		if @pdblWed > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 2) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblWed) +  ',''' + convert(varchar(20),@pdStartDate) +  ''',3,1,' + @sTempEndDate +',3)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Thursdays were found */
		if @pdblThu > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 3) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblThu) +  ',''' + convert(varchar(20),@pdStartDate) + ''',4,1,' + @sTempEndDate +',4)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Fridays were found */
		if @pdblFri > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 4) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblFri) + ',''' + convert(varchar(20),@pdStartDate) + ''',5,1,' + @sTempEndDate +',5)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Saturdays were found */
		if @pdblSat > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 5) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblSat) + ','''+ convert(varchar(20),@pdStartDate) + ''',6,1,' + @sTempEndDate +',6)'
			execute(@sSQL)
			end

		/* Add new records depending on how many Sundays were found */
		if @pdblSun > 0
			begin
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''' + DATENAME(weekday, 5) + ''',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pdblSun) + ',''' + convert(varchar(20),@pdStartDate) + ''',7,1,' + @sTempEndDate +',0)'
			execute(@sSQL)
			end

		/* Calculate total duraton of absence */
		set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun

		if @pfDuration > 0
			begin
			/* Write records for average, totals and count */
			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Total'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pfDuration) + ',''' + convert(varchar(20),@pdStartDate) + ''',9,1,' + @sTempEndDate +',8)'
			execute(@sSQL)

			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Count'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),1) + ',''' + convert(varchar(20),@pdStartDate) + ''',10,1,' + @sTempEndDate +',10)'
			execute(@sSQL)

			set @sSQL = 'INSERT INTO ' + @pcReportTableName + ' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES (' + Convert(varchar(10),@piParentID) + ',''Average'',''' + @pcType + ''', ''' + @pcRecordDescription + ''', ' + Convert(varchar(10),@pfDuration) + ',''' + convert(varchar(20),@pdStartDate) + ''',9,1,' + @sTempEndDate +',9)'
			execute(@sSQL)
			end

		/* Process next record */
		Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed

		end

	/* Delete this record from our collection as it's now been processed */
	set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Processed IS NULL'
	execute(@sSQL)

	Set @sSQL = 'DECLARE CalculateAverage CURSOR STATIC FOR SELECT Ver,(SUM(Value) / COUNT(Value)) / COUNT(Value) FROM ' + @pcReportTableName + ' WHERE hor = ''Average'' GROUP BY Ver'
	execute(@sSQL)
	open CalculateAverage

	Fetch Next From CalculateAverage Into @psVer, @pfCount
	while @@FETCH_STATUS = 0
		begin
  			Set @sSQL = 'UPDATE ' + @pcReportTableName + ' SET Value = ' + Convert(varchar(10),@pfCount) + ' WHERE Ver =  ''' + @psVer + ''' AND Hor = ''Average'''
		execute(@sSQL)
			Fetch Next From CalculateAverage Into @psVer, @pfCount
		end

	/* Tidy up */
	close AbsenceBreakdownCursor
	close CalculateAverage
	deallocate AbsenceBreakdownCursor
	deallocate CalculateAverage

END