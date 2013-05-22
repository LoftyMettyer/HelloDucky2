CREATE PROCEDURE [dbo].[sp_ASR_Bradford_CalculateDurations]
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
)
AS
BEGIN

	SET NOCOUNT ON;

	declare @pdStartDate as datetime
	declare @pdEndDate as datetime
	declare @pcStartSession as char(2)
	declare @pcEndSession as char(2)

	declare @pfDuration as float
	declare @piID as integer
	declare @pfIncludedAmount as float
	declare @sSQL as varchar(MAX)
	declare @pbIncludedRecalculate as bit
	declare @pbTodaysDate as datetime

	/* Set an end date in the event of a blank one */
	set @pbTodaysDate = getDate()

	/* Open the passed in table */
	set @sSQL = 'DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, End_Date, End_Session, Personnel_ID,Duration FROM ' + @pcReportTableName + ' FOR UPDATE OF Included_Days, Duration'
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
				execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pbTodaysDate, 'PM', @piID
				set @sSQL = 'UPDATE ' + @pcReportTableName + ' SET Duration = ' + convert(char(10), @pfIncludedAmount) + ' WHERE CURRENT OF BradFordIndexCursor'
				execute(@sSQL)
				set @pdEndDate = @pbTodaysDate
				set @pbIncludedRecalculate = 1
			end

		/* Start date is before reporting period */
		if @pdStartDate < @pdReportStart
			begin
				set @pdStartDate = @pdReportStart
				set @pcStartSession = 'AM'
				set @pbIncludedRecalculate = 1
			end

		/* End date is outside the reporting period */
		if @pdEndDate > @pdReportEnd
			begin
				set @pdEndDate = @pdReportEnd
				set @pcEndSession = 'PM'
				set @pbIncludedRecalculate = 1
			end

		/* If outside of report period, recalculate */
		if @pbIncludedRecalculate = 1
			begin
				execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID
				set @sSQL = 'UPDATE ' + @pcReportTableName + ' SET Included_Days = ' + convert(char(10), @pfIncludedAmount) + ' WHERE CURRENT OF BradFordIndexCursor'
				execute(@sSQL)
			end

		/* Get next record */
		fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
	end

	close BradfordIndexCursor
	deallocate BradfordIndexCursor

END