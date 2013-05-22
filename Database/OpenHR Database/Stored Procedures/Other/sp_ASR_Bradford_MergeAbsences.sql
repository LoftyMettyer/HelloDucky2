CREATE PROCEDURE [dbo].[sp_ASR_Bradford_MergeAbsences]
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
)
AS
BEGIN
	declare @sSql as varchar(MAX);

	/* Variables to hold current absence record */
	declare @pdStartDate as datetime;
	declare @pdEndDate as datetime;
	declare @pcStartSession as char(2);
	declare @pfDuration as float;
	declare @piID as integer;
	declare @piPersonnelID as integer;
	declare @pbContinuous as bit;

	/* Variables to hold last absence record */
	declare @pdLastStartDate as datetime;
	declare @pcLastStartSession as char(2);
	declare @pfLastDuration as float;
	declare @piLastID as integer;
	declare @piLastPersonnelID as integer;

	/* Open the passed in table */
	set @sSQL = 'DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM ' + @pcReportTableName + ' ORDER BY Personnel_ID, Start_Date ASC';
	execute(@sSQL);
	open BradfordIndexCursor;

	/* Loop through the records in the bradford report table */
	Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID;
	while @@FETCH_STATUS = 0
	begin

		if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
		begin
			Set @pdLastStartDate = @pdStartDate;
			Set @pcLastStartSession = @pcStartSession;
			Set @pfLastDuration = @pfDuration;
			Set @piLastID = @piID;

		end
		else
		begin

			Set @pfLastDuration = @pfLastDuration + @pfDuration;

			/* update start date */
			set @sSQL = 'UPDATE ' + @pcReportTableName + ' SET Start_Date = ''' + convert(varchar(20),@pdLastStartDate) + ''', Start_Session = ''' + @pcLastStartSession + ''', Duration = ' + Convert(Char(10), @pfLastDuration) + ', Included_Days = ' + Convert(Char(10), @pfLastDuration) + ' Where Absence_ID = ' + Convert(varchar(10),@piId);
			execute(@sSQL);

			/* Delete the previous record from our collection */
			set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Absence_ID = ' + Convert(varchar(10),@piLastId);
			execute(@sSQL);

			Set @piLastID = @piID;

		end

		/* Get next absence record */
		Set @piLastPersonnelID = @piPersonnelID;
		
		Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID;
	end

	close BradfordIndexCursor;
	deallocate BradfordIndexCursor;

END