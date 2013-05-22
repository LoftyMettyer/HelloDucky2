CREATE PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pbOmitBeforeStart	bit,
	@pbOmitAfterEnd	bit,
	@pcReportTableName	char(30)
)
AS
BEGIN

	SET NOCOUNT ON;

	declare @piID as integer;
	declare @pdStartDate as datetime;
	declare @pdEndDate as datetime;
	declare @iDuration as float;
	declare @pbDeleteThisAbsence as bit;
	declare @sSQL as varchar(MAX);

	set @sSQL = 'DECLARE BradfordIndexCursor CURSOR FOR SELECT Absence_ID, Start_Date, End_Date, Duration FROM ' + @pcReportTableName;
	execute(@sSQL);
	open BradfordIndexCursor;

	Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
	while @@FETCH_STATUS = 0
		begin
			set @pbDeleteThisAbsence = 0;
			if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1;
			if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1;
			if @iDuration = 0 set @pbDeleteThisAbsence = 1;

			if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1;
			if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1;

			if @pbDeleteThisAbsence = 1
				begin
					set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Absence_ID = Convert(Int,' + Convert(char(10),@piId) + ')';
					execute(@sSQL);
				end

			Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
		end

	close BradfordIndexCursor;
	deallocate BradfordIndexCursor;

END
