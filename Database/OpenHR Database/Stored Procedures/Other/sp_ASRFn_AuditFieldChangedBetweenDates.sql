CREATE PROCEDURE [dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]
(
	@Result		bit OUTPUT,
	@ColumnID	integer,
	@FromDate	datetime,
	@ToDate		datetime,
	@RecordID	integer
)
AS
BEGIN
	declare @Found as integer;

	set @Result = 0;
		
	set @Found = (SELECT Count(DateTimeStamp) FROM [dbo].[ASRSysAuditTrail]
					WHERE ColumnID = @ColumnID
           				AND RecordID = @RecordID
						AND DateTimeStamp >= @FromDate AND DateTimeStamp <= @ToDate+1);

	if @found > 0 set @Result = 1;

END