
CREATE PROCEDURE sp_ASRFn_AddMonthsToDate 
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime,
	@piNumber 	integer
)
AS
BEGIN
	SET @pdtResult = dateAdd(mm, @piNumber, @pdtDate)
END




GO

