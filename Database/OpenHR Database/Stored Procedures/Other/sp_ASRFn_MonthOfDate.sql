
CREATE PROCEDURE sp_ASRFn_MonthOfDate 
(
	@piResult	integer OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @piResult = datepart(mm, @pdtDate)
END




GO

