
CREATE PROCEDURE sp_ASRFn_DayOfDate 
(
	@piResult	integer OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @piResult = datepart(dd, @pdtDate)
END




GO

