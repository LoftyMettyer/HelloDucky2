
CREATE PROCEDURE sp_ASRFn_DayOfTheWeek 
(
	@piResult 	integer OUTPUT,
	@pdtDate	datetime
)
AS
BEGIN
	SET @piResult = datepart(dw, @pdtDate)
END



GO

