
CREATE PROCEDURE sp_ASRFn_YearOfDate 
(
	@piResult 	integer OUTPUT,
	@pdtDate	datetime
)
AS
BEGIN
	SET @piResult = datepart(yy, @pdtDate)
END



GO

