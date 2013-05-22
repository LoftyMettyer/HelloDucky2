
CREATE PROCEDURE sp_ASRFn_AddYearsToDate 
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime,
	@piNumber 	integer
)
AS
BEGIN
	SET @pdtResult = dateAdd(yy, @piNumber, @pdtDate)
END




GO

