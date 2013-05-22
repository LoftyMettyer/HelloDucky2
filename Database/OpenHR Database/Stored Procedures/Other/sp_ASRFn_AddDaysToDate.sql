
CREATE PROCEDURE sp_ASRFn_AddDaysToDate 
(
	@pdtResult	datetime OUTPUT,
	@pdtDate 	datetime,
	@piNumber 	integer
)
AS
BEGIN
	SET @pdtResult = dateAdd(dd, @piNumber, @pdtDate)
END





GO

