
CREATE PROCEDURE sp_ASRFn_FirstDayOfYear
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @pdtResult = dateadd(dd, 1 - datepart(dy, @pdtDate), @pdtDate)
END

GO

