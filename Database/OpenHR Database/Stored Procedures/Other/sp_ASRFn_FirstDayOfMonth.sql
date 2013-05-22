
CREATE PROCEDURE sp_ASRFn_FirstDayOfMonth
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @pdtResult = dateadd(dd, 1 - datepart(dd, @pdtDate), @pdtDate)

END

GO

