
CREATE PROCEDURE sp_ASRFn_LastDayOfYear
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @pdtResult = dateadd(dd, -1, dateadd(yy, 1, dateadd(dd, 1 - datepart(dy, @pdtDate), @pdtDate)))
END

GO

