
CREATE PROCEDURE sp_ASRFn_LastDayOfMonth
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	SET @pdtResult = dateadd(dd, -1, dateadd(mm, 1, dateadd(dd, 1 - datepart(dd, @pdtDate), @pdtDate)))

END

GO

