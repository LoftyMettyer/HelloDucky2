CREATE PROCEDURE sp_ASRFn_DaysBetweenTwoDates 
(
	@piResult	integer OUTPUT,
	@pdtDate1 	datetime,
	@pdtDate2 	datetime
)
AS
BEGIN
	SET @pdtDate1 = convert(datetime, convert(varchar(20), @pdtDate1, 101))
	SET @pdtDate2 = convert(datetime, convert(varchar(20), @pdtDate2, 101))

	/* Get the total number of days difference. */
	SET @piResult = dateDiff(dd, @pdtDate1, @pdtDate2)+1
END
GO

