
CREATE PROCEDURE sp_ASRFn_RemainingMonthsSinceWholeYears 
(
	@piResult	integer OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	DECLARE @dtToday	datetime

	SET @dtToday = getdate()
	SET @pdtDate = convert(datetime, convert(varchar(20), @pdtDate, 101))

	/* Get the number of whole months */
	SET @piResult = month(@dtToday) - month(@pdtDate)
 
	/* Test the day value */
	IF day(@pdtDate) > day(@dtToday)
	BEGIN
		SET @piResult = @piResult - 1
	END

	IF @piResult < 0
	BEGIN
		SET @piResult = @piResult + 12
	END

END




GO

