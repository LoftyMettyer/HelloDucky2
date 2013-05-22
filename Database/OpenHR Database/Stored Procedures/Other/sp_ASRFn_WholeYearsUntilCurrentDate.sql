
CREATE PROCEDURE sp_ASRFn_WholeYearsUntilCurrentDate
(
	@piResult	integer OUTPUT,
	@pdtDate	datetime
)
AS
BEGIN
	DECLARE @dtToday	datetime

	SET @dtToday = getdate()
	SET @pdtDate = convert(datetime, convert(varchar(20), @pdtDate, 101))

	/* Get the number of whole years */
	SET @piResult = year(@dtToday) - year(@pdtDate)
  
	/* See if the date passed in months are greater than todays month */
	IF month(@pdtDate) > month(@dtToday)
	BEGIN
		SET @piResult = @piResult - 1
	END
  
	/* See if the months are equal and if they are test the day value */
	IF month(@pdtDate) = month(@dtToday)
	BEGIN
		IF day(@pdtDate) > day(@dtToday)
		BEGIN
			SET @piResult = @piResult - 1
		END
	END
END





GO

