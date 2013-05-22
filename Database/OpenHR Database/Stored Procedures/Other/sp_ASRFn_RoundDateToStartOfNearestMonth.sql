
CREATE PROCEDURE sp_ASRFn_RoundDateToStartOfNearestMonth 
(
	@pdtResult 	datetime OUTPUT,
	@pdtDate 	datetime
)
AS
BEGIN
	DECLARE @dtDateNextMonth	datetime, /* start of next month */
		@dtDateThisMonth 	datetime /* start of this month */

	SET @pdtDate = convert(datetime, convert(varchar(20), @pdtDate, 101))

	/* Create a date with one month added to the date and move it to the first day of that month */
	SET @dtDateNextMonth = dateAdd(mm, 1, @pdtDate)
	SET @dtDateNextMonth = dateAdd(dd, -1 * (day(@dtDateNextMonth) - 1), @dtDateNextMonth)

	/* Create a date which is the first of the month passed in */
	SET @dtDateThisMonth = dateAdd(dd, -1 * (day(@pdtDate) - 1), @pdtDate)
    
	/* See which is the greatest gap between the two start month dates and the passed in date */
	IF (@pdtDate - (@dtDateThisMonth) + 1) < ((@dtDateNextMonth) - (@pdtDate))
	BEGIN
		SET @pdtResult = @dtDateThisMonth
	END
	ELSE
	BEGIN
		SET @pdtResult = @dtDateNextMonth
	END
END

GO

