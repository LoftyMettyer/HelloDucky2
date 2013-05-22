
CREATE PROCEDURE sp_ASRFn_WholeMonthsBetweenTwoDates 
(
	@piResult	integer OUTPUT,
	@pdtDate1 	datetime,
	@pdtDate2 	datetime
)
AS
BEGIN
	SET @pdtDate1 = convert(datetime, convert(varchar(20), @pdtDate1, 101))
	SET @pdtDate2 = convert(datetime, convert(varchar(20), @pdtDate2, 101))

	IF @pdtDate1 < @pdtDate2
	BEGIN
		/* Get the total number of months*/
		SET @piResult = dateDiff(mm, @pdtDate1, @pdtDate2)
      
		/* See if the day field of pvParam2 < pvParam1 day field and if so - 1 */
		IF day(@pdtDate2) < day(@pdtDate1)
		BEGIN
			SET @piResult = @piResult -1
		END
	END
END



GO

