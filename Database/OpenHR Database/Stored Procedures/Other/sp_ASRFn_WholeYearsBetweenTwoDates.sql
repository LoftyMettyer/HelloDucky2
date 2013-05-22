
CREATE PROCEDURE sp_ASRFn_WholeYearsBetweenTwoDates 
(
	@piResult	integer OUTPUT,
	@pdtDate1 	datetime,
    	@pdtDate2 	datetime
)
AS
BEGIN
	/* Get the number of whole years */
	SET @piResult = year(@pdtDate2) - year(@pdtDate1)

	/* See if the date passed in months are greater than todays month */
	IF month(@pdtDate1) > month(@pdtDate2)
	BEGIN
		SET @piResult = @piResult - 1
	END
  
	/* See if the months are equal and if they are test the day value */
	IF month(@pdtDate1) = month(@pdtDate2)
	BEGIN
		IF day(@pdtDate1) > day(@pdtDate2)
		BEGIN
			SET @piResult = @piResult - 1	
		END
	END
END




GO

