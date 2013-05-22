CREATE PROCEDURE sp_ASRFn_ServiceYears 
(
	@piResult		integer OUTPUT,
	@pdtFirstDate 		datetime,
	@pdtSecondDate	datetime
)
AS
BEGIN

	DECLARE @pdtTempDate	datetime

	/* If start date is in the future then return zero */
	IF datediff(d,@pdtFirstDate,getdate()) < 1 or @pdtFirstDate IS null
		SET @piResult = 0
	ELSE
	BEGIN
		IF datediff(d,@pdtSecondDate,getdate()) < 1 or @pdtSecondDate IS null
			/* If leaving date is in the future or blank then calculate from todays date minus start date */
			SET @pdtTempDate = getdate()
		ELSE
			/* If leaving date is in past then calculate from leaving date minus start date */
			SET @pdtTempDate = @pdtSecondDate

		EXEC sp_ASRFn_WholeYearsBetweenTwoDates @piResult OUTPUT, @pdtFirstDate, @pdtTempDate
		IF @piResult < 0 SET @piResult = 0

	END

END
GO

