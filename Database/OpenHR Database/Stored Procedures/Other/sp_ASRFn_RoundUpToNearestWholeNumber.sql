CREATE PROCEDURE [dbo].[sp_ASRFn_RoundUpToNearestWholeNumber]
(
	@piResult 		integer OUTPUT,	
	@pdblNumber 	float
)
AS
BEGIN
	SET @piResult = CASE WHEN @pdblNumber < 0 THEN floor(@pdblNumber)
		ELSE ceiling(@pdblNumber)
		END;
END