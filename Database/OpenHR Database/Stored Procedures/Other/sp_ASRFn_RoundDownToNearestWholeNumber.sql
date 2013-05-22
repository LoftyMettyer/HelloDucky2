CREATE PROCEDURE [dbo].[sp_ASRFn_RoundDownToNearestWholeNumber]
(
	@piResult 	integer OUTPUT,	
	@pdblNumber float
)
AS
BEGIN
	SET @piResult = ROUND(@pdblNumber, 0, 1);
END