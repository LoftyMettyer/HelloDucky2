CREATE PROCEDURE [dbo].[sp_ASRFn_RoundToNearestNumber]
(
	@pfReturn 			float OUTPUT,
	@pfNumberToRound 	float,
	@pfNearestNumber	float
)
AS
BEGIN

	DECLARE @pfRemainder float;

	/* Calculate the remainder. Cannot use the % because it only works on integers and not floats. */
	set @pfReturn = 0;
	if @pfNearestNumber <= 0 return
	
	set @pfRemainder = @pfNumberToRound - (floor(@pfNumberToRound / @pfNearestNumber) * @pfNearestNumber);

	/* Formula for rounding to the nearest specified number */
	if ((@pfNumberToRound < 0) AND (@pfRemainder <= (@pfNearestNumber / 2.0)))
		OR ((@pfNumberToRound >= 0) AND (@pfRemainder < (@pfNearestNumber / 2.0)))
			set @pfReturn = @pfNumberToRound - @pfRemainder;
		else
			set @pfReturn = @pfNumberToRound + @pfNearestNumber - @pfRemainder;

END