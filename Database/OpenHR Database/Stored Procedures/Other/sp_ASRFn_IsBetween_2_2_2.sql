
CREATE PROCEDURE sp_ASRFn_IsBetween_2_2_2
(
	@pfResult		bit OUTPUT,
	@pdblNumericTest	float,
	@pdblNumericLower 	float,
	@pdblNumericUpper 	float
)
AS
BEGIN
	IF (@pdblNumericTest >= @pdblNumericLower) AND (@pdblNumericTest <= @pdblNumericUpper)
	BEGIN
		SET @pfResult = 1
	END
	ELSE
	BEGIN
		SET @pfResult = 0
	END
END




GO

