
CREATE PROCEDURE sp_ASRFn_IfThenElse_3_2_2
(
	@pdblResult   		float OUTPUT,
	@pfTestValue		bit,
	@pdblNumeric1		float,
	@pdblNumeric2		float
)
AS
BEGIN
	IF @pfTestValue = 1
	BEGIN
		SET @pdblResult = @pdblNumeric1
	END
	ELSE
	BEGIN
		SET @pdblResult = @pdblNumeric2
	END	
END





GO

