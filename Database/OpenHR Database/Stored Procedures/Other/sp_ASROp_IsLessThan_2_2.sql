
CREATE PROCEDURE sp_ASROp_IsLessThan_2_2
(
	@pfResult   	bit OUTPUT,
	@pdblNumeric1	float,
	@pdblNumeric2	float
)
AS
BEGIN
	IF @pdblNumeric1 < @pdblNumeric2
	BEGIN
		SET @pfResult = 1
	END
	ELSE
	BEGIN
		SET @pfResult = 0
	END	
END


GO

