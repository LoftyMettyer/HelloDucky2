CREATE PROCEDURE sp_ASROp_Modulus 
(
	@pdblResult	float OUTPUT,
	@pdblFirst	float,
	@pdblSecond 	float
)
AS
BEGIN
	IF @pdblSecond = 0 
	BEGIN
		SET @pdblResult = 0
	END
	ELSE
	BEGIN
		SET @pdblResult = @pdblFirst - (CAST((@pdblFirst / @pdblSecond) AS INT) * @pdblSecond)
	END
END
GO

