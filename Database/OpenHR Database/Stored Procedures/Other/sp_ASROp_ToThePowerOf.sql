
CREATE PROCEDURE sp_ASROp_ToThePowerOf 
(
	@pdlbResult	float OUTPUT,
	@pdblFirst	float,
	@pdblSecond	float
)
AS
BEGIN
	IF (@pdblFirst = 0) AND (@pdblSecond < 0)
	BEGIN
		SET @pdlbResult = 0
	END
	ELSE
	BEGIN
		SET @pdlbResult = power(@pdblFirst, @pdblSecond)
	END
END


GO

