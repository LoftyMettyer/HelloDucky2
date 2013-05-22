
CREATE PROCEDURE sp_ASROp_Minus 
(
	@pdblResult	float OUTPUT,
	@pdblFirst	float,
	@pdblSecond	float
)
AS
BEGIN
	SET @pdblResult = @pdblFirst - @pdblSecond
END




GO

