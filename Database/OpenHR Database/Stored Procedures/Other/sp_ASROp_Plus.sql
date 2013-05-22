
CREATE PROCEDURE sp_ASROp_Plus 
(
	@pdblResult	float OUTPUT,
	@pdblFirst	float,
	@pdblSecond	float
)
AS
BEGIN
	SET @pdblResult = @pdblFirst + @pdblSecond
END




GO

