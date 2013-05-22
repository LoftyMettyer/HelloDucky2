
CREATE PROCEDURE sp_ASROp_Not 
(
	@pfResult	bit OUTPUT,
	@pfFirst		bit
)
AS
BEGIN
	SET @pfResult = ~ @pfFirst
END




GO

