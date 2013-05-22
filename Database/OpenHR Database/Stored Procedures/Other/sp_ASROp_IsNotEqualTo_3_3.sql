
CREATE PROCEDURE sp_ASROp_IsNotEqualTo_3_3
(
	@pfResult	bit OUTPUT,
	@pfLogic1	bit,
	@pfLogic2	bit
)
AS
BEGIN
	if @pfLogic1 = @pfLogic2
	BEGIN
		SET @pfResult = 0
	END
	ELSE
	BEGIN
		SET @pfResult = 1
	END	
END



GO

