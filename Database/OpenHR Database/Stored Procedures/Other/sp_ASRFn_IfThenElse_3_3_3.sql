
CREATE PROCEDURE sp_ASRFn_IfThenElse_3_3_3
(
	@pfResult		bit OUTPUT,
	@pfTestValue		bit,
	@pfLogic1		bit,
	@pfLogic2		bit
)
AS
BEGIN
	IF @pfTestValue = 1
	BEGIN
		SET @pfResult = @pfLogic1
	END
	ELSE
	BEGIN
		SET @pfResult = @pfLogic2
	END	
END




GO

