
CREATE PROCEDURE sp_ASRFn_IsPopulated_3
(
	@pfResult	bit OUTPUT,
	@pfLogic	bit
)
AS
BEGIN
	SET @pfResult = 1

	IF @pfLogic = 0 
	BEGIN
		SET @pfResult = 0
	END

	IF @pfLogic IS null
	BEGIN
		SET @pfResult = 0
	END
END


GO

