
CREATE PROCEDURE sp_ASRFn_IsEmpty_3
(
	@pfResult	bit OUTPUT,
	@pfLogic	bit
)
AS
BEGIN
	SET @pfResult = 0

	IF @pfLogic = 0 
	BEGIN
		SET @pfResult = 1
	END

	IF @pfLogic IS null
	BEGIN
		SET @pfResult = 1
	END
END



GO

