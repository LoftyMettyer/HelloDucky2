
CREATE PROCEDURE sp_ASRFn_IsPopulated_2
(
	@pfResult	bit OUTPUT,
	@pdblNumeric	float
)
AS
BEGIN
	SET @pfResult = 1

	IF @pdblNumeric = 0 
	BEGIN
		SET @pfResult = 0
	END

	IF @pdblNumeric IS null
	BEGIN
		SET @pfResult = 0
	END
END

GO

