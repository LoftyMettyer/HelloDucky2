
CREATE PROCEDURE sp_ASRFn_IsEmpty_2
(
	@pfResult	bit OUTPUT,
	@pdblNumeric	float
)
AS
BEGIN
	SET @pfResult = 0

	IF @pdblNumeric = 0 
	BEGIN
		SET @pfResult = 1
	END

	IF @pdblNumeric IS null
	BEGIN
		SET @pfResult = 1
	END
END

GO

