
CREATE PROCEDURE sp_ASRFn_Minimum 
(
	@pdblResult	float OUTPUT,
	@pdblNumeric1 float,
	@pdblNumeric2	float
)
AS
BEGIN
	IF @pdblNumeric1 <= @pdblNumeric2
	BEGIN
		SET @pdblResult = @pdblNumeric1
	END
	ELSE
	BEGIN
		SET @pdblResult = @pdblNumeric2
	END
END


GO

