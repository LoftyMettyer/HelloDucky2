
CREATE PROCEDURE sp_ASRFn_IsEmpty_4
(
	@pfResult	bit OUTPUT,
	@pdtDate	datetime
)
AS
BEGIN
	SET @pfResult = 0

	IF @pdtDate IS null
	BEGIN
		SET @pfResult = 1
	END
END




GO

