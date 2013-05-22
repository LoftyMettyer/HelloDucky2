
CREATE PROCEDURE sp_ASRFn_IsPopulated_4
(
	@pfResult	bit OUTPUT,
	@pdtDate	datetime
)
AS
BEGIN
	SET @pfResult = 1

	IF @pdtDate IS null
	BEGIN
		SET @pfResult = 0
	END
END



GO

