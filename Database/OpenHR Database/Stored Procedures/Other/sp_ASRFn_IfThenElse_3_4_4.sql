
CREATE PROCEDURE sp_ASRFn_IfThenElse_3_4_4
(
	@pdtResult   		datetime OUTPUT,
	@pfTestValue		bit,
	@pdtDate1		datetime,
	@pdtDate2		datetime
)
AS
BEGIN
	IF @pfTestValue = 1
	BEGIN
		SET @pdtResult = @pdtDate1
	END
	ELSE
	BEGIN
		SET @pdtResult = @pdtDate2
	END	
END


GO

