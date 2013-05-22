
CREATE PROCEDURE sp_ASROp_IsEqualTo_4_4
(
	@pfResult   	bit OUTPUT,
	@pdtDate1	datetime,
	@pdtDate2	datetime
)
AS
BEGIN
	IF convert(datetime, convert(varchar(20), @pdtDate1, 101)) = convert(datetime, convert(varchar(20), @pdtDate2, 101))
	BEGIN
		SET @pfResult = 1
	END
	ELSE
	BEGIN
		SET @pfResult = 0
	END
END



GO

