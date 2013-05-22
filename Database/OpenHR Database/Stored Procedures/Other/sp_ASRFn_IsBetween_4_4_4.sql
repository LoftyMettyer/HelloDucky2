
CREATE PROCEDURE sp_ASRFn_IsBetween_4_4_4
(
	@pfResult 		bit OUTPUT,
	@pdtDateTest 		datetime,
	@pdtDateLower		datetime,
	@pdtDateUpper 	datetime
)
AS
BEGIN
	/* Remove the time part of the dattime variables. */
	SET @pdtDateTest = convert(datetime, convert(varchar(20), @pdtDateTest, 101))
	SET @pdtDateLower = convert(datetime, convert(varchar(20), @pdtDateLower, 101))
	SET @pdtDateUpper = convert(datetime, convert(varchar(20), @pdtDateUpper, 101))

	IF (@pdtDateTest >= @pdtDateLower) AND (@pdtDateTest <= @pdtDateUpper)
	BEGIN
		SET @pfResult = 1
	END
	ELSE
	BEGIN
		SET @pfResult = 0
	END
END



GO

