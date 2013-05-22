
CREATE PROCEDURE sp_ASRFn_WeekdaysFromStartAndEndDates 
(
	@piResult	integer OUTPUT,
	@pdtDate1 	datetime,
	@pdtDate2 	datetime
)
AS
BEGIN
	DECLARE @iCounter	integer

	SET @piResult = 0
	SET @iCounter = 0
	SET @pdtDate1 = convert(datetime, convert(varchar(20), @pdtDate1, 101))
	SET @pdtDate2 = convert(datetime, convert(varchar(20), @pdtDate2, 101))

	WHILE @iCounter <= datediff(day, @pdtDate1, @pdtDate2)
	BEGIN
		IF datepart(dw, dateadd(day, @iCounter, @pdtDate1)) <> 1
		BEGIN
			IF datepart(dw, dateadd(day, @iCounter, @pdtDate1)) <> 7
			BEGIN
				SET @piResult = @piResult + 1
			END
		END

		SET @iCounter = @iCounter + 1
	END
END



GO

