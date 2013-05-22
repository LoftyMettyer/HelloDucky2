CREATE PROCEDURE [dbo].[spASRMaternityExpectedReturn] (
	@pdblResult datetime OUTPUT,
	@EWCDate datetime,
	@LeaveStart datetime,
	@BabyBirthDate datetime,
	@Ordinary varchar(MAX)
	)
AS
BEGIN

	IF LOWER(@Ordinary) = 'ordinary'
		IF DateDiff(d,'04/06/2003', @EWCDate) >= 0
			SET @pdblResult = Dateadd(ww,26,@LeaveStart);
		ELSE
			IF DateDiff(d,'04/30/2000', @EWCDate) >= 0
				SET @pdblResult = Dateadd(ww,18,@LeaveStart);
			ELSE
				SET @pdblResult = Dateadd(ww,14,@LeaveStart);
	ELSE
		IF DateDiff(d,'04/06/2003', @EWCDate) >= 0
			SET @pdblResult = Dateadd(ww,52,@LeaveStart);
		ELSE
			--29 weeks from baby birth date (but return on the monday before!)
			SET @pdblResult = DateAdd(d,203 - datepart(dw,DateAdd(d,-2,@BabyBirthDate)),@BabyBirthDate);

END
GO

