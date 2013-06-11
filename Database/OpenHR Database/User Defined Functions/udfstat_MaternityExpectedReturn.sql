CREATE FUNCTION [dbo].[udfstat_MaternityExpectedReturn] (
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(MAX)
			)
	RETURNS datetime			
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @pdblResult datetime;

		IF LOWER(@Ordinary) = 'ordinary'
			IF DATEDIFF(d,'04/06/2003', @EWCDate) >= 0
				SET @pdblResult = DATEADD(ww,26,@LeaveStart);
			ELSE
				IF DATEDIFF(d,'04/30/2000', @EWCDate) >= 0
					SET @pdblResult = DATEADD(ww,18,@LeaveStart);
				ELSE
					SET @pdblResult = DATEADD(ww,14,@LeaveStart);
		ELSE
			IF DATEDIFF(d,'04/06/2003', @EWCDate) >= 0
				SET @pdblResult = DATEADD(ww,52,@LeaveStart);
			ELSE
				--29 weeks from baby birth date (but return on the monday before!)
				SET @pdblResult = DATEADD(d,203 - datepart(dw,DATEADD(d,-2,@BabyBirthDate)),@BabyBirthDate);

		RETURN @pdblResult;

	END