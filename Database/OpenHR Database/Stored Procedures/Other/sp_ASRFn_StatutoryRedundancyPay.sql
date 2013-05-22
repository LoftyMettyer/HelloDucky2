CREATE PROCEDURE [dbo].[sp_ASRFn_StatutoryRedundancyPay]
(
	@pdblRedundancyPay	float OUTPUT,
	@pdtStartDate 		datetime,
	@pdtLeaveDate 		datetime,
	@pdtDOB				datetime,
	@pdblWeeklyRate 	float,
	@pdblStatLimit 		float
)
AS
BEGIN
	DECLARE @dtMinAgeBirthday	datetime,
		@dtServiceFrom			datetime,
		@iServiceYears 			integer,
		@iAgeY					integer,
		@iAgeM 					integer,
		@dblRate1 				float,
		@dblRate2 				float,
		@dblRate3 				float,
		@dtTempDate 			datetime,
		@iTempAgeY				integer,
		@iTemp					integer,
		@dblTemp2 				float,
		@iAfterOct2006			bit,
		@iMinAge				integer;

	SET @pdblRedundancyPay = 0
	SET @iAfterOct2006 = case when datediff(dd,@pdtLeaveDate,'10/01/2006') <= 0 then 1 else 0 end

	if @iAfterOct2006 = 1
		SET @iMinAge = 16
	else
		SET @iMinAge = 18

	/* First three parameters are compulsory, so return 0 and exit if they are not set */
	IF (@pdtStartDate IS null) OR (@pdtLeaveDate IS null) OR (@pdtDOB IS null)
	BEGIN
		RETURN
	END

	SET @pdtStartDate = convert(datetime, convert(varchar(20), @pdtStartDate, 101))
	SET @pdtLeaveDate = convert(datetime, convert(varchar(20), @pdtLeaveDate, 101))
	SET @pdtDOB = convert(datetime, convert(varchar(20), @pdtDOB, 101))


	/* Calc start date */
   	SET @dtServiceFrom = @pdtStartDate
	if @iAfterOct2006 = 0
	BEGIN
		SET @dtMinAgeBirthday = dateadd(yy, @iMinAge, @pdtDOB)
		IF @dtMinAgeBirthday >= @pdtStartDate
			SET @dtServiceFrom = @dtMinAgeBirthday
	END


	/* Calc number of applicable complete yrs the employee has been employed */
	exec sp_ASRFn_WholeYearsBetweenTwoDates @iServiceYears OUTPUT, @dtServiceFrom, @pdtLeaveDate

	/* exit if its less than 2 years */
	IF @iServiceYears < 2 
	BEGIN
		RETURN
	END

	/* calculate the employees years and months to the leave date */
	exec sp_ASRFn_WholeYearsBetweenTwoDates @iAgeY OUTPUT, @pdtDOB, @pdtLeaveDate

	SET @dtTempDate = dateadd(yy, @iAgeY, @pdtDOB)
	exec sp_ASRFn_WholeMonthsBetweenTwoDates @iAgeM OUTPUT, @dtTempDate, @pdtLeaveDate

	/* only count up to 20 years for redundancy */
	exec sp_ASRFn_Minimum @iServiceYears OUTPUT, 20, @iServiceYears

	/* fill in the rates depending on service and age */
	SET @iTempAgeY = @iAgeY
	SET @dblRate1 = 0
	SET @dblRate2 = 0
	SET @dblRate3 = 0

	IF @iTempAgeY >= 41
	BEGIN
		SET @iTemp = @iTempAgeY - 41
		exec sp_ASRFn_Minimum @dblRate1 OUTPUT, @iTemp, @iServiceYears
		SET @iTempAgeY = 41
		SET @iServiceYears = @iServiceYears - @dblRate1
	END

	IF @iTempAgeY >= 22
	BEGIN
		SET @iTemp = @iTempAgeY - 22
		exec sp_ASRFn_Minimum @dblRate2 OUTPUT, @iTemp, @iServiceYears
		SET @iTempAgeY = 22
		SET @iServiceYears = @iServiceYears - @dblRate2
	END

	IF @iTempAgeY >= @iMinAge
	BEGIN
		SET @iTemp = @iTempAgeY - @iMinAge
		exec sp_ASRFn_Minimum @dblRate3 OUTPUT, @iTemp, @iServiceYears
	END

	/* calc the redundancy pay */
	exec sp_ASRFn_Minimum @dblTemp2 OUTPUT, @pdblWeeklyRate, @pdblStatLimit

	SET @pdblRedundancyPay = ((@dblRate1 * 1.5) + (@dblRate2) + (@dblRate3 * 0.5)) * @dblTemp2

	if @iAfterOct2006 = 0
	begin
		IF @iAgeY = 64 
		BEGIN
			SET @pdblRedundancyPay = @pdblRedundancyPay * (12 - @iAgeM) / 12
		END
	end
END