CREATE PROCEDURE [dbo].[spASRParentalLeaveEntitlement] (
	@pdblResult		float OUTPUT,
	@DateOfBirth	datetime,
	@AdoptedDate	datetime,
	@Disabled		bit,
	@Region			varchar(MAX)
)
AS
BEGIN

	DECLARE @Today datetime,
		@ChildAge int,
		@Adopted bit,
		@YearsOfResponsibility int,
		@StartDate datetime,
		@Standard int,
		@Extended int;

	SET @Standard = 65;
	SET @Extended = 90;
	IF @Region = 'Rep of Ireland'
	BEGIN
		SET @Standard = 70;
		SET @Extended = 70;
	END


	--Check if we should used the Date of Birth or the Date of Adoption column...
	SET @Adopted = 0;
	SET @StartDate = @DateOfBirth;
	IF NOT @AdoptedDate IS NULL
	BEGIN
		SET @Adopted = 1;
		SET @StartDate = @AdoptedDate;
	END

	--Set variables based on this date...
	--(years of responsibility = years since born or adopted)
	SET @Today = getdate();
	EXEC [dbo].[sp_ASRFn_WholeYearsBetweenTwoDates] @ChildAge OUTPUT, @DateOfBirth, @Today;
	EXEC [dbo].[sp_ASRFn_WholeYearsBetweenTwoDates] @YearsOfResponsibility OUTPUT, @StartDate, @Today;

	SELECT @pdblResult = CASE
		WHEN @Disabled = 0 And @Adopted = 0 And @ChildAge < 5
			THEN @Standard
		WHEN @Disabled = 0 And @Adopted = 1 And @ChildAge < 18
			And @YearsOfResponsibility < 5 THEN	@Standard
		WHEN @Disabled = 1 And @Adopted = 0 And @ChildAge < 18 
			And DateDiff(d,'12/15/1994',@DateOfBirth) >= 0 THEN	@Extended
		WHEN @Disabled = 1 And @Adopted = 1 And @ChildAge < 18 
		And DateDiff(d,'12/15/1994',@AdoptedDate) >= 0 THEN	@Extended
		ELSE
			0
		END;

END