CREATE FUNCTION [dbo].[udfstat_ParentalLeaveEntitlement] (
		@DateOfBirth	datetime,
		@AdoptedDate	datetime,
		@Disabled		bit,
		@Region			varchar(MAX))
	RETURNS float
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @pdblResult			float,
			@Today					datetime,
			@ChildAge				integer,
			@Adopted				bit,
			@YearsOfResponsibility	integer,
			@StartDate				datetime,
			@Standard				integer,
			@Extended				integer;

		SET @Standard = 65;
		SET @Extended = 90;
		SET @Today = GETDATE();
				
		IF @Region = 'Rep of Ireland'
		BEGIN
			SET @Standard = 70;
			SET @Extended = 70;
		END

		IF DATEDIFF(d,'03-08-2013', @Today) >= 0
		BEGIN
			SET @Standard = 90;
			SET @Extended = 90;
		END

		-- Check if we should used the Date of Birth or the Date of Adoption column...
		SET @Adopted = 0;
		SET @StartDate = @DateOfBirth;
		IF NOT @AdoptedDate IS NULL
		BEGIN
			SET @Adopted = 1;
			SET @StartDate = @AdoptedDate;
		END

		-- Set variables based on this date...
		--( years of responsibility = years since born or adopted)
		SELECT @ChildAge = [dbo].[udfsys_wholeyearsbetweentwodates](@DateOfBirth, @Today);
		SELECT @YearsOfResponsibility = [dbo].[udfsys_wholeyearsbetweentwodates](@StartDate, @Today);

		SELECT @pdblResult = CASE
			WHEN @Disabled = 0 AND @Adopted = 0 AND @ChildAge < 5
				THEN @Standard
			WHEN @Disabled = 0 AND @Adopted = 1 AND @ChildAge < 18
				AND @YearsOfResponsibility < 5 THEN	@Standard
			WHEN @Disabled = 1 AND @Adopted = 0 AND @ChildAge < 18 
				AND DATEDIFF(d,'12/15/1994',@DateOfBirth) >= 0 THEN @Extended
			WHEN @Disabled = 1 AND @Adopted = 1 AND @ChildAge < 18 
				AND DATEDIFF(d,'12/15/1994',@AdoptedDate) >= 0 THEN @Extended
			ELSE 0
			END;

		RETURN ISNULL(@pdblResult,0);

	END