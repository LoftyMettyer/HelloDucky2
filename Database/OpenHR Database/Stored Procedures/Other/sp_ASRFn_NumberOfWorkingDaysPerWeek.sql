CREATE PROCEDURE [dbo].[sp_ASRFn_NumberOfWorkingDaysPerWeek]
(
	@pdblResult 	float OUTPUT,
	@psPattern 		varchar(MAX)		
	/* Working pattern. 14 characters long in the format 'SsMmTtWwTtFfSs'
	where a uppercase letter relates to the morning, and the lowercase letter relates to the afternnon of the appropriate day. 
	A space means that the morning/afternoon is not worked, anything else means that the session is worked. */
)
AS
BEGIN
	DECLARE @iCounter	integer;

	SET @pdblResult = 0;
	SET @iCounter = 0;

	WHILE @iCounter <= LEN(@psPattern)
	BEGIN
		IF SUBSTRING(@psPattern, @iCounter, 1) <> ' '
		BEGIN
			SET @pdblResult = @pdblResult + 0.5;
		END

		SET @iCounter = @iCounter + 1;
	END
END