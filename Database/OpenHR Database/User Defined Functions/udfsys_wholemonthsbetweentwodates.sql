CREATE FUNCTION [dbo].[udfsys_wholemonthsbetweentwodates] 
(
	@date1 	datetime,
	@date2 	datetime
)
RETURNS integer
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result integer;

	-- Clean dates (trim time part)
	SET @date1 = DATEADD(D, 0, DATEDIFF(D, 0, @date1));
	SET @date2 = DATEADD(D, 0, DATEDIFF(D, 0, @date2));

	IF @date1 < @date2
	BEGIN

		-- Get the total number of months
		SET @result = DATEDIFF(mm, @date1, @date2);
      
		-- See if the day field of pvParam2 < pvParam1 day field and if so - 1
		IF DAY(@date2) < DAY(@date1)
		BEGIN
			SET @result = @result -1;
		END
	END
	
	RETURN @result
	
END