CREATE FUNCTION [dbo].[udfsys_wholeyearsbetweentwodates] (
     @date1  datetime,
     @date2  datetime )
RETURNS integer 
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result integer = 0;
	
    -- Get the number of whole years
    SET @result = YEAR(@date2) - YEAR(@date1);

    -- See if the date passed in months are greater than todays month
    IF MONTH(@date1) > MONTH(@date2)
    BEGIN
		SET @result = @result - 1;
    END
    
    -- See if the months are equal and if they are test the day value
    IF MONTH(@date1) = MONTH(@date2)
    BEGIN
        IF DAY(@date1) > DAY(@date2)
            BEGIN
				SET @result = @result - 1;
            END
        END
        
    RETURN @result;

END