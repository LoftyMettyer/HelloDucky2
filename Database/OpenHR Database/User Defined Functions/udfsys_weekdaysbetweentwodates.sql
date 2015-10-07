CREATE FUNCTION [dbo].[udfsys_weekdaysbetweentwodates](
		@datefrom AS datetime,
		@dateto AS datetime)
	RETURNS integer
	WITH SCHEMABINDING
	AS
	BEGIN
	
		DECLARE @result integer;

		SELECT @result = CASE 
			WHEN DATEDIFF (day, @datefrom, @dateto) <= 0 THEN 0
			ELSE DATEDIFF(day, @datefrom, @dateto + 1) 
				- (2 * (DATEDIFF(day, @datefrom - (DATEPART(dw, @datefrom) -1),
					@dateto	- (DATEPART(dw, @dateto) - 1)) / 7))
				- CASE WHEN DATEPART(dw, @datefrom) = 1 THEN 1 ELSE 0 END
				- CASE WHEN DATEPART(dw, @dateto) = 7 THEN 1 ELSE 0	END
				END;
				
		RETURN ISNULL(@result,0);
		
	END