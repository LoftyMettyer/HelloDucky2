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
			ELSE (DATEDIFF(dd, @datefrom, @dateto) + 1)
				- (DATEDIFF(wk, @datefrom, @dateto) * 2)
				- (CASE WHEN DATEPART(dw, @datefrom) = 1 THEN 1 ELSE 0 END)
				- (CASE WHEN DATEPART(dw, @dateto) = 7 THEN 1 ELSE 0 END)
				END;
				
		RETURN ISNULL(@result,0);
		
	END