CREATE FUNCTION [dbo].[udfsys_isfieldpopulated](
	@inputcolumn as nvarchar(MAX))
RETURNS bit
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result bit = 0;
	SELECT @result = (
		CASE 
			WHEN @inputcolumn IS NULL THEN 0 
			ELSE
				CASE
--					WHEN LEN(convert(nvarchar(1),@inputcolumn)) = 0 THEN 0
					WHEN DATALENGTH(@inputcolumn) = 0 THEN 0
					ELSE 1
				END
			END);

	RETURN @result;
	
END