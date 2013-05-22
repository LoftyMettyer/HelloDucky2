CREATE FUNCTION [dbo].[udfsys_propercase](
	@text as nvarchar(max))
RETURNS nvarchar(max)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @reset	bit = 1;
	DECLARE @result varchar(8000) = '';
	DECLARE @i		int = 1;
	DECLARE @c		char(1);
      
	WHILE (@i <= len(@text))
		SELECT @c= substring(@text,@i,1)
			, @result = @result + CASE WHEN @reset=1 THEN UPPER(@c) 
									   ELSE LOWER(@c) END
			, @reset = CASE WHEN @c LIKE '[a-zA-Z]' THEN 0
							ELSE 1
							END
			, @i = @i + 1;

	RETURN @result;
	
END