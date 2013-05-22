CREATE FUNCTION [dbo].[udfsys_nicetime](
	@inputdate as datetime)
RETURNS nvarchar(255)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result varchar(255) = '';

	SELECT @result =convert(char(8), @inputdate, 108)

	RETURN @result;
	
END