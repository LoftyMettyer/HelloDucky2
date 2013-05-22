CREATE FUNCTION [dbo].[udfsys_nicedate](
	@inputdate as datetime)
RETURNS nvarchar(max)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result varchar(MAX) = '';
	SELECT @result = CONVERT(nvarchar(2),DATEPART(day, @inputdate))
		+ ' ' + DATENAME(month, @inputdate) 
		+ ' ' + CONVERT(nvarchar(4),DATEPART(YYYY, @inputdate));

	RETURN @result;
	
END