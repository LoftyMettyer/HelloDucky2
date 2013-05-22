CREATE FUNCTION [dbo].[udfsys_getfieldfromdatabaserecord](
	@searchcolumn AS nvarchar(255),
	@searchexpression AS nvarchar(MAX),
	@returnfield AS nvarchar(255))
RETURNS nvarchar(MAX)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result nvarchar(MAX);
	RETURN @result;

END