CREATE FUNCTION [dbo].[udfsys_parentalleavetaken] (
     @id		integer)
RETURNS integer
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result integer;
	
	SET @result = 0;
        
    RETURN @result;

END