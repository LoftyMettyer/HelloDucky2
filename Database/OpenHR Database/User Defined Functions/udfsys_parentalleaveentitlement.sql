CREATE FUNCTION [dbo].[udfsys_parentalleaveentitlement] (
     @id		integer)
RETURNS integer
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result integer;
	
	SET @result = 10;
        
    RETURN @result;

END