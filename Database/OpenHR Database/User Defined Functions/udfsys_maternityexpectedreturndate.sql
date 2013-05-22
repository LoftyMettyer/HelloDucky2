CREATE FUNCTION [dbo].[udfsys_maternityexpectedreturndate] (
     @id		integer)
RETURNS datetime
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result datetime;
	
	SET @result = GETDATE()
        
    RETURN @result;

END