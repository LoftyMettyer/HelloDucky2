CREATE FUNCTION [dbo].[udfsys_workingdaysbetweentwodates]
(
	@date1 	datetime,
	@date2 	datetime
)
RETURNS integer
WITH SCHEMABINDING
AS
BEGIN
	
	RETURN 0
	
END