CREATE FUNCTION [dbo].[udfsys_servicelength] (
     @startdate  datetime,
     @leavingdate  datetime,
     @period nvarchar(2))
RETURNS integer 
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result integer;
	DECLARE @amount integer;

	-- If start date is in the future ignore
	IF @startdate > GETDATE()
		RETURN 0;
	
	-- Trim the leaving date
	IF @leavingdate IS NULL OR @leavingdate > GETDATE()
		SET @leavingdate = GETDATE();


	SET @amount = [dbo].[udfsys_wholeyearsbetweentwodates]
		(@startdate, @leavingdate);

	-- Years	
	IF @period = 'Y' SET @result = @amount
	
	--Months
	ELSE IF @period = 'M'
		SET @result = [dbo].[udfsys_wholemonthsbetweentwodates]
			(@startdate, @leavingdate) - (@amount * 12);
	
    RETURN @result;

END