CREATE FUNCTION [dbo].[udfsys_statutoryredundancypay](
	@startdate AS datetime,
	@leavingdate AS datetime,
	@dateofbirth AS datetime,
	@weeklyrate AS numeric(10,2),
	@limit as numeric(10,2))
RETURNS numeric(10,2)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result numeric(10,2);
	DECLARE @service_years integer;
	
	--/* First three parameters are compulsory, so return 0 and exit if they are not set */
	IF (@startdate IS null) OR (@leavingdate IS null) OR (@weeklyrate IS null)
	BEGIN
		SET @result = 0;
		RETURN @result;
	END

	-- Calculate service years
	SET @service_years = [dbo].[udfsys_wholeyearsbetweentwodates](@startdate, @leavingdate);

	SET @result = @service_years * @weeklyrate;

	RETURN @result;

END