CREATE FUNCTION [dbo].[udfsys_uniquecode](
	@prefix AS nvarchar(max),
	@coderoot AS numeric(10,2))
RETURNS numeric(10,0)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result numeric(10,0);

	SET @result = 0;

	--SELECT @result = [maxcodesuffix] 
	--	FROM [dbo].[tb_uniquecodes]
	--	WHERE [codeprefix] = @prefix;

	-- Update existing value 
	/*
	You can't run an execute or an update in a UDF, so will have to create an extended
	stored procedure which should be able to do it. Otherwise tack something into the end
	of the update trigger on the base table that calls this function. (code stub already in
	the admin module.
	*/

	RETURN @result;

END
