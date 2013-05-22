CREATE FUNCTION [dbo].[udfsys_firstnamefromforenames] 
(
	@forenames nvarchar(max)
)
RETURNS nvarchar(max)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result nvarchar(max);

	IF (LEN(@forenames) = 0 ) OR (@forenames IS null)
	BEGIN
		SET @result = '';
	END
	ELSE
	BEGIN
		IF CHARINDEX(' ', @forenames) > 0
			SET @result = LEFT(@forenames, CHARINDEX(' ', @forenames));
		ELSE
			SET @result = @forenames;
	END
	
	RETURN @result;
	
END