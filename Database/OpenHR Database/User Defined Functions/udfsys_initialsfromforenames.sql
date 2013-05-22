CREATE FUNCTION [dbo].[udfsys_initialsfromforenames] 
(
	@forenames	varchar(8000),
	@padwithspace bit
)
RETURNS nvarchar(10)
WITH SCHEMABINDING
AS
BEGIN

	DECLARE @result nvarchar(10) = '';
	DECLARE @icounter integer = 1;

	IF LEN(@forenames) > 0 
	BEGIN
		SET @result = UPPER(left(@forenames,1));

		WHILE @icounter < LEN(@forenames)
		BEGIN
			IF SUBSTRING(@forenames, @icounter, 1) = ' '
			BEGIN
				IF @padwithspace = 1
					SET @result = @result + ' ' + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
				ELSE
					SET @result = @result + UPPER(SUBSTRING(@forenames, @icounter+1, 1));
			END
	
			SET @icounter = @icounter +1;
		END

		SET @result = @result + ' '
	
	END

	RETURN @result

END