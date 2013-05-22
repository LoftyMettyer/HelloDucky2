CREATE PROCEDURE [dbo].[sp_ASRFn_IsEmpty]
(
    @result integer OUTPUT,
    @vartotest varchar(MAX)
)
AS
BEGIN
	
	IF LEN(@vartotest) = 0 
		SET @result = 1;

	IF @vartotest IS NULL
		SET @result = 1;

	IF LEN(@vartotest) > 0
		SET @result = 0;

	SELECT @result AS result;

END
