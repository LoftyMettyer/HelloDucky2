CREATE FUNCTION dbo.udfASRIsServer64Bit
()
RETURNS int
AS
BEGIN

	DECLARE @bIs64Bit bit
	SELECT @bIs64Bit = CASE PATINDEX ('%X64)%' , @@version)
			WHEN 0 THEN 0
			ELSE 1
		END
	RETURN @bIs64Bit

END
GO