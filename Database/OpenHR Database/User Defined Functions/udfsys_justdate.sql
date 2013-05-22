CREATE FUNCTION udfsys_justdate
	(@date datetime)
RETURNS datetime
WITH SCHEMABINDING
AS
BEGIN
	RETURN DATEADD(D, 0, DATEDIFF(D, 0, @date));
END