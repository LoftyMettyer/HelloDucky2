CREATE FUNCTION [dbo].[udfASRSQLVersion]
	(
	)
	RETURNS integer
	AS
	BEGIN
		RETURN convert(int,convert(float,substring(@@version,charindex('-',@@version)+2,2)))
	END