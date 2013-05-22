CREATE FUNCTION [dbo].[udfsys_getfunctionparametertype]
	(@functionid integer, @parameterindex integer)
RETURNS integer
AS
BEGIN

	DECLARE @result integer;

	SELECT @result = [parametertype] FROM ASRSysFunctionParameters
		WHERE @functionid = [functionID] AND @parameterindex = [parameterIndex];

	RETURN @result;

END