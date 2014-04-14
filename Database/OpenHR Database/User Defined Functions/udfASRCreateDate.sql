CREATE FUNCTION [dbo].[udfASRCreateDate](@day float, @month float, @year float)
RETURNS datetime
AS
BEGIN

	DECLARE @date varchar(20);

	IF @day = 0 OR @month = 0 OR @year = 0 RETURN NULL;
	SET @date = CONVERT(varchar(2), @month) + '/' + CONVERT(varchar(2), @day) + '/' + CONVERT(varchar(4), @year);

	IF ISDATE(@date) = 0
		RETURN NULL;

	RETURN @date;

END