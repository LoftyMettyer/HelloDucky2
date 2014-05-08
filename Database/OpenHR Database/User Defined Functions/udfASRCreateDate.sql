CREATE FUNCTION [dbo].[udfASRCreateDate](@day float, @month float, @year float)
RETURNS datetime
AS
BEGIN

	DECLARE @date varchar(20);

	IF @day < 1 OR @month < 1 OR @year < 1 OR @month > 12 OR @day > 31 OR @year > 9999  RETURN NULL;

	SET @date = CONVERT(varchar(2), @month) + '/' + CONVERT(varchar(2), @day) + '/' + CONVERT(varchar(4), @year);

	IF ISDATE(@date) = 0
		RETURN NULL;

	RETURN @date;

END