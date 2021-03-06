CREATE FUNCTION [dbo].[udfASRGetDate]()
	RETURNS datetime
	AS
BEGIN

		DECLARE @dtDate datetime

		SELECT TOP 1 @dtDate = convert(datetime, convert(varchar(20), last_batch, 101))
		FROM master..sysprocesses
		ORDER BY last_batch DESC

		RETURN @dtDate

	END
