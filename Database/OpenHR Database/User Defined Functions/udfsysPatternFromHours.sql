CREATE FUNCTION [dbo].[udfsysPatternFromHours] (
	@PatternType	varchar(5),
	@Sunday_Hours numeric(4,2),
	@Monday_Hours numeric(4,2),
	@Tuesday_Hours numeric(4,2),
	@Wednesday_Hours numeric(4,2),
	@Thursday_Hours numeric(4,2),
	@Friday_Hours numeric(4,2),
	@Saturday_Hours numeric(4,2))
RETURNS varchar(28)
AS 
BEGIN

	DECLARE @value varchar(28);

	IF @PatternType = 'Days'
		SET @value = CASE WHEN @Sunday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Monday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Tuesday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Wednesday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Thursday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Friday_Hours > 0 THEN '1' ELSE '0' END +
						CASE WHEN @Saturday_Hours > 0 THEN '1' ELSE '0' END;
	ELSE
		SET @value = CASE WHEN ISNULL(@Sunday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Sunday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Monday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Monday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Tuesday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Tuesday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Wednesday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Wednesday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Thursday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Thursday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Friday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Friday_Hours), 5), '.','') ELSE '0000' END +
						CASE WHEN ISNULL(@Saturday_Hours,0) > 0 THEN REPLACE(RIGHT('00000' + CONVERT(varchar(5), @Saturday_Hours), 5), '.','') ELSE '0000' END;

	RETURN @value;

END






