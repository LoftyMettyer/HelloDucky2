CREATE PROCEDURE [dbo].[sp_ASRIntGetExprOperators]
AS
BEGIN
	/* Return a recordset of tab-delimted operator definitions ;
	<operator id><tab><operator name><tab><operator category> */
	SELECT 
		convert(varchar(100), operatorID) + char(9) +
		name + 
		CASE 
			WHEN len(shortcutKeys) > 0 THEN ' (' + shortcutKeys + ')'
			ELSE ''
		END + char(9) +
		category AS [definitionString]
	FROM [dbo].[ASRSysOperators];
END