CREATE PROCEDURE [dbo].[sp_ASRIntGetTables] AS
BEGIN
	SELECT tableID, tableName
	FROM [dbo].[ASRSysTables]
	ORDER BY tableName;
END