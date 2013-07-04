CREATE PROCEDURE [dbo].[sp_ASRIntGetTables] AS
BEGIN

	SET NOCOUNT ON;

	SELECT tableID, tableName
		FROM [dbo].[ASRSysTables]
		ORDER BY tableName;
END