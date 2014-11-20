CREATE PROCEDURE [dbo].[spASRIntGetColumnControlValues]
	@ColumnIDs nvarchar(100)
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @sql nvarchar(MAX);

	SET @sql = 'SELECT columnID, Value, sequence FROM ASRSysColumnControlValues WHERE columnID IN (' + @ColumnIDs + ')'
	EXECUTE sp_executeSQL @sql;
END