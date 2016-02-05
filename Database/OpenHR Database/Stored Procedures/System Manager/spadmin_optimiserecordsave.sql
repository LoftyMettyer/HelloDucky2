CREATE PROCEDURE [dbo].[spadmin_optimiserecordsave]
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sCode nvarchar(MAX);

	SET @sCode = '';
	SELECT @sCode = @sCode + 'UPDATE dbo.[tbuser_' + [tablename] + '] SET [updflag] = 0 WHERE [id] = 0;' + CHAR(13)
		FROM ASRSysTables
		WHERE [TableType] IN (1,2)
		ORDER BY [tabletype];

	EXECUTE sp_executesql @sCode;
	
END