CREATE PROCEDURE [dbo].[sp_ASRIntGetTableName] (
	@piTableID		integer,
	@psTableName	varchar(255) 	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;
	
	SELECT @psTableName = tableName
	FROM [dbo].[ASRSysTables]
	WHERE tableID = @piTableID;
END