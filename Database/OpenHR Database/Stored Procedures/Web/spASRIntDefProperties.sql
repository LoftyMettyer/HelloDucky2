CREATE PROCEDURE [dbo].[spASRIntDefProperties] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @Name	nvarchar(255);

	-- Access details of object
	SELECT CreatedDate, SavedDate, RunDate,
		Createdby, 
		Savedby, 
		Runby 
	FROM [dbo].[ASRSysUtilAccessLog]
	WHERE UtilID = @intID AND [Type] = @intType;

	-- Get usage of this object
	EXEC sp_ASRIntDefUsage @intType, @intID;

END

