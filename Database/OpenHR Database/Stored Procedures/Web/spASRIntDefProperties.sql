CREATE PROCEDURE [dbo].[spASRIntDefProperties] (
	@intType int, 
	@intID int
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @Name	nvarchar(255);

	-- Access details of object
	SELECT convert(varchar, CreatedDate,103) + ' ' + convert(varchar, CreatedDate,108) as [CreatedDate], 
		convert(varchar, SavedDate,103) + ' ' + convert(varchar, SavedDate,108) as [SavedDate], 
		convert(varchar, RunDate,103) + ' ' + convert(varchar, RunDate,108) as [RunDate], 
		Createdby, 
		Savedby, 
		Runby 
	FROM [dbo].[ASRSysUtilAccessLog]
	WHERE UtilID = @intID AND [Type] = @intType;

	-- Get usage of this object
	EXEC sp_ASRIntDefUsage @intType, @intID;

END

