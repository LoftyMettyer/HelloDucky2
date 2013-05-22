CREATE PROCEDURE [dbo].[sp_ASRIntDefProperties] (
	@intType int, 
	@intID int
)
AS
BEGIN
	/* Return a recordset of the details with which to populate the intranet defproperties page. */
	SELECT convert(varchar, CreatedDate,103) + ' ' + convert(varchar, CreatedDate,108) as 'CreatedDate', 
		convert(varchar, SavedDate,103) + ' ' + convert(varchar, SavedDate,108) as 'SavedDate', 
		convert(varchar, RunDate,103) + ' ' + convert(varchar, RunDate,108) as 'RunDate', 
		Createdby, 
		Savedby, 
		Runby 
	FROM [dbo].[ASRSysUtilAccessLog]
	WHERE UtilID = @intID 
		AND [Type] = @intType;
END