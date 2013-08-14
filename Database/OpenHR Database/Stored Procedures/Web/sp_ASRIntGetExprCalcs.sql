CREATE PROCEDURE [dbo].[sp_ASRIntGetExprCalcs] (
	@piCurrentExprID	integer,
	@piBaseTableID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the calc definitions. */
	DECLARE @sUserName	sysname;

	SET @sUserName = SYSTEM_USER;

	SELECT Name + char(9) +
		convert(varchar(255), exprID) + char(9) +
		userName AS [definitionString],
		[Description]
	FROM [dbo].[ASRSysExpressions]
	WHERE ExprID <> @piCurrentExprID
		AND Type = 10
		AND TableID = @piBaseTableID
		AND parentComponentID = 0
		AND (Username = @sUserName OR access <> 'HD')
	ORDER BY name;
END