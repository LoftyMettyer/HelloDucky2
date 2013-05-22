CREATE PROCEDURE [dbo].[sp_ASRIntGetExprCalcs] (
	@piCurrentExprID	integer,
	@piBaseTableID		integer
)
AS
BEGIN
	/* Return a recordset of the calc definitions. */
	DECLARE @sUserName	sysname;

	SET @sUserName = SYSTEM_USER;

	SELECT name + char(9) +
		convert(varchar(255), exprID) + char(9) +
		userName AS [definitionString],
		[description]
	FROM [dbo].[ASRSysExpressions]
	WHERE exprID <> @piCurrentExprID
		AND type = 10
		AND tableID = @piBaseTableID
		AND parentComponentID = 0
		AND (Username = @sUserName OR access <> 'HD')
	ORDER BY name;
END