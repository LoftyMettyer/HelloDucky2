CREATE PROCEDURE dbo.[spASRGetCalculationsForTable](@piTableID as integer)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT ExprID AS ID,
			Name,
			0 AS DataType,
			0 AS Size,
			0 AS Decimals,
			Access,
			Username
	 FROM ASRSysExpressions
		WHERE type = 10 AND (returnType = 0 OR type = 10) AND parentComponentID = 0	AND TableID  = @piTableID
		ORDER BY Name;

END
