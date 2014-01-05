CREATE PROCEDURE dbo.spASRIntGetExpressionAndComponents
	(@ExpressionID integer, @ExpressionType integer)
AS
BEGIN

	IF @ExpressionType = 14
		SELECT name, 0 AS tableID, returnType, type, parentComponentID, Username,
			access, description, ViewInColour, CONVERT(integer, timestamp) AS intTimestamp, '' AS tableName
			FROM ASRSysExpressions
			WHERE exprID = @ExpressionID;

	ELSE
		SELECT e.name, ISNULL(e.TableID, 0) AS [tableID]
			, ISNULL(e.returnType, 0) AS [returntype]
			, ISNULL(e.type, 0) AS [type]
			, ISNULL(e.parentComponentID, 0) AS [parentComponentID]
			, ISNULL(e.Username, SYSTEM_USER) AS [username]
			, ISNULL(e.access, 'RW') AS [access]
			, ISNULL(e.description,'') AS [description]
			, ISNULL(e.ViewInColour,0) AS [ViewInColour]
			, CONVERT(integer, e.timestamp) AS [intTimestamp]
			, ISNULL(t.tableName,'') AS [tablename]
			FROM ASRSysExpressions e
				LEFT OUTER JOIN ASRSysTables t ON e.TableID = t.tableID
			WHERE exprID = @ExpressionID;

	-- Components for this expression
	SELECT * FROM ASRSysExprComponents WHERE exprID = @ExpressionID ORDER BY componentID;

END
