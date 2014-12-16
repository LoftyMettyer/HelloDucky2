CREATE PROCEDURE spASRIntGetMailMergeDS(@id AS integer)
AS
BEGIN

	SET NOCOUNT ON;

	-- Definition
	SELECT m.*, t.TableName, t.RecordDescExprID
		FROM ASRSysMailMergeName m
		JOIN ASRSYSTables t ON (t.TableID = m.TableID) WHERE MailMergeID = @id;

	-- Columns
	SELECT 0 AS [IsExpression],  c.ColumnID AS ColExpId
		, t.TableID AS [tableid], t.Tablename AS [TableName], c.ColumnName AS [Name]
		, c.DataType AS [Type], m.Size, m.Decimals, c.Use1000Separator
	FROM ASRSysMailMergeColumns m
		INNER JOIN ASRSysColumns c ON (c.ColumnID = m.ColumnID) 
		INNER JOIN ASRSysTables t ON (t.TableID = c.TableID) WHERE m.Type = 'C' AND m.MailMergeID = @id
	UNION    
	SELECT 1 AS [IsExpression], e.ExprID AS [ColExpId],  0 AS [TableID]
		, 'Calculation_' AS [Table], e.Name AS [Name]
		, e.ReturnType as [Type], m.Size, m.Decimals, 0 AS [Use1000Separator]
	FROM ASRSysMailMergeColumns m
		LEFT OUTER JOIN ASRSysExpressions e ON (e.ExprID = m.ColumnID)
		WHERE m.Type = 'E' AND m.MailMergeID = @id
	ORDER BY [TableName], [Name];

	-- Sort Order
	SELECT t.TableID, t.TableName, c.ColumnID AS ColExpId, c.ColumnName AS [Name], mc.SortOrder 
		FROM ASRSysMailMergeColumns mc
		INNER JOIN ASRSysColumns c ON mc.ColumnID = c.ColumnID
		INNER JOIN ASRSysTables t ON c.TableID = t.TableID
		WHERE mc.MailMergeID = @id AND SortOrderSequence > 0
	ORDER BY SortOrderSequence;

END
