CREATE PROCEDURE spASRIntGetMailMergeDS(@id as integer)
AS
BEGIN

	SELECT m.*, t.TableName, t.RecordDescExprID
		FROM ASRSysMailMergeName m
		JOIN ASRSYSTables t ON (t.TableID = m.TableID) WHERE MailMergeID = @id;

	SELECT 'Col' AS [ColExp],  c.ColumnID AS ColExpId
		, t.TableID AS [tableid], t.Tablename AS [Table], c.ColumnName AS [Name]
		, c.DataType AS [Type], m.Size, m.Decimals, c.Use1000Separator
	FROM ASRSysMailMergeColumns m
		JOIN ASRSysColumns c ON (c.ColumnID = m.ColumnID) 
		JOIN ASRSysTables t ON (t.TableID = c.TableID) WHERE m.Type = 'C' AND m.MailMergeID = @id
	UNION    
	SELECT 'Exp' AS [ColExp], e.ExprID AS [ColExpId],  0 AS [TableID]
		, 'Calculation_' AS [Table], e.Name AS [Name]
		, e.ReturnType as [Type], m.Size, m.Decimals, 0 AS [Use1000Separator]
	FROM ASRSysMailMergeColumns m
		LEFT OUTER JOIN ASRSysExpressions e ON (e.ExprID = m.ColumnID)
		WHERE m.Type = 'E' AND m.MailMergeID = @id
	ORDER BY [ColExp], [Table], [Name];

END
