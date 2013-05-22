CREATE PROCEDURE [dbo].[spASRIntGetColumnsFromTablesAndViews]
AS
BEGIN
	SELECT c.columnName, c.columnType, c.dataType
		, c.columnID, ISNULL(c.uniqueCheckType,0) AS uniqueCheckType
		, t.tableName AS tableViewName
	FROM dbo.ASRSysColumns c
	INNER JOIN ASRSysTables t ON c.tableID = t.tableID
	UNION 
	SELECT c.columnName, c.columnType, c.dataType
		, c.columnID, ISNULL(c.uniqueCheckType,0) AS uniqueCheckType
		, v.viewName AS tableViewName 
	FROM dbo.ASRSysColumns c
	INNER JOIN ASRSysViews v ON c.tableID = v.viewTableID 
	LEFT OUTER JOIN ASRSysViewColumns vc ON (v.viewID = vc.viewID 
			AND c.columnID = vc.columnID)
	WHERE vc.inView = 1 OR c.columnType = 3 
	ORDER BY tableViewName;

END


