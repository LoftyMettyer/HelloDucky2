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


	
--	SET NOCOUNT ON;

	--DECLARE @tablesAndViews table (columnName varchar(128), columnType smallint, dataType smallint, columnID int
	--	, uniqueCheckType bit, ViewName varchar(255));

	--INSERT @tablesAndViews (columnName, columnType, dataType, columnID, uniqueCheckType, ViewName)
	--	SELECT c.columnName, c.columnType, c.dataType
	--		, c.columnID, ISNULL(c.uniqueCheckType,0)
	--		, t.tableName
	--	FROM dbo.ASRSysColumns c
	--		INNER JOIN ASRSysTables t ON c.tableID = t.tableID;

	--INSERT @tablesAndViews (columnName, columnType, dataType, columnID, uniqueCheckType, ViewName)
	--	SELECT c.columnName, c.columnType, c.dataType
	--		, c.columnID, ISNULL(c.uniqueCheckType,0)
	--		, v.viewName
	--	FROM dbo.ASRSysColumns c
	--		INNER JOIN ASRSysViews v ON c.tableID = v.viewTableID 
	--		INNER JOIN ASRSysViewColumns vc ON v.viewID = vc.viewID AND c.columnID = vc.columnID
	--	WHERE vc.inView = 1 OR c.columnType = 3;

	--SELECT tv.columnName, tv.columnType, tv.ColumnID, tv.dataType, tv.dataType, tv.uniqueCheckType
	--	,  sysobjects.name AS tableViewName, syscolumns.name AS columnName, p.action
	--	, CASE p.ProtectType WHEN 204 THEN 1 WHEN 205 THEN 1 ELSE 0 END AS permission
	--FROM #SysProtects p 
	--	INNER JOIN sysobjects ON p.id = sysobjects.id 
	--	INNER JOIN syscolumns ON p.id = syscolumns.id 
	--	INNER JOIN @tablesAndViews tv ON tv.columnname = syscolumns.name AND tv.viewName = sysobjects.name
	--WHERE (p.action = 193 or p.action = 197) 
	--	AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0 
	--	AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0) 
	--	OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0 
	--	AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
	-- ORDER BY tableViewName;



END


