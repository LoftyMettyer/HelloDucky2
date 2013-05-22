CREATE PROCEDURE [dbo].[spASRGetAllTableAndViewColumns] 
AS
BEGIN

	SELECT ASRSysColumns.columnName, ASRSysColumns.columnType, ASRSysColumns.dataType
	, ASRSysColumns.columnID, ASRSysColumns.uniqueCheckType, ASRSysColumns.DefaultDisplayWidth
	, ASRSysColumns.Size, ASRSysColumns.Decimals, ASRSysColumns.Use1000Separator
	, ASRSysColumns.OLEType, ASRSysTables.tableName AS tableViewName 
	FROM ASRSysColumns 
	INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID 
	UNION SELECT ASRSysColumns.columnName, ASRSysColumns.columnType, ASRSysColumns.dataType
	, ASRSysColumns.columnID, ASRSysColumns.uniqueCheckType, ASRSysColumns.DefaultDisplayWidth
	, ASRSysColumns.Size, ASRSysColumns.Decimals, ASRSysColumns.Use1000Separator
	, ASRSysColumns.OLEType, ASRSysViews.viewName AS tableViewName 
	FROM ASRSysColumns 
	INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysViews.viewTableID 
	LEFT OUTER JOIN ASRSysViewColumns ON (ASRSysViews.viewID = ASRSysViewColumns.viewID 
		AND ASRSysColumns.columnID = ASRSysViewColumns.columnID) 
	WHERE ASRSysViewColumns.inView = 1 OR ASRSysColumns.columnType = 3

END