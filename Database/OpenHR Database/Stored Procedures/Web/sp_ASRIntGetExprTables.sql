CREATE PROCEDURE [dbo].[sp_ASRIntGetExprTables] (
	@piTableID	integer
)
AS
BEGIN
	/* Return a recordset of tab-delimted table definitions ;
	<table id><tab><table name><tab><table type><tab><related to base table ?><tab><is child of base table ?> */
	SELECT 
		convert(varchar(255), tableID) + char(9) +
		tableName + char(9) +
		convert(varchar(255), tableType) + char(9) +
		CASE 
			WHEN (tableID = @piTableID) OR (children.childID IS NOT null) OR (parents.parentID IS NOT null) THEN '1'
			ELSE '0'
		END + char(9) +
		CASE 
			WHEN (children.childID IS NOT null) THEN '1'
			ELSE '0'
		END AS [definitionString]
	FROM [dbo].[ASRSysTables]
	LEFT OUTER JOIN ASRSysRelations children ON	(ASRSysTables.tableid = children.childID AND children.parentID = @piTableID)
	LEFT OUTER JOIN ASRSysRelations parents ON	(ASRSysTables.tableid = parents.parentID AND parents.childID = @piTableID)
	ORDER BY tableName;
END