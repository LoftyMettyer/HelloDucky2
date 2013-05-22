CREATE PROCEDURE sp_ASRGetRelatedTables (
	@tableID		int = 0)
AS

SELECT tableID, tableName 
FROM ASRSysTables
JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID
WHERE ASRSysRelations.parentID= @tableID
UNION
SELECT tableID, tableName 
FROM ASRSysTables
JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID
WHERE ASRSysRelations.childID= @tableID


GO

