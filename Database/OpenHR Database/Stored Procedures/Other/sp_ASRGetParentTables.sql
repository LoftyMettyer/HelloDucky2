CREATE PROCEDURE sp_ASRGetParentTables (
	@piTableID int = 0)
AS
BEGIN
	SELECT tableID, tableName 
	FROM ASRSysTables
	JOIN ASRSysRelations 
		ON ASRSysTables.tableID = ASRSysRelations.parentID
	WHERE ASRSysRelations.childID= @piTableID
END





GO

