CREATE PROCEDURE sp_ASRGetQuickEntryTables (
	@piScreenID int)
AS
BEGIN
	SELECT DISTINCT ASRSysTables.tableName, 
		ASRSysControls.tableID
	FROM ASRSysScreens 
	INNER JOIN ASRSysControls 
		ON ASRSysScreens.screenID = ASRSysControls.screenID 
		AND ASRSysScreens.tableID <> ASRSysControls.tableID 
	INNER JOIN ASRSysTables 
		ON ASRSysControls.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.screenID = @piScreenID
END



GO

