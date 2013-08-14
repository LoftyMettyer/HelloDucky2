CREATE PROCEDURE sp_ASRGetQuickEntryTables (
	@piScreenID int)
AS
BEGIN
	SELECT DISTINCT ASRSysTables.tableName, 
		ASRSysControls.tableID
	FROM ASRSysScreens 
	INNER JOIN ASRSysControls 
		ON ASRSysScreens.ScreenID = ASRSysControls.screenID 
		AND ASRSysScreens.tableID <> ASRSysControls.tableID 
	INNER JOIN ASRSysTables 
		ON ASRSysControls.tableID = ASRSysTables.tableID
	WHERE ASRSysScreens.ScreenID = @piScreenID
END



GO

