CREATE PROCEDURE [dbo].[sp_ASRGetHistoryScreens]
	(@piParentScreenID	integer)
AS
BEGIN
	/* Return a recordset of the history screens that hang off the given parent screen. */
	SELECT ASRSysTables.tableName, 
		ASRSysTables.tableID,
		childScreens.screenID,
		childScreens.name,
		childScreens.pictureID
	FROM ASRSysScreens parentScreen
	INNER JOIN ASRSysHistoryScreens 
		ON parentScreen.screenID = ASRSysHistoryScreens.parentScreenID
	INNER JOIN ASRSysScreens childScreens 
		ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID
	INNER JOIN ASRSysTables 
		ON childScreens.tableID = ASRSysTables.tableID
	WHERE parentScreen.screenID = @piParentScreenID
		AND childScreens.quickEntry = 0;
END