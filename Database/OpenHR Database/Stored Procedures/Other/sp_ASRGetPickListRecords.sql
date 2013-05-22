CREATE PROCEDURE [dbo].[sp_ASRGetPickListRecords] (
	@piPickListID int)
AS
BEGIN
	SELECT ASRSysPickListItems.recordID AS id
	FROM ASRSysPickListItems 
	INNER JOIN ASRSysPickListName 
		ON ASRSysPickListItems.pickListID = ASRSysPickListName.pickListID
	WHERE ASRSysPickListName.pickListID = @piPickListID;
END