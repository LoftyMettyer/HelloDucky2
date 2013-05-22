CREATE PROCEDURE [dbo].[spASROutlookEventRefresh](
	@LinkID		integer,
	@FolderID	integer,
	@TableID	integer,
	@RecordID	integer)
AS
BEGIN

	IF EXISTS(SELECT * FROM ASRSysOutlookEvents WHERE LinkID = @LinkID AND FolderID = @FolderID AND TableID = @TableID AND RecordID = @RecordID)
		UPDATE ASRSysOutlookEvents SET Refresh = 1 WHERE LinkID = @LinkID AND FolderID = @FolderID AND TableID = @TableID AND RecordID = @RecordID;
	ELSE
		INSERT ASRSysOutlookEvents(LinkID, FolderID, TableID, RecordID, Refresh, Deleted) VALUES (@LinkID,@FolderID, @TableID, @RecordID, 1, 0);

END