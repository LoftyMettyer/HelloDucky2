CREATE TRIGGER [dbo].[DEL_ASRSysViews]
ON [dbo].[ASRSysViews]
INSTEAD OF DELETE
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM [tbsys_views] WHERE viewid IN (SELECT viewid FROM deleted);
END