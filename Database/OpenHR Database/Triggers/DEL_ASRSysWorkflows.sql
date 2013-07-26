CREATE TRIGGER [dbo].[DEL_ASRSysWorkflows]
ON [dbo].[ASRSysWorkflows]
INSTEAD OF DELETE
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM [tbsys_workflows] WHERE id IN (SELECT id FROM deleted);
	DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT id FROM deleted) AND objecttype = 10;

END