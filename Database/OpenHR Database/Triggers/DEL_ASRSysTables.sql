CREATE TRIGGER [DEL_ASRSysTables]
ON [dbo].[ASRSysTables]
INSTEAD OF DELETE
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM [tbsys_tables] WHERE tableid IN (SELECT tableid FROM deleted);
	DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT tableid FROM deleted) AND objecttype = 1;

END
