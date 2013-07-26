CREATE TRIGGER [dbo].[DEL_ASRSysColumns]
ON [dbo].[ASRSysColumns]
INSTEAD OF DELETE
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM [tbsys_columns] WHERE columnid IN (SELECT columnid FROM deleted);
	DELETE FROM [tbsys_scriptedobjects] WHERE targetid IN (SELECT columnid FROM deleted) AND objecttype = 2;
			
END