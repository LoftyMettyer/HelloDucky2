CREATE TRIGGER [dbo].[INS_ASRSysWorkflows]
ON [dbo].[ASRSysWorkflows]
INSTEAD OF INSERT
AS
BEGIN
	
	SET NOCOUNT ON;
	
	-- Update objects table
	IF NOT EXISTS(SELECT [guid]
		FROM dbo.[tbsys_scriptedobjects] o
		INNER JOIN inserted i ON i.id = o.targetid AND o.objecttype = 10)
	BEGIN
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT NEWID(), 10, [id], dbo.[udfsys_getownerid](), '01/01/1900',1,0, GETDATE()
				FROM inserted;
	END

	-- Update base table								
	INSERT dbo.[tbsys_workflows] ([id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid]) 
		SELECT [id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid] FROM inserted;

END