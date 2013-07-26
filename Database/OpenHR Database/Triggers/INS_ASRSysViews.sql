CREATE TRIGGER [dbo].[INS_ASRSysViews]
ON [dbo].[ASRSysViews]
INSTEAD OF INSERT
AS
BEGIN

	SET NOCOUNT ON;

	-- Update objects table
	IF NOT EXISTS(SELECT [guid]
		FROM dbo.[tbsys_scriptedobjects] o
		INNER JOIN inserted i ON i.viewid = o.targetid AND o.objecttype = 3)
	BEGIN
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT NEWID(), 3, [viewid], dbo.[udfsys_getownerid](), '01/01/1900',1,0, GETDATE()
				FROM inserted;
	END

	-- Update base table								
	INSERT dbo.[tbsys_views] ([ViewID], [ViewName], [ViewDescription], [ViewTableID], [ViewSQL], [ExpressionID]) 
		SELECT [ViewID], [ViewName], [ViewDescription], [ViewTableID], [ViewSQL], [ExpressionID] FROM inserted;

END