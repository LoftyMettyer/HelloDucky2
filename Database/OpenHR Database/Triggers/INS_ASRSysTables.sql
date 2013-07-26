CREATE TRIGGER [INS_ASRSysTables]
ON [dbo].[ASRSysTables]
INSTEAD OF INSERT
AS
BEGIN
	SET NOCOUNT ON;

	-- Update objects table
	IF NOT EXISTS(SELECT [guid]
		FROM dbo.[tbsys_scriptedobjects] o
		INNER JOIN inserted i ON i.tableid = o.targetid AND o.objecttype = 1)
	BEGIN
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT NEWID(), 1, [tableid], dbo.[udfsys_getownerid](), '01/01/1900',1,0, GETDATE()
				FROM inserted;
	END

	-- Update base table								
	INSERT dbo.[tbsys_tables] ([TableID], [TableType], [DefaultOrderID], [RecordDescExprID], [DefaultEmailID], [TableName], [ManualSummaryColumnBreaks], [AuditInsert], [AuditDelete], [isremoteview]) 
		SELECT [TableID], [TableType], [DefaultOrderID], [RecordDescExprID], [DefaultEmailID], [TableName], [ManualSummaryColumnBreaks], [AuditInsert], [AuditDelete], [isremoteview] FROM inserted;

END
