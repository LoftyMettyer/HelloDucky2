
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[PERSONNEL_RECORDS]') AND xtype = 'TR')
		DROP TRIGGER [dbo].[fusion_table1];

go

CREATE TRIGGER fusion_table1
   ON dbo.tbuser_Personnel_Records
   AFTER INSERT, UPDATE
AS 
BEGIN
	SET NOCOUNT ON;

	DECLARE @LocalId integer,
			@ParentID integer,
			@startingtrigger integer;

	IF TRIGGER_NESTLEVEL(OBJECT_ID('fusion_table1')) = 1  AND TRIGGER_NESTLEVEL() = 2
	BEGIN

		-- Cursor over inserted virtual table causing message to be triggered for each
		DECLARE MessageCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Id, ISNULL(ID,0) FROM inserted;	

		OPEN MessageCursor;
		
		FETCH NEXT FROM MessageCursor INTO @LocalId, @ParentID;
			
		WHILE @@FETCH_STATUS = 0 
		BEGIN 

			IF ISNULL(@ParentID,0) > 0
			BEGIN		
				EXEC fusion.[pSendMessageCheckContext] @MessageType='StaffChange', @LocalId=@ParentID
				--EXEC fusion.[pSendMessageCheckContext] @MessageType='StaffPostChange', @LocalId=@LocalId
			END

			FETCH NEXT FROM MessageCursor INTO @LocalId, @ParentID;
		END

		CLOSE MessageCursor;
		DEALLOCATE MessageCursor;
		
	END

END

EXEC sp_settriggerorder @triggername=N'fusion_table1', @order=N'Last', @stmttype=N'INSERT'
EXEC sp_settriggerorder @triggername=N'fusion_table1', @order=N'Last', @stmttype=N'UPDATE'
