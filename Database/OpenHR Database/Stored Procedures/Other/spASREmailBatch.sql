CREATE PROCEDURE [dbo].[spASREmailBatch]
AS
BEGIN

	DECLARE @QueueID	integer,
		@LinkID			integer,
		@RecordID		integer,
		@ColumnID		integer,
		@ColumnValue	integer,
		@RecDescID		integer,
		@RecDesc		nvarchar(MAX),
		@sSQL			nvarchar(MAX),
		@EmailDate		datetime,
		@hResult		integer,
		@blnEnabled		integer;

	SELECT @blnEnabled = [SettingValue] FROM [dbo].[ASRSysSystemSettings]
		WHERE [Section] = 'email' and [SettingKey] = 'overnight enabled';

	IF @blnEnabled = 0
	BEGIN
		RETURN
	END

	-- Clear Servers Inbox
	-- Doing this just before sending messages means that any failure return messages will
	-- stay in the servers inbox until this sp is run again - could be useful for support ?

	-- DECLARE @message_id varchar(255)
	-- EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	-- WHILE not @message_ID is null
	-- BEGIN
	--	EXEC master.dbo.xp_deletemail @message_id
	--	SET @message_id = null
	--	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
	-- END


	/* Purge email queue */
	EXEC sp_ASRPurgeRecords 'EMAIL', 'ASRSysEmailQueue', 'DateDue';

	/* Send all emails waiting to be sent regardless of username */
	EXEC spASREmailImmediate '';

END