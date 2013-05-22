CREATE PROCEDURE [dbo].[spASRAccordPopulateTransaction] (
	@piTransactionID	integer OUTPUT,
	@piTransferType		integer,
	@piTransactionType	integer ,
	@piDefaultStatus	integer,
	@piHRProRecordID	integer,
	@iTriggerLevel		integer,
	@pbSendAllFields	bit OUTPUT)
AS
BEGIN	

	-- Return the required user or system setting.
	DECLARE @iCount			integer,
		@bNewTransaction	bit,
		@iStatus			integer,
		@bCreate			bit,
		@bForceAsUpdate		bit;

	SET @piTransactionID = null;
	SET @bCreate = 1;
	SET @bForceAsUpdate = 0;

	SELECT @piTransactionID = [TransactionID]
		FROM [dbo].[ASRSysAccordTransactionProcessInfo]
		WHERE [spid] = @@SPID AND [TransferType] = @piTransferType
			AND [RecordID] = @piHRProRecordID;

	-- Could be a null if the trigger was fired from a non Accord module enabled table, e.g. a child updating a parent field
	IF @piTransactionID IS null SET @bNewTransaction = 1;
	ELSE SET @bNewTransaction = 0;

	-- Get a transaction ID for this process and update the temporary Accord table
	IF @bNewTransaction = 1
	BEGIN
		SELECT @iCount = COUNT(*)
			FROM [dbo].[ASRSysSystemSettings]
			WHERE [section] = 'AccordTransfer' AND [settingKey] = 'NextTransactionID';
		
		IF @iCount = 0
			INSERT [dbo].[ASRSysSystemSettings] (Section, SettingKey, SettingValue)
				VALUES ('AccordTransfer','NextTransactionID',1);
		ELSE
			UPDATE [dbo].[ASRSysSystemSettings] SET [SettingValue] = [SettingValue] + 1
				WHERE [section] = 'AccordTransfer' AND [settingKey] =  'NextTransactionID';

		SELECT @piTransactionID = [settingValue]
			FROM[dbo].[ASRSysSystemSettings]
			WHERE [section] = 'AccordTransfer' AND [settingKey] =  'NextTransactionID';

		-- If update, has it already been sent?
		IF @piTransactionType = 1
		BEGIN

			SELECT TOP 1 @iStatus = [Status]
			FROM [dbo].[ASRSysAccordTransactions]
			WHERE [HRProRecordID] = @piHRProRecordID
				AND [TransferType] = @piTransferType
			ORDER BY [CreatedDateTime] DESC;

			IF @iStatus IS NULL OR @iStatus = 23
			BEGIN
				SET @piTransactionType = 0;
				SET @pbSendAllFields = 1;
			END
			ELSE IF @iStatus = 20
			BEGIN
				IF EXISTS(SELECT [Status]
					FROM [dbo].[ASRSysAccordTransactions]
					WHERE [HRProRecordID] = @piHRProRecordID
						AND [Status] IN (10, 11) AND [TransferType] = @piTransferType)
				BEGIN
					SET @piTransactionType = 1;
				END
				ELSE
				BEGIN
					SET @piTransactionType = 0;
				END
				
				SET @pbSendAllFields = 1;
				
			END
			
		END

		SELECT @bForceAsUpdate = [ForceAsUpdate] FROM [dbo].[ASRSysAccordTransferTypes]
			WHERE [TransferTypeID] = @piTransferType;

		IF @bForceAsUpdate = 1 AND @piTransactionType = 0 SET @piTransactionType = 1;

		-- Are we trying to delete something thats never been sent?
		IF @piTransactionType = 2
		BEGIN
			SELECT TOP 1 @iStatus = [Status] FROM [dbo].[ASRSysAccordTransactions]
			WHERE [HRProRecordID] = @piHRProRecordID AND [TransferType] = @piTransferType
			ORDER BY [CreatedDateTime] DESC;
		
			IF @iStatus IS NULL	SET @bCreate = 0;
			ELSE SET @pbSendAllFields = 1;
		END

		-- Insert a record into the Accord Transfer table.
		IF @bCreate = 1
		BEGIN
			INSERT INTO [dbo].[ASRSysAccordTransactions] ([TransactionID], [TransferType], [TransactionType], [CreatedUser], [CreatedDateTime], [Status], [HRProRecordID], [Archived])
				VALUES (@piTransactionID, @piTransferType, @piTransactionType, SYSTEM_USER, GETDATE(), @piDefaultStatus, @piHRProRecordID, 0);

			INSERT [dbo].[ASRSysAccordTransactionProcessInfo] ([SPID], [TransactionID], [TransferType], [RecordID])
				VALUES (@@SPID, @piTransactionID, @piTransferType, @piHRProRecordID);
		END

	END
END