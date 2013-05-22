CREATE PROCEDURE [dbo].[spASRAccordSetLatestToType] (
		@piTransferType		integer ,
		@piHRProRecordID	integer,
		@piTransactionType	integer)
AS
BEGIN	

	SET NOCOUNT ON;
	
	DECLARE @iTransactionID integer;

	-- Get our transaction
	SELECT TOP 1 @iTransactionID = TransactionID FROM ASRSysAccordTransactions
		WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
		ORDER BY CreatedDateTime DESC

	-- Force the transaction type
	UPDATE dbo.[ASRSysAccordTransactions] SET TransactionType = @piTransactionType
		WHERE TransactionID = @iTransactionID;

	-- If new type then ensure that old data is blank
	IF @piTransactionType = 0
		UPDATE dbo.[ASRSysAccordTransactionData] SET [OldData] = '' WHERE TransactionID = @iTransactionID;

END