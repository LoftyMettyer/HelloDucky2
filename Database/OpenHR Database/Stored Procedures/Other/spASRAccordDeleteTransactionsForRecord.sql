CREATE PROCEDURE [dbo].[spASRAccordDeleteTransactionsForRecord]
	(@iRecordID int
	, @iTransferType int)
	AS
	BEGIN
		SET NOCOUNT ON
		DELETE FROM ASRSysAccordTransactions WHERE HrProRecordID = @iRecordID
			AND TransferType = @iTransferType
	END
GO

