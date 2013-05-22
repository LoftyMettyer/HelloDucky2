CREATE PROCEDURE [dbo].[spASRAccordVoidPreviousTransactions] (
	@piTransferType int ,
	@piHRProRecordID int)
AS
BEGIN	

	SET NOCOUNT ON;

	UPDATE ASRSysAccordTransactions SET [Status] = 31
		WHERE [HRProRecordID] = @piHRProRecordID AND [TransferType] = @piTransferType;

END