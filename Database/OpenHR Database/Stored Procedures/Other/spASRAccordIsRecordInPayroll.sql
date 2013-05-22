CREATE PROCEDURE [dbo].[spASRAccordIsRecordInPayroll]
	(@iRecordID int,
	@iTransferType int,
	@ProhibitDelete int OUTPUT)
AS
	BEGIN
	SET NOCOUNT ON
	
	SET @ProhibitDelete = 0

	IF EXISTS(SELECT Status FROM ASRSysAccordTransactions
				WHERE HRProRecordID = @iRecordID
					AND Status IN (10,11)
					AND TransferType = @iTransferType)
		SET @ProhibitDelete = 1

END	
GO

