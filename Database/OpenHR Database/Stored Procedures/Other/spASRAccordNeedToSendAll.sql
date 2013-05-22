CREATE PROCEDURE [dbo].[spASRAccordNeedToSendAll] 
	(@iTransferType int, 
	@iRecordID int,
	@bResend bit OUTPUT)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @Status integer;

	SELECT TOP 1 @Status = [Status] FROM [dbo].[ASRSysAccordTransactions]
		WHERE [HRProRecordID] = @iRecordID AND [TransferType] = @iTransferType
		ORDER BY [CreatedDateTime] DESC;

	-- Nothing found
	IF @Status IS NULL SET @bResend = 1;

	-- Previous transaction failed
	IF @Status IN (20) SET @bResend = 0;

	--	Previous transaction went as update - should be new
	IF @Status IN (22, 23, 31) SET @bResend = 1;

	-- Pending, success, or success with warnings, blocked
	IF @Status IN (1, 10, 11, 21, 30) SET @bResend = 0;

END