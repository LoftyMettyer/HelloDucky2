CREATE PROCEDURE spASRAccordPurgeTemp (
			@piTriggerLevel int)
	AS
	BEGIN	
	
		-- This stored procedure is called from every table trigger and resets the Accord transaction id whenever the trigger level is 1
		IF @piTriggerLevel = 1 DELETE FROM ASRSysAccordTransactionProcessInfo WHERE spid = @@SPID
	END
GO

