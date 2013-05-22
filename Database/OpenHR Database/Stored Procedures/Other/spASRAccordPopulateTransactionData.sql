CREATE PROCEDURE [dbo].[spASRAccordPopulateTransactionData] (
		@piTransactionID int,
		@piColumnID int,
		@psOldValue varchar(255),
		@psNewValue varchar(255)
		)
	AS
	BEGIN	
		DECLARE @iRecCount int

		SET NOCOUNT ON

		SELECT @iRecCount = COUNT(FieldID) FROM ASRSysAccordTransactionData WHERE @piTransactionID = TransactionID and FieldID = @piColumnID

		-- Insert a record into the Accord Transaction table.	
		IF @iRecCount = 0
			INSERT INTO ASRSysAccordTransactionData
				([TransactionID],[FieldID], [OldData], [NewData])
			VALUES 
				(@piTransactionID,@piColumnID,@psOldValue,@psNewValue)
		ELSE
			UPDATE ASRSysAccordTransactionData SET [OldData] = @psOldValue
				WHERE @piTransactionID = TransactionID and FieldID = @piColumnID
	END
GO

