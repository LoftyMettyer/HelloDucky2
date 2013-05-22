CREATE TABLE [dbo].[ASRSysAccordTransactionProcessInfo](
	[SPID] [smallint] NOT NULL,
	[TransactionID] [numeric](18, 0) NOT NULL,
	[TransferType] [smallint] NOT NULL,
	[RecordID] [int] NOT NULL
) ON [PRIMARY]