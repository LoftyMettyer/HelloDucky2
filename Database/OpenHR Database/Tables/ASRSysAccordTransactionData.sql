CREATE TABLE [dbo].[ASRSysAccordTransactionData](
	[TransactionID] [int] NOT NULL,
	[FieldID] [smallint] NOT NULL,
	[OldData] [varchar](max) NULL,
	[NewData] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IDX_TransactionID]
    ON [dbo].[ASRSysAccordTransactionData]([TransactionID] ASC, [FieldID] ASC);
GO