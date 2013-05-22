CREATE TABLE [dbo].[ASRSysAccordTransactionWarnings](
	[TransactionID] [int] NOT NULL,
	[FieldID] [smallint] NOT NULL,
	[WarningMessage] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]