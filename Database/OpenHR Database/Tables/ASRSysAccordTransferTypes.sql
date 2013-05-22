CREATE TABLE [dbo].[ASRSysAccordTransferTypes](
	[TransferTypeID] [int] NOT NULL,
	[TransferType] [nvarchar](40) NULL,
	[FilterID] [int] NULL,
	[ASRBaseTableID] [int] NULL,
	[IsVisible] [bit] NULL,
	[ForceAsUpdate] [bit] NULL,
 CONSTRAINT [PK_TransferTypeID] PRIMARY KEY NONCLUSTERED 
(
	[TransferTypeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]