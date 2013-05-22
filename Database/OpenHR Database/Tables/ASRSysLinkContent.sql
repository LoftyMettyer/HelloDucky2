CREATE TABLE [dbo].[ASRSysLinkContent](
	[ID] [int] NULL,
	[ContentID] [int] NULL,
	[Sequence] [int] NULL,
	[FixedText] [varchar](max) NULL,
	[FieldCode] [varchar](1) NULL,
	[FieldID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]