CREATE TABLE [dbo].[ASRSysSSIViews](
	[ViewID] [int] NOT NULL,
	[ButtonLinkPromptText] [varchar](200) NULL,
	[ButtonLinkButtonText] [varchar](200) NULL,
	[HypertextLinkText] [varchar](200) NULL,
	[DropdownListLinkText] [varchar](200) NULL,
	[ButtonLink] [bit] NOT NULL,
	[HypertextLink] [bit] NOT NULL,
	[DropdownListLink] [bit] NOT NULL,
	[SingleRecordView] [bit] NOT NULL,
	[Sequence] [int] NOT NULL,
	[LinksLinkText] [varchar](200) NULL,
	[pageTitle] [varchar](200) NULL,
	[TableID] [int] NULL,
	[WFOutOfOffice] [bit] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysSSIViews] ADD  DEFAULT ((1)) FOR [WFOutOfOffice]