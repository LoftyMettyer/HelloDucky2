CREATE TABLE [dbo].[ASRSysDiaryLinks](
	[diaryID] [int] NOT NULL,
	[columnID] [int] NOT NULL,
	[comment] [varchar](255) NULL,
	[offset] [smallint] NOT NULL,
	[period] [smallint] NOT NULL,
	[reminder] [bit] NOT NULL,
	[FilterID] [int] NULL,
	[EffectiveDate] [datetime] NULL,
	[CheckLeavingDate] [bit] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysDiaryLinks] ADD  CONSTRAINT [DF_ASRSysDiaryLinks_FilterID]  DEFAULT (0) FOR [FilterID]
GO
ALTER TABLE [dbo].[ASRSysDiaryLinks] ADD  CONSTRAINT [DF_ASRSysDiaryLinks_EffectiveDate]  DEFAULT ('01/01/1980') FOR [EffectiveDate]