CREATE TABLE [dbo].[ASRSysPasswords](
	[Username] [varchar](50) NOT NULL,
	[LastChanged] [datetime] NULL,
	[ForceChange] [bit] NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysPasswords] ADD  CONSTRAINT [DF_ASRSysPasswords_ForceChange]  DEFAULT (0) FOR [ForceChange]