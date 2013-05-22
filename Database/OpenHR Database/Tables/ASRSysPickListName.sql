CREATE TABLE [dbo].[ASRSysPickListName](
	[PickListID] [int] IDENTITY(1,1) NOT NULL,
	[Description] [varchar](255) NULL,
	[TableID] [int] NULL,
	[Access] [varchar](2) NULL,
	[UserName] [varchar](50) NULL,
	[TimeStamp] [timestamp] NULL,
	[Name] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysPickListName] ADD  DEFAULT ('') FOR [Name]