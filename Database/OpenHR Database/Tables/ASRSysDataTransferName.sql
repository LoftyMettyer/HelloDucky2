CREATE TABLE [dbo].[ASRSysDataTransferName](
	[DataTransferID] [int] IDENTITY(1,1) NOT NULL,
	[Description] [varchar](255) NULL,
	[FromTableID] [int] NOT NULL,
	[AllRecords] [bit] NOT NULL,
	[FilterID] [int] NOT NULL,
	[PickListID] [int] NOT NULL,
	[ToTableID] [int] NOT NULL,
	[UserName] [varchar](50) NULL,
	[TimeStamp] [timestamp] NULL,
	[Name] [varchar](50) NOT NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysDataTransferName] ADD  DEFAULT ('') FOR [Name]