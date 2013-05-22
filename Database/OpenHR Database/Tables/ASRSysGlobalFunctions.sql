CREATE TABLE [dbo].[ASRSysGlobalFunctions](
	[FunctionID] [int] IDENTITY(1,1) NOT NULL,
	[Description] [varchar](255) NULL,
	[Type] [varchar](1) NOT NULL,
	[TableID] [int] NOT NULL,
	[ChildTableID] [int] NULL,
	[AllRecords] [bit] NOT NULL,
	[FilterID] [int] NULL,
	[PickListID] [int] NULL,
	[TimeStamp] [timestamp] NULL,
	[UserName] [varchar](50) NULL,
	[Name] [varchar](50) NOT NULL,
	[BypassTrigger] [bit] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysGlobalFunctions] ADD  DEFAULT ('') FOR [Name]