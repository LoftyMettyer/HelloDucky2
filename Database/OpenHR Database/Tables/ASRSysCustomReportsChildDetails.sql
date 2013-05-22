CREATE TABLE [dbo].[ASRSysCustomReportsChildDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CustomReportID] [int] NOT NULL,
	[ChildTable] [int] NOT NULL,
	[ChildFilter] [int] NULL,
	[ChildMaxRecords] [int] NULL,
	[ChildOrder] [int] NULL
) ON [PRIMARY]