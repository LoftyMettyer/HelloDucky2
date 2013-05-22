CREATE TABLE [dbo].[ASRSysRecordProfileTables](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[RecordProfileID] [int] NOT NULL,
	[TableID] [int] NOT NULL,
	[FilterID] [int] NULL,
	[OrderID] [int] NULL,
	[MaxRecords] [int] NULL,
	[Orientation] [int] NULL,
	[PageBreak] [bit] NULL,
	[Sequence] [int] NULL
) ON [PRIMARY]