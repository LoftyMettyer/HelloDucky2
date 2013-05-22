CREATE TABLE [dbo].[ASRSysOutlookEvents](
	[LinkID] [int] NULL,
	[FolderID] [int] NULL,
	[TableID] [int] NULL,
	[RecordID] [int] NULL,
	[Refresh] [bit] NULL,
	[Deleted] [bit] NULL,
	[ErrorMessage] [varchar](max) NULL,
	[StoreID] [varchar](2000) NULL,
	[EntryID] [varchar](2000) NULL,
	[Folder] [varchar](255) NULL,
	[Subject] [varchar](255) NULL,
	[StartDate] [datetime] NULL,
	[EndDate] [datetime] NULL,
	[RefreshDate] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]