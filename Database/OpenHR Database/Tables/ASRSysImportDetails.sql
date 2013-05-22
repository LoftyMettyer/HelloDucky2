CREATE TABLE [dbo].[ASRSysImportDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ImportID] [int] NOT NULL,
	[Type] [char](1) NOT NULL,
	[TableID] [int] NULL,
	[ColExprID] [int] NOT NULL,
	[KeyField] [bit] NOT NULL,
	[Size] [int] NULL,
	[LookupEntries] [bit] NULL
) ON [PRIMARY]