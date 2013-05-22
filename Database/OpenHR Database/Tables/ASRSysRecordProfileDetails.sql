CREATE TABLE [dbo].[ASRSysRecordProfileDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[RecordProfileID] [int] NOT NULL,
	[Sequence] [int] NULL,
	[Type] [char](1) NULL,
	[ColumnID] [int] NULL,
	[Heading] [varchar](50) NULL,
	[Size] [int] NULL,
	[DP] [int] NULL,
	[IsNumeric] [bit] NULL,
	[TableID] [int] NULL
) ON [PRIMARY]