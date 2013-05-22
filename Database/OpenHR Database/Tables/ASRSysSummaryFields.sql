CREATE TABLE [dbo].[ASRSysSummaryFields](
	[HistoryTableID] [int] NOT NULL,
	[ParentColumnID] [int] NOT NULL,
	[Sequence] [int] NOT NULL,
	[StartOfGroup] [bit] NOT NULL,
	[ID] [int] NOT NULL,
	[StartOfColumn] [bit] NULL,
 CONSTRAINT [PK_ASRSysSummaryFields] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]