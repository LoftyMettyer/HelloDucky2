CREATE TABLE [dbo].[ASRSysMailMergeColumns](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[MailMergeID] [int] NOT NULL,
	[Type] [char](1) NOT NULL,
	[ColumnID] [int] NOT NULL,
	[sortOrder] [varchar](4) NULL,
	[SortOrderSequence] [int] NULL,
	[Size] [int] NULL,
	[Decimals] [int] NULL,
	[ColumnOrder] [int] NULL,
	[StartOnNewLine] [bit] NULL,
	[Headingtext] [varchar](50) NULL,
 CONSTRAINT [PK_ASRSysMailMergeColumns] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]