CREATE TABLE [dbo].[ASRSysMatchReportDetails](
	[MatchReportID] [int] NOT NULL,
	[ColType] [char](1) NOT NULL,
	[ColExprID] [int] NOT NULL,
	[ColSize] [int] NOT NULL,
	[ColDecs] [int] NOT NULL,
	[ColHeading] [varchar](255) NOT NULL,
	[ColSequence] [int] NOT NULL,
	[SortOrderSeq] [int] NOT NULL,
	[SortOrderDirection] [varchar](4) NULL
) ON [PRIMARY]