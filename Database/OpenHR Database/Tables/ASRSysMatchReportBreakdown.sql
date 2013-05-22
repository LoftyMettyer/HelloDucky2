CREATE TABLE [dbo].[ASRSysMatchReportBreakdown](
	[MatchReportID] [int] NOT NULL,
	[MatchRelationID] [int] NOT NULL,
	[ColType] [varchar](1) NOT NULL,
	[ColExprID] [int] NOT NULL,
	[ColSize] [int] NOT NULL,
	[ColDecs] [int] NOT NULL,
	[ColHeading] [varchar](255) NOT NULL,
	[ColSequence] [int] NOT NULL
) ON [PRIMARY]