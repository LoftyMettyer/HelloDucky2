CREATE TABLE [dbo].[ASRSysMatchReportTables](
	[MatchReportID] [int] NOT NULL,
	[MatchRelationID] [int] IDENTITY(1,1) NOT NULL,
	[Table1ID] [int] NOT NULL,
	[Table2ID] [int] NOT NULL,
	[RequiredExprID] [int] NOT NULL,
	[PreferredExprID] [int] NOT NULL,
	[MatchScoreExprID] [int] NOT NULL
) ON [PRIMARY]