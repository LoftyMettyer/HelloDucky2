CREATE TABLE [dbo].[ASRSysTalentReportDetails](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[TalentReportID] [int] NOT NULL,
			[ColType] varchar(1) NOT NULL,
			[ColExprID] [int] NOT NULL,
			[ColSize] [int] NULL,
			[ColDecs] [int] NULL,
			[ColHeading] [varchar](255) NULL,
			[ColSequence] [int] NULL,
			[SortOrderSeq] [int] NULL,
			[SortOrderDirection] [varchar](4) NULL)