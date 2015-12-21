CREATE TABLE [dbo].[ASRSysTalentReportColumns](
			[ID] [int] IDENTITY(1,1) PRIMARY KEY NOT NULL,
			[TalentReportID] [int] NOT NULL,
			[Type] varchar(1) NOT NULL,
			[ColumnID] [int] NOT NULL,
			[sortOrder] [varchar](4) NULL,
			[SortOrderSequence] [int] NULL,
			[Size] [int] NULL,
			[Decimals] [int] NULL,
			[ColumnOrder] [int] NULL,
			[StartOnNewLine] [bit] NULL)