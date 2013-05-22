﻿CREATE TABLE [dbo].[ASRSysCustomReportsDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CustomReportID] [int] NOT NULL,
	[Sequence] [int] NOT NULL,
	[Type] [char](1) NOT NULL,
	[ColExprID] [int] NOT NULL,
	[Heading] [varchar](50) NOT NULL,
	[Size] [int] NOT NULL,
	[DP] [int] NOT NULL,
	[IsNumeric] [bit] NOT NULL,
	[Avge] [bit] NOT NULL,
	[Cnt] [bit] NOT NULL,
	[Tot] [bit] NOT NULL,
	[SortOrderSequence] [int] NOT NULL,
	[SortOrder] [varchar](4) NULL,
	[Boc] [bit] NOT NULL,
	[Poc] [bit] NOT NULL,
	[Voc] [bit] NOT NULL,
	[Srv] [bit] NOT NULL,
	[Repetition] [int] NULL,
	[Hidden] [bit] NULL,
	[GroupWithNextColumn] [bit] NULL
) ON [PRIMARY]