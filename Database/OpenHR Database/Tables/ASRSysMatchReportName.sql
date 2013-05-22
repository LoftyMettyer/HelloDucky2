﻿CREATE TABLE [dbo].[ASRSysMatchReportName](
	[MatchReportID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Description] [varchar](255) NOT NULL,
	[Table1ID] [int] NOT NULL,
	[Table1AllRecords] [bit] NOT NULL,
	[Table1Picklist] [int] NOT NULL,
	[Table1Filter] [int] NOT NULL,
	[Table2ID] [int] NOT NULL,
	[Table2AllRecords] [bit] NOT NULL,
	[Table2Picklist] [int] NOT NULL,
	[Table2Filter] [int] NOT NULL,
	[UserName] [varchar](50) NOT NULL,
	[NumRecords] [int] NOT NULL,
	[OutputPreview] [bit] NOT NULL,
	[OutputFormat] [int] NOT NULL,
	[OutputScreen] [bit] NOT NULL,
	[OutputPrinter] [bit] NOT NULL,
	[OutputPrinterName] [varchar](255) NOT NULL,
	[OutputSave] [bit] NOT NULL,
	[OutputSaveExisting] [int] NOT NULL,
	[OutputEmail] [bit] NOT NULL,
	[OutputEmailAddr] [int] NOT NULL,
	[OutputEmailSubject] [varchar](255) NOT NULL,
	[OutputFilename] [varchar](255) NOT NULL,
	[Timestamp] [timestamp] NOT NULL,
	[MatchReportType] [int] NULL,
	[ScoreMode] [int] NULL,
	[ScoreCheck] [bit] NULL,
	[ScoreLimit] [int] NULL,
	[EqualGrade] [bit] NULL,
	[ReportingStructure] [bit] NULL,
	[PrintFilterHeader] [bit] NULL,
	[OutputEmailAttachAs] [varchar](255) NULL
) ON [PRIMARY]