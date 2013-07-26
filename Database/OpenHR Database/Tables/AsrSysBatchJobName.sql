﻿CREATE TABLE [dbo].[AsrSysBatchJobName](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Scheduled] [bit] NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Description] [varchar](255) NULL,
	[Frequency] [int] NOT NULL,
	[Period] [varchar](1) NOT NULL,
	[StartDate] [datetime] NULL,
	[Indefinitely] [bit] NOT NULL,
	[EndDate] [datetime] NULL,
	[Weekends] [bit] NOT NULL,
	[Access] [varchar](2) NULL,
	[Username] [varchar](50) NULL,
	[LastCompleted] [datetime] NULL,
	[RunOnce] [bit] NOT NULL,
	[RoleToPrompt] [varchar](50) NULL,
	[Timestamp] [timestamp] NULL,
	[LockSpid] [int] NULL,
	[LockLoginTime] [datetime] NULL,
	[EmailFailed] [int] NULL,
	[EmailSuccess] [int] NULL,
	[IsBatch] [bit] NULL,
	[OutputPreview] [bit] NULL,
	[OutputFormat] [int] NULL,
	[OutputScreen] [bit] NULL,
	[OutputPrinter] [bit] NULL,
	[OutputPrinterName] [varchar](255) NULL,
	[OutputSave] [bit] NULL,
	[OutputSaveExisting] [int] NULL,
	[OutputEmail] [bit] NULL,
	[OutputEmailAddr] [int] NULL,
	[OutputEmailSubject] [varchar](255) NULL,
	[OutputFilename] [varchar](255) NULL,
	[OutputEmailAttachAs] [varchar](255) NULL,
	[OutputTitlePage] [varchar](255) NULL,
	[OutputReportPackTitle] [varchar](255) NULL,
	[OutputOverrideFilter] [varchar](255) NULL,
	[OutputTOC] [bit] NULL,
	[OutputCoverSheet] [bit] NULL,
	[OverrideFilterID] [int] NULL,
	[OutputRetainPivotOrChart] [bit] NULL,
	[OutputRetainCharts] [bit] NULL
 CONSTRAINT [PK_ASRSysBatchJobName_ID] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) 

GO

CREATE CLUSTERED INDEX [IDX_Name]
    ON [dbo].[AsrSysBatchJobName]([Name] ASC);
GO

CREATE TRIGGER DEL_ASRSysBatchJobName ON dbo.ASRSysBatchJobName 
		FOR DELETE AS

			DELETE FROM ASRSysBatchJobDetails WHERE BatchJobNameID IN (SELECT ID FROM Deleted)
			DELETE FROM ASRSysBatchJobAccess WHERE ID IN (SELECT ID FROM Deleted)
GO