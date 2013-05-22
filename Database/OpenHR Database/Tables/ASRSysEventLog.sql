CREATE TABLE [dbo].[ASRSysEventLog](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DateTime] [datetime] NOT NULL,
	[Status] [int] NOT NULL,
	[Username] [varchar](50) NOT NULL,
	[Type] [int] NOT NULL,
	[Name] [varchar](150) NULL,
	[Mode] [bit] NOT NULL,
	[BatchName] [varchar](50) NULL,
	[SuccessCount] [int] NULL,
	[FailCount] [int] NULL,
	[BatchRunID] [int] NULL,
	[EndTime] [datetime] NULL,
	[Duration] [numeric](18, 0) NULL,
	[BatchJobID] [int] NULL,
	[ReportPack] [bit] NULL,
 CONSTRAINT [PK_ASRSysEventLog_ID] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysEventLog] ADD  DEFAULT (1) FOR [Type]
GO
ALTER TABLE [dbo].[ASRSysEventLog] ADD  DEFAULT ('') FOR [Name]
GO
ALTER TABLE [dbo].[ASRSysEventLog] ADD  DEFAULT (0) FOR [Mode]