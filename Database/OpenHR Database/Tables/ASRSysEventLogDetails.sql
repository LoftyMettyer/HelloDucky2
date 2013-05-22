CREATE TABLE [dbo].[ASRSysEventLogDetails](
	[Notes] [varchar](max) NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[EventLogID] [int] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysEventLogDetails] ADD  DEFAULT (0) FOR [EventLogID]