CREATE TABLE [dbo].[ASRSysWorkflowInstances](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[WorkflowID] [int] NOT NULL,
	[InitiatorID] [int] NOT NULL,
	[InitiationDateTime] [datetime] NULL,
	[CompletionDateTime] [datetime] NULL,
	[UserName] [varchar](256) NULL,
	[Status] [int] NULL,
	[parent1TableID] [int] NULL,
	[parent1RecordID] [int] NULL,
	[parent2TableID] [int] NULL,
	[parent2RecordID] [int] NULL,
	[pageno] [int] NULL
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysWorkflowInstances] ADD  CONSTRAINT [DF_ASRSysWorkflowInstances_InitiationDateTime]  DEFAULT (getdate()) FOR [InitiationDateTime]