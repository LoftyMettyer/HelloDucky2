CREATE TABLE [dbo].[ASRSysWorkflowLinks](
	[ID] [int] NOT NULL,
	[WorkflowID] [int] NOT NULL,
	[StartElementID] [int] NOT NULL,
	[EndElementID] [int] NOT NULL,
	[StartOutboundFlowCode] [int] NULL
) ON [PRIMARY]