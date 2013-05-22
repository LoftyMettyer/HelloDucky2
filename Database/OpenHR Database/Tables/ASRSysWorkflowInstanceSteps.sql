CREATE TABLE [dbo].[ASRSysWorkflowInstanceSteps](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[InstanceID] [int] NOT NULL,
	[ElementID] [int] NOT NULL,
	[Status] [int] NOT NULL,
	[ActivationDateTime] [datetime] NULL,
	[Message] [varchar](max) NULL,
	[CompletionDateTime] [datetime] NULL,
	[UserEmail] [varchar](max) NULL,
	[UserName] [varchar](256) NULL,
	[DecisionFlow] [smallint] NULL,
	[CompletionCount] [int] NULL,
	[FailedCount] [int] NULL,
	[TimeoutCount] [int] NULL,
	[EmailCC] [varchar](8000) NULL,
	[HypertextLinkedSteps] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]