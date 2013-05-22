CREATE TABLE [dbo].[ASRSysWorkflowStepDelegation](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[StepID] [int] NOT NULL,
	[DelegateEmail] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]