CREATE TABLE [dbo].[ASRSysWorkflowElementValidations](
	[ID] [int] NOT NULL,
	[ElementID] [int] NOT NULL,
	[ExprID] [int] NOT NULL,
	[Type] [smallint] NOT NULL,
	[Message] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]