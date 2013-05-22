CREATE TABLE [dbo].[ASRSysEmailQueue](
	[QueueID] [int] IDENTITY(1,1) NOT NULL,
	[LinkID] [int] NULL,
	[ColumnID] [int] NULL,
	[RecordID] [int] NULL,
	[DateDue] [datetime] NULL,
	[DateSent] [datetime] NULL,
	[UserName] [varchar](50) NULL,
	[RecordDesc] [varchar](255) NULL,
	[ColumnValue] [varchar](max) NULL,
	[Immediate] [bit] NULL,
	[RecalculateRecordDesc] [bit] NULL,
	[TableID] [int] NULL,
	[RepTo] [varchar](max) NULL,
	[RepCC] [varchar](max) NULL,
	[RepBCC] [varchar](4000) NULL,
	[MsgText] [varchar](max) NULL,
	[Subject] [varchar](max) NULL,
	[Attachment] [varchar](max) NULL,
	[WorkflowInstanceID] [int] NULL,
	[Type] [int] NULL,
 CONSTRAINT [PK_ASRSysEmailQueue] PRIMARY KEY CLUSTERED 
(
	[QueueID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]