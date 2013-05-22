CREATE TABLE [dbo].[ASRSysWorkflowQueue](
	[QueueID] [int] IDENTITY(1,1) NOT NULL,
	[LinkID] [int] NULL,
	[Immediate] [bit] NULL,
	[RecordID] [int] NULL,
	[DateDue] [datetime] NULL,
	[DateInitiated] [datetime] NULL,
	[RecordDesc] [varchar](255) NULL,
	[UserName] [varchar](50) NULL,
	[RecalculateRecordDesc] [bit] NULL,
	[parent1TableID] [int] NULL,
	[parent1RecordID] [int] NULL,
	[parent2TableID] [int] NULL,
	[parent2RecordID] [int] NULL,
	[InstanceID] [int] NULL,
 CONSTRAINT [IDX_QueueID] PRIMARY KEY CLUSTERED 
(
	[QueueID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]