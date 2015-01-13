CREATE TABLE [dbo].[tbsys_tables](
	[TableID] [int] NOT NULL,
	[TableType] [smallint] NOT NULL,
	[lastUpdated] [datetime] NULL,
	[DefaultOrderID] [int] NULL,
	[RecordDescExprID] [int] NULL,
	[DefaultEmailID] [int] NULL,
	[TableName] [varchar](128) NULL,
	[ManualSummaryColumnBreaks] [bit] NULL,
	[AuditInsert] [bit] NULL,
	[AuditDelete] [bit] NULL,
	[isremoteview] [bit] NULL,
 [InsertTriggerDisabled] BIT NULL, 
    [UpdateTriggerDisabled] BIT NULL, 
    [DeleteTriggerDisabled] BIT NULL, 
    [CopyWhenParentRecordIsCopied] BIT NULL, 
    CONSTRAINT [PK_ASRSysTables] PRIMARY KEY CLUSTERED 
(
	[TableID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)