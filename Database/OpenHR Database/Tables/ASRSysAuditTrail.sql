CREATE TABLE [dbo].[ASRSysAuditTrail](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[UserName] [varchar](255) NOT NULL,
	[DateTimeStamp] [datetime] NOT NULL,
	[RecordID] [int] NOT NULL,
	[RecordDesc] [varchar](255) NULL,
	[OldValue] [varchar](max) NULL,
	[NewValue] [varchar](max) NULL,
	[Tablename] [varchar](200) NULL,
	[Columnname] [varchar](200) NULL,
	[CMGExportDate] [datetime] NULL,
	[CMGCommitDate] [datetime] NULL,
	[ColumnID] [int] NULL,
	[Deleted] [bit] NULL,
	[tableid] INT NULL,
 CONSTRAINT [PK_ASRSysAuditTrail] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)

GO

CREATE NONCLUSTERED INDEX [IDX_CMG]
    ON [dbo].[ASRSysAuditTrail]([ColumnID] ASC, [RecordID] ASC);
GO

CREATE NONCLUSTERED INDEX [IDX_DateTimeStamp]
    ON [dbo].[ASRSysAuditTrail]([DateTimeStamp] ASC);
GO