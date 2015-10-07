CREATE TABLE [dbo].[ASRSysAuditAccess](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DateTimeStamp] [datetime] NOT NULL,
	[UserGroup] [varchar](50) NOT NULL,
	[UserName] [varchar](50) NOT NULL,
	[ComputerName] [varchar](255) NOT NULL,
	[HRProModule] [varchar](8) NOT NULL,
	[Action] [varchar](20) NOT NULL,
 CONSTRAINT [PK_ASRSysAuditAccess] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)

GO

CREATE NONCLUSTERED INDEX [IDX_DateTimeStamp]
		ON [dbo].[ASRSysAuditAccess]([DateTimeStamp] ASC);
GO




