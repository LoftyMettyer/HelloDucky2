CREATE TABLE [dbo].[ASRSysAuditPermissions](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[UserName] [varchar](50) NOT NULL,
	[DateTimeStamp] [datetime] NOT NULL,
	[GroupName] [varchar](50) NOT NULL,
	[ViewTableName] [varchar](128) NULL,
	[ColumnName] [varchar](128) NULL,
	[Action] [varchar](5) NOT NULL,
	[Permission] [varchar](6) NOT NULL,
 CONSTRAINT [PK_ASRSysAuditPermissions] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]