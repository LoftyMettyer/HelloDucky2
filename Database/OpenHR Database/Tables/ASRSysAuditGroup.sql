CREATE TABLE [dbo].[ASRSysAuditGroup](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[UserName] [varchar](50) NOT NULL,
	[DateTimeStamp] [datetime] NOT NULL,
	[GroupName] [varchar](50) NULL,
	[UserLogin] [varchar](50) NULL,
	[Action] [varchar](20) NOT NULL,
 CONSTRAINT [PK_ASRSysAuditGroup] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]