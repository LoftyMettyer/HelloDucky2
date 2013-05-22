CREATE TABLE [dbo].[ASRSysIntPoll](
	[spid] [int] NOT NULL,
	[hitTime] [datetime] NOT NULL,
	[dbid] [int] NULL,
	[uid] [int] NULL,
	[loginTime] [datetime] NULL,
	[loginName] [varchar](256) NULL,
 CONSTRAINT [PK_ASRSysIntPoll] PRIMARY KEY CLUSTERED 
(
	[spid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]