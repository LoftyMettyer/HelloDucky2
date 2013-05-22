CREATE TABLE [dbo].[ASRSysMessages](
	[loginName] [varchar](256) NULL,
	[message] [varchar](200) NULL,
	[spid] [int] NULL,
	[dbid] [int] NULL,
	[uid] [int] NULL,
	[loginTime] [datetime] NULL,
	[id] [int] IDENTITY(1,1) NOT NULL,
	[messageTime] [datetime] NULL,
	[messageFrom] [varchar](256) NULL,
	[messageSource] [varchar](256) NULL
) ON [PRIMARY]