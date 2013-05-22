CREATE TABLE [dbo].[tbsys_userusage](
	[objecttype] [smallint] NULL,
	[objectid] [int] NULL,
	[username] [varchar](255) NULL,
	[lastrun] [datetime] NULL,
	[runcount] [int] NULL,
	[lastaction] [int] NULL
) ON [PRIMARY]