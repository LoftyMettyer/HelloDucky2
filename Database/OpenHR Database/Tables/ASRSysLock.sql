﻿CREATE TABLE [dbo].[ASRSysLock](
	[Priority] [int] NULL,
	[Description] [varchar](50) NULL,
	[Username] [varchar](50) NULL,
	[Hostname] [varchar](50) NULL,
	[Lock_Time] [datetime] NULL,
	[Login_Time] [datetime] NULL,
	[SPID] [int] NULL, 
    [Module] INT NULL, 
    [NotifyGroups] NVARCHAR(MAX) NULL
) ON [PRIMARY]