CREATE TABLE [dbo].[tbsys_intransactiontrigger](
	[spid] [int] NOT NULL,
	[tablefromid] [int] NOT NULL,
	[nestlevel] [int] NOT NULL,
	[actiontype] [tinyint] NOT NULL
) ON [PRIMARY]