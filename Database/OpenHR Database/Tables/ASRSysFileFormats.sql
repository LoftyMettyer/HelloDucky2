﻿CREATE TABLE [dbo].[ASRSysFileFormats](
	[ID] [int] NULL,
	[Destination] [varchar](255) NULL,
	[Description] [varchar](255) NULL,
	[Extension] [varchar](255) NULL,
	[Office2003] [int] NULL,
	[Office2007] [int] NULL,
	[Default] [bit] NULL,
	[Direction] [tinyint] NULL
) ON [PRIMARY]