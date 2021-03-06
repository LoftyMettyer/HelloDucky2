﻿CREATE TABLE [dbo].[ASRSysChildViews2](
	[childViewID] [int] IDENTITY(1,1) NOT NULL,
	[tableID] [int] NOT NULL,
	[type] [int] NULL,
	[role] [varchar](256) NOT NULL,
 CONSTRAINT [PK_ASRSysChildViews2] PRIMARY KEY CLUSTERED 
(
	[childViewID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)

GO

CREATE NONCLUSTERED INDEX [IDX_TableID]
    ON [dbo].[ASRSysChildViews2]([tableID] ASC);
GO