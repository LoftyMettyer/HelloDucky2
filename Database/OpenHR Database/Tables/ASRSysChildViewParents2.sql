﻿CREATE TABLE [dbo].[ASRSysChildViewParents2](
	[ChildViewID] [int] NOT NULL,
	[ParentType] [char](10) NOT NULL,
	[ParentID] [int] NOT NULL,
	[ParentTableID] [int] NULL
) ON [PRIMARY]

GO

CREATE NONCLUSTERED INDEX [IDX_ChildViewID]
    ON [dbo].[ASRSysChildViewParents2]([ChildViewID] ASC);
GO