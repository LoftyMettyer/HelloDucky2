CREATE TABLE [dbo].[ASRSysPermissionItems](
	[itemID] [int] NOT NULL,
	[description] [varchar](50) NOT NULL,
	[listOrder] [int] NOT NULL,
	[categoryID] [int] NOT NULL,
	[itemKey] [varchar](50) NOT NULL
) ON [PRIMARY]