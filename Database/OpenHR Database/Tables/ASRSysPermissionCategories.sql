CREATE TABLE [dbo].[ASRSysPermissionCategories](
	[description] [varchar](50) NOT NULL,
	[picture] [varbinary](max) NULL,
	[listOrder] [int] NOT NULL,
	[categoryKey] [varchar](50) NOT NULL,
	[categoryID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]