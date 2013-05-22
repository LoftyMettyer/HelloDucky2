CREATE TABLE [dbo].[tbsys_systemobjects](
	[objecttype] [int] NULL,
	[tablename] [nvarchar](255) NULL,
	[viewname] [nvarchar](255) NULL,
	[description] [nvarchar](max) NULL,
	[nextid] [int] NULL,
	[allowselect] [bit] NULL,
	[allowupdate] [bit] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]