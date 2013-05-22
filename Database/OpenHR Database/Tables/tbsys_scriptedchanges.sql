CREATE TABLE [dbo].[tbsys_scriptedchanges](
	[id] [uniqueidentifier] NULL,
	[sequence] [int] NULL,
	[file] [nvarchar](max) NULL,
	[uploaddate] [datetime] NULL,
	[runtype] [int] NULL,
	[lastrundate] [datetime] NULL,
	[runonce] [bit] NULL,
	[runinversion] [nvarchar](10) NULL,
	[description] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]