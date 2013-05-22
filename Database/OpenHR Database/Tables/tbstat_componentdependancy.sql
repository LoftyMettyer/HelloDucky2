CREATE TABLE [dbo].[tbstat_componentdependancy](
	[id] [int] NOT NULL,
	[type] [int] NOT NULL,
	[modulekey] [nvarchar](50) NOT NULL,
	[parameterkey] [nvarchar](50) NOT NULL,
	[code] [nvarchar](max) NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]