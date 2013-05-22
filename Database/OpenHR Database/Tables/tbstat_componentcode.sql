CREATE TABLE [dbo].[tbstat_componentcode](
	[id] [int] NOT NULL,
	[code] [nvarchar](max) NULL,
	[precode] [nvarchar](max) NULL,
	[aftercode] [nvarchar](50) NULL,
	[datatype] [int] NULL,
	[name] [nvarchar](255) NOT NULL,
	[isoperator] [bit] NULL,
	[operatortype] [tinyint] NULL,
	[recordidrequired] [bit] NULL,
	[overnightonly] [bit] NULL,
	[istimedependant] [bit] NOT NULL,
	[calculatepostaudit] [bit] NOT NULL,
	[isgetfieldfromdb] [bit] NULL,
	[isuniquecode] [bit] NULL,
	[performancerating] [int] NOT NULL,
	[maketypesafe] [bit] NOT NULL,
	[casecount] [tinyint] NULL,
	[dependsonbankholiday] [bit] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[tbstat_componentcode] ADD  DEFAULT ((0)) FOR [istimedependant]
GO
ALTER TABLE [dbo].[tbstat_componentcode] ADD  DEFAULT ((0)) FOR [calculatepostaudit]
GO
ALTER TABLE [dbo].[tbstat_componentcode] ADD  DEFAULT ((0)) FOR [performancerating]
GO
ALTER TABLE [dbo].[tbstat_componentcode] ADD  DEFAULT ((0)) FOR [maketypesafe]
GO
ALTER TABLE [dbo].[tbstat_componentcode] ADD  DEFAULT ((0)) FOR [dependsonbankholiday]