CREATE TABLE [dbo].[tbsys_scriptedobjects](
	[guid] [uniqueidentifier] NOT NULL,
	[parentguid] [uniqueidentifier] NULL,
	[objecttype] [int] NOT NULL,
	[targetid] [int] NULL,
	[ownerid] [uniqueidentifier] NOT NULL,
	[effectivedate] [datetime] NULL,
	[disabledate] [datetime] NULL,
	[revision] [int] NOT NULL,
	[lastupdated] [datetime] NULL,
	[lastupdatedby] [nvarchar](255) NULL,
	[locked] [bit] NOT NULL,
	[tag] [xml] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]