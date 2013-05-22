CREATE TABLE [fusion].[MessageDefinition](
	[ID] [smallint] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Description] [varchar](max) NOT NULL,
	[Version] [tinyint] NOT NULL,
	[AllowPublish] [bit] NOT NULL,
	[AllowSubscribe] [bit] NOT NULL,
	[TableID] [int] NULL,
	[StopDeletion] [bit] NOT NULL,
	[BypassValidation] [bit] NOT NULL,
 CONSTRAINT [PK_MessageCategory] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]