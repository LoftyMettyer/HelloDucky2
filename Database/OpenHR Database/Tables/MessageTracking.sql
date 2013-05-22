CREATE TABLE [fusion].[MessageTracking](
	[MessageType] [varchar](50) NOT NULL,
	[BusRef] [uniqueidentifier] NOT NULL,
	[LastGeneratedDate] [datetime] NULL,
	[LastProcessedDate] [datetime] NULL,
	[LastGeneratedXml] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]