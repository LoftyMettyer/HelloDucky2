USE [OpenHR5]
GO

/****** Object:  Table [fusion].[Category]    Script Date: 07/30/2012 17:26:21 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [fusion].[Category](
	[ID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[TableID] [int] NULL,
 CONSTRAINT [PK_Category] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


USE [OpenHR5]
GO

/****** Object:  Table [fusion].[Element]    Script Date: 07/30/2012 17:26:27 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [fusion].[Element](
	[ID] [int] NOT NULL,
	[CategoryID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Description] [varchar](max) NULL,
	[DataType] [int] NOT NULL,
	[MinSize] [int] NULL,
	[MaxSize] [int] NULL,
	[ColumnID] [int] NULL,
	[Lookup] [bit] NOT NULL,
 CONSTRAINT [PK_Element] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [fusion].[Element]  WITH CHECK ADD  CONSTRAINT [FK_Element_Category] FOREIGN KEY([CategoryID])
REFERENCES [fusion].[Category] ([ID])
GO

ALTER TABLE [fusion].[Element] CHECK CONSTRAINT [FK_Element_Category]
GO

USE [OpenHR5]
GO

/****** Object:  Table [fusion].[Message]    Script Date: 07/30/2012 17:26:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [fusion].[Message](
	[ID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Description] [varchar](max) NULL,
	[Schema] [varbinary](max) NULL,
	[Skeleton] [nvarchar](max) NULL,
	[Version] [int] NOT NULL,
	[AllowPublish] [bit] NOT NULL,
	[AllowSubscribe] [bit] NOT NULL,
	[Publish] [bit] NOT NULL,
	[Subscribe] [bit] NOT NULL,
	[StopDeletion] [bit] NOT NULL,
	[BypassValidation] [bit] NOT NULL,
 CONSTRAINT [PK_MessageID] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

USE [OpenHR5]
GO

/****** Object:  Table [fusion].[MessageElements]    Script Date: 07/30/2012 17:26:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [fusion].[MessageElements](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[MessageID] [int] NOT NULL,
	[NodeKey] [varchar](255) NOT NULL,
	[Position] [int] NOT NULL,
	[DataType] [int] NOT NULL,
	[Nillable] [bit] NOT NULL,
	[MinOccurs] [int] NOT NULL,
	[MaxOccurs] [int] NOT NULL,
	[MinSize] [int] NULL,
	[MaxSize] [int] NULL,
	[Lookup] [bit] NOT NULL,
	[ElementID] [int] NULL,
 CONSTRAINT [PK_MessageElementID] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [fusion].[MessageElements]  WITH CHECK ADD  CONSTRAINT [FK_Message_ElementID] FOREIGN KEY([ElementID])
REFERENCES [fusion].[Element] ([ID])
GO

ALTER TABLE [fusion].[MessageElements] CHECK CONSTRAINT [FK_Message_ElementID]
GO

ALTER TABLE [fusion].[MessageElements]  WITH CHECK ADD  CONSTRAINT [FK_MessageID] FOREIGN KEY([MessageID])
REFERENCES [fusion].[Message] ([ID])
GO

ALTER TABLE [fusion].[MessageElements] CHECK CONSTRAINT [FK_MessageID]
GO

USE [OpenHR5]
GO

/****** Object:  Table [fusion].[MessageTracking]    Script Date: 07/30/2012 17:30:01 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [fusion].[MessageTracking](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[MessageType] [varchar](50) NOT NULL,
	[BusRef] [uniqueidentifier] NOT NULL,
	[LastGeneratedDate] [datetime] NULL,
	[LastProcessedDate] [datetime] NULL,
	[LastGeneratedXml] [varchar](max) NULL,
	[Username] [varchar](255) NULL,
 CONSTRAINT [PK_MessageTracking] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

