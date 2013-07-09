﻿CREATE TABLE [dbo].[tbsys_mobileformlayout](
	[ID] [int] NOT NULL,
	[HeaderBackColor] [int] NOT NULL,
	[HeaderPictureID] [int] NULL,
	[HeaderPictureLocation] [tinyint] NOT NULL,
	[HeaderLogoID] [int] NULL,
	[HeaderLogoWidth] [int] NOT NULL,
	[HeaderLogoHeight] [int] NOT NULL,
	[HeaderLogoHorizontalOffset] [int] NOT NULL,
	[HeaderLogoVerticalOffset] [int] NOT NULL,
	[HeaderLogoHorizontalOffsetBehaviour] [tinyint] NOT NULL,
	[HeaderLogoVerticalOffsetBehaviour] [tinyint] NOT NULL,
	[MainBackColor] [int] NOT NULL,
	[MainPictureID] [int] NULL,
	[MainPictureLocation] [tinyint] NOT NULL,
	[FooterBackColor] [int] NOT NULL,
	[FooterPictureID] [int] NULL,
	[FooterPictureLocation] [tinyint] NOT NULL,
	[TodoTitleFontName] [varchar](255) NOT NULL,
	[TodoTitleFontSize] [float] NOT NULL,
	[TodoTitleFontBold] [bit] NOT NULL,
	[TodoTitleFontItalic] [bit] NOT NULL,
	[TodoDescFontName] [varchar](255) NOT NULL,
	[TodoDescFontSize] [float] NOT NULL,
	[TodoDescFontBold] [bit] NOT NULL,
	[TodoDescFontItalic] [bit] NOT NULL,
	[HomeItemFontName] [varchar](255) NOT NULL,
	[HomeItemFontSize] [float] NOT NULL,
	[HomeItemFontBold] [bit] NOT NULL,
	[HomeItemFontItalic] [bit] NOT NULL,
	[TodoTitleForeColor] [int] NULL,
	[TodoDescForeColor] [int] NULL,
	[HomeItemForeColor] [int] NULL,
	[TodoTitleFontUnderline] [bit] NULL,
	[TodoTitleFontStrikeout] [bit] NULL,
	[TodoDescFontUnderline] [bit] NULL,
	[TodoDescFontStrikeout] [bit] NULL,
	[HomeItemFontUnderline] [bit] NULL,
	[HomeItemFontStrikeout] [bit] NULL,
 CONSTRAINT [PK_tbsys_mobileformlayout] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)