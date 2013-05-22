CREATE TABLE [dbo].[ASRSysScreens](
	[ScreenID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[TableID] [int] NOT NULL,
	[OrderID] [int] NOT NULL,
	[Height] [int] NOT NULL,
	[Width] [int] NOT NULL,
	[PictureID] [int] NULL,
	[FontName] [varchar](50) NULL,
	[FontSize] [smallint] NULL,
	[FontBold] [bit] NOT NULL,
	[FontItalic] [bit] NOT NULL,
	[FontStrikeThru] [bit] NOT NULL,
	[FontUnderline] [bit] NOT NULL,
	[GridX] [int] NOT NULL,
	[GridY] [int] NOT NULL,
	[AlignToGrid] [bit] NOT NULL,
	[DfltForeColour] [int] NULL,
	[DfltFontName] [varchar](50) NULL,
	[DfltFontSize] [smallint] NULL,
	[DfltFontBold] [tinyint] NULL,
	[DfltFontItalic] [tinyint] NULL,
	[QuickEntry] [bit] NOT NULL,
	[SSIntranet] [bit] NULL,
	[category] [nvarchar](255) NULL,
	[groupscreens] [bit] NULL,
	[description] [nvarchar](max) NULL,
 CONSTRAINT [PK_ASRSysScreens] PRIMARY KEY NONCLUSTERED 
(
	[ScreenID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]