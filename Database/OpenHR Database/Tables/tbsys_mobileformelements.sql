CREATE TABLE [dbo].[tbsys_mobileformelements](
	[ID] [int] NOT NULL,
	[Form] [tinyint] NOT NULL,
	[Type] [tinyint] NOT NULL,
	[Name] [varchar](50) NULL,
	[Caption] [varchar](500) NULL,
	[FontName] [varchar](255) NULL,
	[FontSize] [float] NULL,
	[FontBold] [bit] NULL,
	[FontItalic] [bit] NULL,
	[ForeColor] [int] NULL,
	[PictureID] [int] NULL,
	[FontUnderline] [bit] NULL,
	[FontStrikeout] [bit] NULL,
 CONSTRAINT [PK_tbsys_mobileformelements] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]