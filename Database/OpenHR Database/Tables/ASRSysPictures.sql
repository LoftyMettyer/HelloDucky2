﻿CREATE TABLE [dbo].[ASRSysPictures](
	[PictureID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Picture] [varbinary](max) NULL,
	[PictureType] [smallint] NOT NULL,
	[GUID] [uniqueidentifier] NULL,
 CONSTRAINT [PK_ASRSysPictures] PRIMARY KEY CLUSTERED 
(
	[PictureID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) TEXTIMAGE_ON [PRIMARY]