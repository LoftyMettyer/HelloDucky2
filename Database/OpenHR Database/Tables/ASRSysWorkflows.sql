CREATE TABLE [dbo].[ASRSysWorkflows](
	[id] [int] NOT NULL,
	[name] [varchar](255) NOT NULL,
	[description] [char](255) NULL,
	[enabled] [bit] NOT NULL,
	[initiationType] [smallint] NULL,
	[baseTable] [int] NULL,
	[queryString] [varchar](max) NULL,
	[PictureID] [int] NULL,
	[HasTargetIdentifier] BIT NULL, 
	[Guid] uniqueidentifier,
	[Locked] bit,
    CONSTRAINT [IDX_ID] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) TEXTIMAGE_ON [PRIMARY]