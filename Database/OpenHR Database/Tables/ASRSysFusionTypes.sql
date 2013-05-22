CREATE TABLE [dbo].[ASRSysFusionTypes](
	[FusionTypeID] [int] NOT NULL,
	[FusionType] [nvarchar](40) NULL,
	[FilterID] [int] NULL,
	[ASRBaseTableID] [int] NULL,
	[IsVisible] [bit] NULL,
	[Version] [numeric](5, 2) NULL,
 CONSTRAINT [PK_FusionTypeID] PRIMARY KEY NONCLUSTERED 
(
	[FusionTypeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]