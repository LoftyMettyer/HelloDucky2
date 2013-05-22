CREATE TABLE [dbo].[ASRSysExportDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[ExportID] [int] NOT NULL,
	[Type] [char](1) NOT NULL,
	[TableID] [char](10) NULL,
	[ColExprID] [int] NOT NULL,
	[Data] [varchar](max) NULL,
	[FillerLength] [int] NOT NULL,
	[SortOrderSequence] [int] NOT NULL,
	[SortOrder] [varchar](4) NULL,
	[CMGColumnCode] [varchar](50) NULL,
	[Decimals] [int] NULL,
	[Heading] [varchar](50) NULL,
	[ConvertCase] [smallint] NULL,
	[SuppressNulls] [bit] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]