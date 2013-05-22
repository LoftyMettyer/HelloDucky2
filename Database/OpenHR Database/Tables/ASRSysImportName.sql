CREATE TABLE [dbo].[ASRSysImportName](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Description] [varchar](255) NULL,
	[BaseTable] [int] NOT NULL,
	[FileType] [int] NOT NULL,
	[FileName] [varchar](256) NOT NULL,
	[Delimiter] [varchar](7) NOT NULL,
	[Encapsulator] [char](1) NULL,
	[MultipleRecordAction] [bit] NOT NULL,
	[Username] [varchar](50) NOT NULL,
	[Timestamp] [timestamp] NULL,
	[OtherDelimiter] [char](1) NULL,
	[DateFormat] [varchar](3) NULL,
	[ImportType] [int] NULL,
	[FilterID] [int] NULL,
	[DateSeparator] [varchar](6) NULL,
	[BypassTrigger] [bit] NULL,
	[HeaderLines] [int] NULL,
	[FooterLines] [int] NULL,
 CONSTRAINT [PK_ASRSysImportName] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]