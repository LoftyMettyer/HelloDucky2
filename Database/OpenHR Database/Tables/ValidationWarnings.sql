CREATE TABLE [fusion].[ValidationWarnings](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TableID] [smallint] NOT NULL,
	[RecordID] [int] NOT NULL,
	[MessageName] [varchar](255) NOT NULL,
	[ValidationMessage] [varchar](max) NULL,
	[CreatedDateTime] [datetime] NOT NULL,
 CONSTRAINT [PK_ValidationWarnings] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]