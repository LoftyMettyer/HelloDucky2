CREATE TABLE [dbo].[ASRSysEmailLinks](
	[LinkID] [int] NOT NULL,
	[Title] [varchar](255) NULL,
	[FilterID] [int] NULL,
	[EffectiveDate] [datetime] NULL,
	[Attachment] [varchar](255) NULL,
	[Type] [int] NULL,
	[TableID] [int] NULL,
	[DateColumnID] [int] NULL,
	[DateOffset] [int] NULL,
	[DatePeriod] [int] NULL,
	[RecordInsert] [bit] NULL,
	[RecordUpdate] [bit] NULL,
	[RecordDelete] [bit] NULL,
	[SubjectContentID] [int] NULL,
	[BodyContentID] [int] NULL,
	[DateAmendment] [bit] NULL,
 CONSTRAINT [PK_ASREmailLinks] PRIMARY KEY CLUSTERED 
(
	[LinkID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]