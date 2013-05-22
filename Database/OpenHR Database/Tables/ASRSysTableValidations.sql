CREATE TABLE [dbo].[ASRSysTableValidations](
	[ValidationID] [int] NOT NULL,
	[TableID] [int] NOT NULL,
	[Type] [tinyint] NOT NULL,
	[EventStartDateColumnID] [int] NULL,
	[EventStartSessionColumnID] [int] NULL,
	[EventEndDateColumnID] [int] NULL,
	[EventEndSessionColumnID] [int] NULL,
	[FilterID] [int] NULL,
	[Severity] [tinyint] NULL,
	[Message] [nvarchar](max) NULL,
	[EventTypeColumnID] [int] NULL,
	[ColumnID] [int] NULL,
	[ValidationGUID] [uniqueidentifier] NULL,
 CONSTRAINT [PK_ASRSysTableValidations] PRIMARY KEY CLUSTERED 
(
	[ValidationID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]