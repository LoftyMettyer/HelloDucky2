CREATE TABLE [dbo].[ASRSysCalendarReportEvents](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[EventKey] [varchar](50) NOT NULL,
	[CalendarReportID] [int] NOT NULL,
	[Name] [varchar](50) NULL,
	[TableID] [int] NOT NULL,
	[FilterID] [int] NOT NULL,
	[EventStartDateID] [int] NOT NULL,
	[EventStartSessionID] [int] NULL,
	[EventEndDateID] [int] NOT NULL,
	[EventEndSessionID] [int] NULL,
	[EventDurationID] [int] NULL,
	[LegendType] [int] NULL,
	[LegendCharacter] [varchar](2) NULL,
	[LegendLookupTableID] [int] NULL,
	[LegendLookupColumnID] [int] NULL,
	[LegendLookupCodeID] [int] NULL,
	[LegendEventColumnID] [int] NULL,
	[EventDesc1ColumnID] [int] NULL,
	[EventDesc2ColumnID] [int] NULL,
 CONSTRAINT [PK_ASRSysCalendarReportEvents] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]