CREATE TABLE [dbo].[ASRSysCalendarReportOrder](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CalendarReportID] [int] NOT NULL,
	[TableID] [int] NOT NULL,
	[ColumnID] [int] NOT NULL,
	[OrderSequence] [int] NOT NULL,
	[OrderType] [varchar](4) NOT NULL
) ON [PRIMARY]