CREATE TABLE [dbo].[ASRSysDiaryEvents](
	[DiaryEventsID] [int] IDENTITY(1,1) NOT NULL,
	[TableID] [int] NULL,
	[ColumnID] [int] NULL,
	[RowID] [int] NULL,
	[EventDate] [datetime] NOT NULL,
	[EventTime] [char](5) NULL,
	[Alarm] [bit] NOT NULL,
	[Access] [char](2) NULL,
	[CopiedFromID] [int] NULL,
	[TimeStamp] [timestamp] NULL,
	[EventNotes] [varchar](max) NULL,
	[EventTitle] [varchar](255) NOT NULL,
	[UserName] [varchar](50) NULL,
	[ColumnValue] [datetime] NULL,
	[LinkID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysDiaryEvents] ADD  DEFAULT ('') FOR [EventTitle]