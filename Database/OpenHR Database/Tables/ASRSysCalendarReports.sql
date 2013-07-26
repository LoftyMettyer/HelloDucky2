CREATE TABLE [dbo].[ASRSysCalendarReports](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NULL,
	[Description] [varchar](255) NULL,
	[BaseTable] [int] NULL,
	[AllRecords] [bit] NOT NULL,
	[Picklist] [int] NULL,
	[Filter] [int] NULL,
	[UserName] [varchar](50) NULL,
	[Description1] [int] NOT NULL,
	[Description2] [int] NULL,
	[Region] [int] NULL,
	[GroupByDesc] [bit] NOT NULL,
	[StartType] [int] NOT NULL,
	[FixedStart] [datetime] NULL,
	[StartFrequency] [int] NULL,
	[StartPeriod] [int] NULL,
	[EndType] [int] NOT NULL,
	[FixedEnd] [datetime] NULL,
	[EndFrequency] [int] NULL,
	[EndPeriod] [int] NULL,
	[ShowBankHolidays] [bit] NOT NULL,
	[ShowCaptions] [bit] NOT NULL,
	[ShowWeekends] [bit] NOT NULL,
	[IncludeWorkingDaysOnly] [bit] NOT NULL,
	[OutputPreview] [bit] NOT NULL,
	[OutputFormat] [int] NOT NULL,
	[OutputScreen] [bit] NOT NULL,
	[OutputPrinter] [bit] NOT NULL,
	[OutputPrinterName] [varchar](255) NOT NULL,
	[OutputSave] [bit] NOT NULL,
	[OutputSaveExisting] [int] NOT NULL,
	[OutputEmail] [bit] NOT NULL,
	[OutputEmailAddr] [int] NOT NULL,
	[OutputEmailSubject] [varchar](255) NOT NULL,
	[OutputFilename] [varchar](255) NOT NULL,
	[Timestamp] [timestamp] NULL,
	[IncludeBankHolidays] [bit] NULL,
	[PrintFilterHeader] [bit] NULL,
	[StartDateExpr] [int] NULL,
	[EndDateExpr] [int] NULL,
	[DescriptionExpr] [int] NULL,
	[OutputEmailAttachAs] [varchar](255) NULL,
	[DescriptionSeparator] [varchar](6) NULL,
	[StartOnCurrentMonth] [bit] NULL,
 CONSTRAINT [PK_ASRSysCalendarReports] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)

GO

CREATE NONCLUSTERED INDEX [IDX_BaseTableID]
		ON [dbo].[ASRSysCalendarReports]([BaseTable] ASC);
GO

CREATE TRIGGER DEL_ASRSysCalendarReports ON dbo.ASRSysCalendarReports 
FOR DELETE 
AS
BEGIN
	DELETE FROM ASRSysCalendarReportEvents WHERE ASRSysCalendarReportEvents.CalendarReportID IN (SELECT ID FROM deleted)
	DELETE FROM ASRSysCalendarReportOrder WHERE ASRSysCalendarReportOrder.CalendarReportID IN (SELECT ID FROM deleted)
	DELETE FROM ASRSysCalendarReportAccess WHERE ASRSysCalendarReportAccess.ID IN (SELECT ID FROM Deleted)
END
GO