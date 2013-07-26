CREATE TABLE [dbo].[ASRSysCalendarReportAccess](
	[GroupName] [varchar](256) NOT NULL,
	[Access] [varchar](2) NOT NULL,
	[ID] [int] NOT NULL
) 

GO

CREATE CLUSTERED INDEX [IDX_ID]
		ON [dbo].[ASRSysCalendarReportAccess]([ID] ASC);
GO
