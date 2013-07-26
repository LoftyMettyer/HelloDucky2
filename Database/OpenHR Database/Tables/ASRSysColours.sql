CREATE TABLE [dbo].[ASRSysColours](
	[ColOrder] [int] NULL,
	[ColValue] [int] NULL,
	[ColDesc] [varchar](50) NULL,
	[WordColourIndex] [int] NULL,
	[CalendarLegendColour] [bit] NULL
) 

GO

CREATE CLUSTERED INDEX [IDX_ColOrder]
		ON [dbo].[ASRSysColours]([ColOrder] ASC);
GO