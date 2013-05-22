CREATE TABLE [dbo].[ASRSysWorkflowElementColumns](
	[ID] [int] NOT NULL,
	[ElementID] [int] NOT NULL,
	[ColumnID] [int] NOT NULL,
	[ValueType] [int] NOT NULL,
	[Value] [varchar](max) NULL,
	[WFFormIdentifier] [varchar](200) NULL,
	[WFValueIDentifier] [varchar](200) NULL,
	[DBColumnID] [int] NULL,
	[DBRecord] [int] NULL,
	[CalcID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]