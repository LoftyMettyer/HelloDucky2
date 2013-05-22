CREATE TABLE [dbo].[ASRSysUtilAccessLog](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Type] [int] NULL,
	[UtilID] [int] NULL,
	[CreatedDate] [datetime] NULL,
	[SavedDate] [datetime] NULL,
	[RunDate] [datetime] NULL,
	[CreatedBy] [varchar](50) NULL,
	[SavedBy] [varchar](50) NULL,
	[RunBy] [varchar](50) NULL,
	[CreatedHost] [varchar](50) NULL,
	[SavedHost] [varchar](50) NULL,
	[RunHost] [varchar](50) NULL
) ON [PRIMARY]