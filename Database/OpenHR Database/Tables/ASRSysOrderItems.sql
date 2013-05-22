CREATE TABLE [dbo].[ASRSysOrderItems](
	[OrderID] [int] NOT NULL,
	[ColumnID] [int] NOT NULL,
	[Type] [char](1) NOT NULL,
	[Sequence] [smallint] NOT NULL,
	[Ascending] [bit] NOT NULL
) ON [PRIMARY]