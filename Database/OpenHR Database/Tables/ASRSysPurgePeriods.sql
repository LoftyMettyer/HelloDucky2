CREATE TABLE [dbo].[ASRSysPurgePeriods](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[PurgeKey] [varchar](15) NULL,
	[Unit] [char](1) NULL,
	[Period] [int] NULL,
	[LastPurgeDate] [datetime] NULL
) ON [PRIMARY]