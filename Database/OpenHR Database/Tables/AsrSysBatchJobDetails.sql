CREATE TABLE [dbo].[AsrSysBatchJobDetails](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[BatchJobNameID] [int] NOT NULL,
	[JobType] [varchar](50) NOT NULL,
	[JobID] [int] NOT NULL,
	[Parameter] [varchar](255) NOT NULL,
	[JobOrder] [int] NOT NULL
) ON [PRIMARY]