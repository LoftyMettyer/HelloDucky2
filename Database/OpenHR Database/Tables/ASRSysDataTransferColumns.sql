CREATE TABLE [dbo].[ASRSysDataTransferColumns](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DataTransferID] [int] NOT NULL,
	[FromTableID] [int] NOT NULL,
	[FromColumnID] [int] NOT NULL,
	[FromText] [varchar](max) NULL,
	[FromSysDate] [bit] NOT NULL,
	[ToTableID] [int] NOT NULL,
	[ToColumnID] [int] NOT NULL,
 CONSTRAINT [PK_ASRSysDataTransferColumns] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysDataTransferColumns] ADD  CONSTRAINT [DF_ASRSysDataTransferColumns_FromSysDate]  DEFAULT (0) FOR [FromSysDate]