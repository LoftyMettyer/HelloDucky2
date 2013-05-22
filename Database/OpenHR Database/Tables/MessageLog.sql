CREATE TABLE [fusion].[MessageLog](
	[MessageType] [varchar](50) NOT NULL,
	[MessageRef] [uniqueidentifier] NOT NULL,
	[ReceivedDate] [datetime] NOT NULL,
	[Originator] [varchar](50) NULL
) ON [PRIMARY]