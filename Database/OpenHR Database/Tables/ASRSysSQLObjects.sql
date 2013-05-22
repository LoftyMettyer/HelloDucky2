CREATE TABLE [dbo].[ASRSysSQLObjects](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Type] [varchar](16) NOT NULL,
	[DateCreated] [datetime] NOT NULL,
	[Owner] [varchar](255) NOT NULL
) ON [PRIMARY]