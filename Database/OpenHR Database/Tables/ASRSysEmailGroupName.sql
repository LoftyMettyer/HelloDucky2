CREATE TABLE [dbo].[ASRSysEmailGroupName](
	[EmailGroupID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Description] [varchar](255) NULL,
	[UserName] [varchar](50) NOT NULL,
	[Access] [varchar](2) NOT NULL,
	[TimeStamp] [timestamp] NULL
) ON [PRIMARY]