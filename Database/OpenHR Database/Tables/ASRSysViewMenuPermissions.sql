CREATE TABLE [dbo].[ASRSysViewMenuPermissions](
	[TableID] [int] NOT NULL,
	[TableName] [varchar](128) NOT NULL,
	[groupName] [varchar](255) NOT NULL,
	[HideFromMenu] [bit] NOT NULL
) ON [PRIMARY]