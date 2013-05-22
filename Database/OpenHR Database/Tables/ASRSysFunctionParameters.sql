CREATE TABLE [dbo].[ASRSysFunctionParameters](
	[functionID] [int] NOT NULL,
	[parameterIndex] [smallint] NOT NULL,
	[parameterType] [smallint] NOT NULL,
	[parameterName] [varchar](255) NOT NULL
) ON [PRIMARY]