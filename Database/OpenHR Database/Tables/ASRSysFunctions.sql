CREATE TABLE [dbo].[ASRSysFunctions](
	[functionID] [int] NOT NULL,
	[functionName] [varchar](255) NOT NULL,
	[returnType] [smallint] NOT NULL,
	[timeDependent] [bit] NOT NULL,
	[category] [varchar](20) NULL,
	[spName] [varchar](50) NULL,
	[nonStandard] [bit] NOT NULL,
	[runtime] [bit] NULL,
	[UDF] [bit] NULL,
	[ShortcutKeys] [varchar](20) NULL,
	[ExcludeExprTypes] [varchar](50) NULL,
	[UDFName] [nvarchar](255) NULL,
	[IncludeExprTypes] [varchar](50) NULL
) ON [PRIMARY]