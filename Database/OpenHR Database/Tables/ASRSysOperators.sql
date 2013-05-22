CREATE TABLE [dbo].[ASRSysOperators](
	[OperatorID] [int] NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[returnType] [smallint] NULL,
	[precedence] [smallint] NULL,
	[operandCount] [int] NULL,
	[category] [varchar](20) NULL,
	[SQLCode] [varchar](30) NULL,
	[SQLType] [varchar](1) NULL,
	[spName] [varchar](50) NULL,
	[checkDivideByZero] [bit] NOT NULL,
	[SQLFixedParam1] [varchar](20) NULL,
	[CastAsFloat] [bit] NOT NULL,
	[ShortcutKeys] [varchar](20) NULL,
 CONSTRAINT [PK_OperatorID] PRIMARY KEY CLUSTERED 
(
	[OperatorID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysOperators] ADD  DEFAULT (0) FOR [CastAsFloat]