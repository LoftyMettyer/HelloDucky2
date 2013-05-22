CREATE TABLE [dbo].[ASRSysGlobalItems](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[FunctionID] [int] NOT NULL,
	[ColumnID] [int] NOT NULL,
	[ValueType] [int] NOT NULL,
	[ExprID] [int] NULL,
	[Value] [varchar](max) NULL,
	[RefColumnID] [int] NULL,
	[LookupTableID] [int] NULL,
	[LookupColumnID] [int] NULL,
 CONSTRAINT [PK_ASRSysGlobalItems] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]