CREATE TABLE [dbo].[ASRSysExpressions](
	[ExprID] [int] NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[ReturnType] [int] NOT NULL,
	[ReturnSize] [int] NOT NULL,
	[ReturnDecimals] [int] NOT NULL,
	[Type] [int] NOT NULL,
	[ParentComponentID] [int] NULL,
	[Description] [varchar](255) NULL,
	[Timestamp] [timestamp] NULL,
	[TableID] [int] NOT NULL,
	[Username] [varchar](50) NULL,
	[Access] [varchar](2) NULL,
	[ExpandedNode] [bit] NULL,
	[ViewInColour] [bit] NULL,
	[UtilityID] [int] NULL,
 CONSTRAINT [PK_ASRSysExpressions] PRIMARY KEY NONCLUSTERED 
(
	[ExprID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysExpressions] ADD  CONSTRAINT [DF_ASRSysExpressions_ReturnDecimals]  DEFAULT (0) FOR [ReturnDecimals]
GO
ALTER TABLE [dbo].[ASRSysExpressions] ADD  DEFAULT (0) FOR [TableID]