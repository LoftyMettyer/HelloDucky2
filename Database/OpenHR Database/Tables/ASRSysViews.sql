CREATE TABLE [dbo].[ASRSysViews](
	[ViewID] [int] NOT NULL,
	[ViewName] [varchar](255) NOT NULL,
	[ViewDescription] [varchar](255) NULL,
	[ViewTableID] [int] NOT NULL,
	[ViewSQL] [varchar](255) NOT NULL,
	[ExpressionID] [int] NULL,
	[Guid] uniqueidentifier,
	[Locked] bit,
 CONSTRAINT [PK_ASRSysViews] PRIMARY KEY CLUSTERED 
(
	[ViewID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
)