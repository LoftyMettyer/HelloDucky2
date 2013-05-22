CREATE TABLE [dbo].[tbsys_mobilegroupworkflows](
	[UserGroupID] [int] NOT NULL,
	[WorkflowID] [int] NOT NULL,
	[Pos] [int] NOT NULL,
 CONSTRAINT [PK_tbsys_mobilegroupworkflows] PRIMARY KEY CLUSTERED 
(
	[UserGroupID] ASC,
	[WorkflowID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]