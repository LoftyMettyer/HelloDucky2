CREATE TABLE [dbo].[ASRSysPageSizes](
	[PageSizeID] [int] NOT NULL,
	[Name] [char](20) NOT NULL,
	[Width] [float] NOT NULL,
	[Height] [float] NOT NULL,
	[DisplayOrder] [int] NOT NULL,
	[WordTemplateID] [int] NOT NULL,
	[IsEnvelope] [bit] NOT NULL,
 CONSTRAINT [PK_PageSizeID] PRIMARY KEY CLUSTERED 
(
	[PageSizeID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]