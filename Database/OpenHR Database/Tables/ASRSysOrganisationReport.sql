CREATE TABLE [dbo].[ASRSysOrganisationReport](
		[ID] [int] IDENTITY(1,1) NOT NULL,
		[Name] [varchar](50) NOT NULL,
		[Description] [varchar](255) NOT NULL,
		[BaseViewID] [int] NOT NULL,
		[UserName] [varchar](50) NOT NULL,
		[Timestamp] [timestamp] NOT NULL)