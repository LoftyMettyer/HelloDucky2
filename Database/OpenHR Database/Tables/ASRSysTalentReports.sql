﻿CREATE TABLE [dbo].[ASRSysTalentReports](
			[ID] [int] PRIMARY KEY IDENTITY(1,1) NOT NULL,
			[Name] [varchar](50) NOT NULL DEFAULT (''''),
			[Description] [varchar](255) NULL,
			[BaseTableID] [int] NOT NULL,
			[BaseSelection] [int] NOT NULL,
			[BasePicklistID] [int] NULL,
			[BaseFilterID] [int] NULL,
			[BaseChildTableID] [int] NOT NULL,
			[BaseChildColumnID] [int] NOT NULL,
			[BasePreferredRatingColumnID] [int] NOT NULL,
			[BaseMinimumRatingColumnID] [int] NOT NULL,
			[MatchTableID] [int] NOT NULL,
			[MatchSelection] [int] NOT NULL,
			[MatchPicklistID] [int] NULL,
			[MatchFilterID] [int] NULL,
			[MatchChildTableID] [int] NOT NULL,
			[MatchChildColumnID] [int] NOT NULL,
			[MatchChildRatingColumnID] [int] NOT NULL,
			[MatchAgainstType] [int] NOT NULL,
			[IncludeUnmatched] [bit] NULL,
			[UserName] [varchar](50) NOT NULL,
			[TimeStamp] [timestamp] NULL,
			[EmailAddrID] [int] NULL,
			[EmailSubject] [varchar](max) NULL,
			[EmailAttachmentName] [varchar](max) NULL,
			[IsLabel] [bit] NULL,
			[LabelTypeID] [int] NULL,
			[PromptStart] [int] NULL,
			[OutputFormat] [int] NULL,
			[OutputScreen] [bit] NULL,
			[OutputSave] [bit] NULL,
			[OutputFilename] [varchar](255) NULL, 
      [OutputEmail] [bit] NULL, 
			[MinimumScore] INT NULL)