﻿CREATE TABLE [dbo].[ASRSysMailMergeName](
	[MailMergeID] [int] IDENTITY(1,1) NOT NULL,
	[Description] [varchar](255) NULL,
	[TableID] [int] NOT NULL,
	[Selection] [int] NOT NULL,
	[PicklistID] [int] NULL,
	[FilterID] [int] NULL,
	[EmailSubject] [varchar](max) NULL,
	[TemplateFileName] [varchar](256) NOT NULL,
	[UserName] [varchar](50) NOT NULL,
	[EMailAsAttachment] [bit] NOT NULL,
	[SuppressBlanks] [bit] NOT NULL,
	[PauseBeforeMerge] [bit] NOT NULL,
	[TimeStamp] [timestamp] NULL,
	[Name] [varchar](50) NOT NULL,
	[EmailAddrID] [int] NULL,
	[EmailAttachmentName] [varchar](max) NULL,
	[IsLabel] [bit] NULL,
	[LabelTypeID] [int] NULL,
	[PromptStart] [int] NULL,
	[OutputFormat] [int] NULL,
	[OutputScreen] [bit] NULL,
	[OutputPrinter] [bit] NULL,
	[OutputPrinterName] [varchar](255) NULL,
	[OutputSave] [bit] NULL,
	[OutputFilename] [varchar](255) NULL,
	[DocumentMapID] [int] NULL,
	[ManualDocManHeader] [bit] NULL, 
    [UploadTemplate] VARBINARY(MAX) NULL, 
    [UploadTemplateName] NVARCHAR(255) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[ASRSysMailMergeName] ADD  DEFAULT ('') FOR [Name]