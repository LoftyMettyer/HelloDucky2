﻿CREATE TABLE [dbo].[ASRSysExportName](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Description] [varchar](255) NULL,
	[BaseTable] [int] NULL,
	[AllRecords] [bit] NOT NULL,
	[Picklist] [int] NULL,
	[Filter] [int] NULL,
	[Parent1Table] [int] NULL,
	[Parent1Filter] [int] NULL,
	[Parent2Table] [int] NULL,
	[Parent2Filter] [int] NULL,
	[ChildTable] [int] NULL,
	[ChildFilter] [int] NULL,
	[ChildMaxRecords] [int] NULL,
	[OutputType] [char](1) NULL,
	[OutputName] [varchar](255) NULL,
	[Delimiter] [varchar](7) NULL,
	[Quotes] [bit] NOT NULL,
	[Header] [int] NULL,
	[DateFormat] [char](3) NULL,
	[DateYearDigits] [char](1) NULL,
	[UserName] [varchar](50) NULL,
	[TimeStamp] [timestamp] NULL,
	[DateSeparator] [varchar](6) NULL,
	[HeaderText] [varchar](MAX) NULL,
	[CMGExportFileCode] [varchar](10) NULL,
	[CMGExportUpdateAudit] [bit] NULL,
	[CMGExportRecordID] [int] NULL,
	[Parent1AllRecords] [bit] NULL,
	[Parent1Picklist] [int] NULL,
	[Parent2AllRecords] [bit] NULL,
	[Parent2Picklist] [int] NULL,
	[Footer] [int] NULL,
	[FooterText] [varchar](MAX) NULL,
	[AppendToFile] [bit] NULL,
	[ForceHeader] [bit] NULL,
	[OmitHeader] [bit] NULL,
	[OutputFormat] [int] NULL,
	[OutputSave] [bit] NULL,
	[OutputSaveExisting] [int] NULL,
	[OutputEmail] [bit] NULL,
	[OutputEmailAddr] [int] NULL,
	[OutputEmailSubject] [varchar](255) NULL,
	[OutputFilename] [varchar](255) NULL,
	[OutputEmailAttachAs] [varchar](255) NULL,
	[OtherDelimiter] [varchar](1) NULL, 
    [TransformFile] NVARCHAR(MAX) NULL, 
    [XMLDataNodeName] NVARCHAR(50) NULL, 
    [LastSuccessfulOutput] DATETIME NULL, 
    [AuditChangesOnly] BIT NULL, 
    [StripDelimiterFromData] BIT NULL, 
    [SplitFile] BIT NULL, 
    [SplitFileSize] INT NULL, 
    [XSDFileName] NVARCHAR(255) NULL, 
    [PreserveTransformPath] BIT NULL, 
    [PreserveXSDPath] BIT NULL, 
    [SplitXMLNodesFile] BIT NULL, 
    [LinkedServer] NVARCHAR(255) NULL, 
    [LinkedCatalog] NVARCHAR(255) NULL,
		[LinkedTable] NVARCHAR(255) NULL
) ON [PRIMARY]