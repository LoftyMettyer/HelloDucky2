CREATE TABLE [dbo].[ASRSysMailMergeTemplate](
		[Id] [int] IDENTITY(1,1) NOT NULL,
		[MailMergeID] [int] NOT NULL,
		[Template] [varbinary](max) NOT NULL,
		[TemplateName] NVARCHAR(255) NOT NULL, 
		[UploadDate] [datetime] NOT NULL,
		[UploadedUser] [nvarchar](255) NOT NULL,

    CONSTRAINT [PK_ASRSysMailMergeTemplate] PRIMARY KEY CLUSTERED ([Id] ASC)) 