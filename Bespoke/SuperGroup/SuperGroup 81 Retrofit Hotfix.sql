
/* ------------------------------------------------------- */
PRINT 'Step - Export additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'TransformFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD TransformFile nvarchar(MAX) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XMLDataNodeName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XMLDataNodeName nvarchar(50) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LastSuccessfulOutput')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LastSuccessfulOutput datetime NULL;';
		
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'AuditChangesOnly')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD AuditChangesOnly bit NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XSDFileName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XSDFileName nvarchar(255) NULL, PreserveTransformPath bit, PreserveXSDPath bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitXMLNodesFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitXMLNodesFile bit;';

	ALTER TABLE [ASRSysExportName] ALTER COLUMN [HeaderText] varchar(MAX);
	ALTER TABLE [ASRSysExportName] ALTER COLUMN [FooterText] varchar(MAX);

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'StripDelimiterFromData')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD StripDelimiterFromData bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFile bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFileSize')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFileSize int;';





