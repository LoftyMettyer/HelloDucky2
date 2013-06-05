DECLARE @fusionschemaID integer

SELECT @fusionschemaID = [SCHEMA_ID] FROM sys.schemas WHERE [name] = 'fusion'


	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spsys_getmodulesetting]')	AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_getmodulesetting];

	IF EXISTS (SELECT id FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[spsys_getaccordmodulesetting]')	AND xtype = 'P')
		DROP PROCEDURE [dbo].[spsys_getaccordmodulesetting];



	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spsys_getaccordmodulesetting](
			@moduleKey AS varchar(50),
			@parameterKey AS varchar(50),
			@paramterType AS varchar(50),			
			@parameterValue AS nvarchar(MAX) OUTPUT)
		AS
		BEGIN
			SELECT @parameterValue = [parameterValue] FROM [asrsysModuleSetup] WHERE [ModuleKey] = @moduleKey 
				AND [ParameterKey] = @parameterKey AND [ParameterType] = @paramterType
		END';


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spsys_getmodulesetting](
			@moduleKey AS varchar(50),
			@parameterKey AS varchar(50),
			@paramterType AS varchar(50),			
			@parameterValue AS nvarchar(MAX) OUTPUT)
		AS
		BEGIN
			SELECT @parameterValue = [parameterValue] FROM [asrsysModuleSetup] WHERE [ModuleKey] = @moduleKey 
				AND [ParameterKey] = @parameterKey AND [ParameterType] = @paramterType
		END';




	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageDefinition' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageDefinition]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageElements' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageElements]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageRelations' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageRelations]
		
	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'Element' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[Element]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'Message' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[Message]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'Category' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[Category]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageTracking' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageTracking]

	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageDefinition' AND type in (N'V') AND schema_id = @fusionschemaID)
		DROP VIEW [fusion].[MessageDefinition]

	-- Category
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[Category](
		[ID] [int] NOT NULL,
		[Name] [varchar](255) NOT NULL,
		[TableID] [int] NULL,
		[TranslationName] varchar(255) NOT NULL,
		[IsDataList] bit NOT NULL
		CONSTRAINT [PK_Category] PRIMARY KEY CLUSTERED ([ID] ASC))'

	-- Element
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[Element](
		[ID] [int] NOT NULL,
		[CategoryID] [int] NOT NULL,
		[Name] [varchar](255) NOT NULL,
		[Description] [varchar](max) NULL,
		[DataType] [int] NOT NULL,
		[MinSize] [int] NULL,
		[MaxSize] [int] NULL,
		[Precision] [int] NULL,
		[ColumnID] [int] NULL,
		[Lookup] [bit] NOT NULL,
		CONSTRAINT [PK_Element] PRIMARY KEY CLUSTERED ([ID] ASC))'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[Element]  WITH CHECK ADD  CONSTRAINT [FK_Element_Category] FOREIGN KEY([CategoryID])
		REFERENCES [fusion].[Category] ([ID])'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[Element] CHECK CONSTRAINT [FK_Element_Category]'

	-- Message
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[Message](
		[ID] [int] NOT NULL,
		[Name] [varchar](255) NOT NULL,
		[Description] [varchar](max) NULL,
		[Schema] [varbinary](max) NULL,
		[Version] [int] NOT NULL,
		[AllowPublish] [bit] NOT NULL,
		[AllowSubscribe] [bit] NOT NULL,
		[Publish] [bit] NOT NULL,
		[Subscribe] [bit] NOT NULL,
		[StopDeletion] [bit] NOT NULL,
		[BypassValidation] [bit] NOT NULL,
		[xmlns] nvarchar(255) NULL,
		[xmlschemaLocation] nvarchar(255) NULL,
		[xmlxsi] nvarchar(255),
		[DataNodeKey] varchar(255)
		CONSTRAINT [PK_MessageID] PRIMARY KEY CLUSTERED ([ID] ASC))'

	-- MessageElements
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[MessageElements](
		[ID] [int] IDENTITY(1,1) NOT NULL,
		[MessageID] [int] NOT NULL,
		[NodeKey] [varchar](255) NOT NULL,
		[Position] [int] NULL,
		[DataType] [int] NULL,
		[Nillable] [bit] NOT NULL,
		[MinOccurs] [int] NOT NULL,
		[MaxOccurs] [int] NOT NULL,
		[MinSize] [int] NULL,
		[MaxSize] [int] NULL,
		[Lookup] [bit] NULL,
		[ElementID] [int] NULL
		CONSTRAINT [PK_MessageElementID] PRIMARY KEY CLUSTERED ([ID] ASC))'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageElements] WITH CHECK ADD CONSTRAINT [FK_Message_ElementID] FOREIGN KEY([ElementID])
		REFERENCES [fusion].[Element] ([ID])'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageElements] CHECK CONSTRAINT [FK_Message_ElementID]'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageElements]  WITH CHECK ADD  CONSTRAINT [FK_MessageID] FOREIGN KEY([MessageID])
		REFERENCES [fusion].[Message] ([ID])'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageElements] CHECK CONSTRAINT [FK_MessageID]'

	-- MessageRelations
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[MessageRelations](
		[ID]			int IDENTITY(1,1) NOT NULL,
		[MessageID]		int NOT NULL,
		[NodeKey]		varchar(255) NOT NULL,
		[CategoryID]	int NOT NULL,
		[IsPrimaryKey]	bit NOT NULL
		CONSTRAINT [PK_MessageCategoryID] PRIMARY KEY CLUSTERED ([ID] ASC))'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageRelations] WITH CHECK ADD CONSTRAINT [FK_Message_CategoryID] FOREIGN KEY([CategoryID])
		REFERENCES [fusion].[Category] ([ID])'

	EXEC sp_executeSQL N'ALTER TABLE [fusion].[MessageRelations] CHECK CONSTRAINT [FK_Message_CategoryID]'


	-- MessageTracking
	EXEC sp_executeSQL N'CREATE TABLE [fusion].[MessageTracking](
		[ID] [int] IDENTITY(1,1) NOT NULL,
		[MessageType] [varchar](50) NOT NULL,
		[BusRef] [uniqueidentifier] NOT NULL,
		[LastGeneratedDate] [datetime] NULL,
		[LastProcessedDate] [datetime] NULL,
		[LastGeneratedXml] [varchar](max) NULL,
		[Username] [varchar](255) NULL,
		CONSTRAINT [PK_MessageTracking] PRIMARY KEY CLUSTERED ([ID] ASC))'

	-- MessageDefinition
	EXEC sp_executeSQL N'CREATE VIEW fusion.MessageDefinition
	AS
		SELECT m.name AS xmlmessageID,
			me.NodeKey AS xmlnodekey,
			me.Position,
			me.Nillable AS nilable,
			me.minOccurs,
			me.maxOccurs,
			ISNULL(c.TableID, 0) AS TableID,
			ISNULL(e.ColumnID, 0) AS ColumnID,
			e.DataType,
			me.MinSize,
			me.MaxSize,
			'''' AS value
			FROM fusion.[MessageElements] me
				INNER JOIN fusion.Message m ON m.ID = me.MessageID
				INNER JOIN fusion.Element e ON e.ID = me.ElementID
				INNER JOIN fusion.Category c ON c.ID = e.categoryID'


	-- Functions and procedures that we created in v5.0
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spGetMessageDefinitions]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spGetMessageDefinitions];

	EXECUTE sp_executesql N'CREATE PROCEDURE fusion.[spGetMessageDefinitions]
	AS
	BEGIN
		SELECT [ID], [name], [description],
			[version], [allowpublish], [allowsubscribe], [bypassvalidation], [stopdeletion],
			0 AS [tableid]
			
		 FROM fusion.[Message]
	END';



/*
select * from fusion.[Message]
select * from fusion.[Element]
select * from fusion.[MessageElements]
select * from fusion.category
*/



