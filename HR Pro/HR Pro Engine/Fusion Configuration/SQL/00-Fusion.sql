	DECLARE @fusionschemaID integer,
			@columnID integer,
			@tableID integer;

	SELECT @fusionschemaID = [SCHEMA_ID] FROM sys.schemas WHERE [name] = 'fusion'


	-- Drop the stored procedures
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessage]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessage];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessageCheckContext]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessageCheckContext];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastProcessedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastMessageDates]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogCheck]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogCheck];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogAdd]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogAdd];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateSetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateSetBusRef];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetLocalId]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetLocalId];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetBusRef];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pConvertData]') AND xtype = 'P')
		DROP FUNCTION [fusion].[pConvertData]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSetFusionContext]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[pSetFusionContext]

	-- Drop views
	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageDefinition' AND type in (N'V') AND schema_id = @fusionschemaID)
		DROP VIEW [fusion].[MessageDefinition]




	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageDefinition' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageDefinition]

	-- Category table
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'Category' AND type in (N'U') AND schema_id = @fusionschemaID)
	BEGIN

		EXEC sp_executeSQL N'CREATE TABLE [fusion].[Category](
			[ID] [int] NOT NULL,
			[Name] [varchar](255) NOT NULL,
			[TableID] [int] NULL,
			[TranslationName] varchar(255) NOT NULL,
			[IsDataList] bit NOT NULL
			CONSTRAINT [PK_Category] PRIMARY KEY CLUSTERED ([ID] ASC))'

		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_PERSONNEL', 'Param_TablePersonnel', 'PType_TableID', @tableID OUTPUT;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (0, N'Employee', @tableID, 'staff', 0);

		SELECT @tableID = NULLIF(ISNULL(ASRBaseTableID, NULL),0) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 1;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (1, N'Salary', @tableID, 'salary', 0);

		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_PERSONNEL', 'Param_FieldsHWorkingPatternTable', 'PType_TableID', @tableID OUTPUT;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (2, N'Working History (Contracts)', @tableID, 'contract', 0);

		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_ABSENCE', 'Param_TableAbsence', 'PType_TableID', @tableID OUTPUT;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (3, N'Absence', @tableID, 'absence', 0);

		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_TRAININGBOOKING', 'Param_TrainBookTable', 'PType_TableID', @tableID OUTPUT;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (4, N'Training', @tableID, 'training', 0);

		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_TRAININGBOOKING', 'Param_CourseTable', 'PType_TableID', @tableID OUTPUT;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (5, N'Courses', @tableID, 'course', 0);

		SET @tableID = NULL;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (6, N'Documents', @tableID, 'document', 0);
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (7, N'Timesheets', @tableID, 'timesheet', 0);
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (8, N'Contacts', @tableID, 'contact', 0);

		SELECT @tableID = NULLIF(c.lookupTableID,0) FROM asrsysAccordtransferfielddefinitions a
			INNER JOIN ASRSysColumns c ON c.ColumnID = a.ASRColumnID
			WHERE a.[transfertypeid] = 0 AND a.[transferfieldid] = 4
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (9, N'Title', @tableID, 'persontitle', 0);

		SET @tableID = NULL;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (10, N'Contact Type', @tableID, 'personcontacttype', 1);

		SELECT @tableID = NULLIF(c.lookupTableID,0) FROM ASRSysColumns c
			INNER JOIN ASRSysModuleSetup m ON m.parametervalue = c.columnID
			WHERE m.moduleKey = 'MODULE_PERSONNEL' AND m.ParameterKey = 'Param_FieldsDepartment' AND m.ParameterType = 'PType_ColumnID'
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (11, N'Department', @tableID, 'staffdepartment', 1);

		SELECT @tableID = NULLIF(ISNULL(ASRBaseTableID, NULL),0) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 102;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (12, N'Marital Status', @tableID, 'personmarital', 1);

		SELECT @tableID = NULLIF(ISNULL(ASRBaseTableID, NULL),0) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 112;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (13, N'Ethnicity', @tableID, 'personethnicity', 1);

		SELECT @tableID = NULLIF(ISNULL(ASRBaseTableID, NULL),0) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 101;
		INSERT [fusion].[Category] ([ID], [Name], [TableID], [TranslationName], [IsDataList]) VALUES (14, N'Leaving Reason', @tableID, 'staffleavingreason', 1);





	END

	-- Element table
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'Element' AND type in (N'U') AND schema_id = @fusionschemaID)
	BEGIN

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


		-- Staff elements
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 0
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (1, 0, N'Company Code', '', 12, NULL, NULL, 1, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsEmployeeNumber' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (2, 0, N'Staff Number', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsSurname' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (3, 0, N'Surname', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsForename' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (4, 0, N'Forenames', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 4
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (5, 0, N'Title', '', 12, NULL, NULL, 1, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (6, 0, N'NI Number', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 6
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (7, 0, N'Date of Birth', '', 11, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 7
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (8, 0, N'Gender', '', 12, NULL, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 8
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (9, 0, N'Marital Status', '', 12, NULL, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 9
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (10, 0, N'Address Line 1', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 10
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (11, 0, N'Address Line 2', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 11
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (12, 0, N'Address Line 3', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 12
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (13, 0, N'Address Line 4', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 13
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (14, 0, N'Address Line 5', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 14
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (15, 0, N'Post Code', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 15
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (16, 0, N'Telephone (Home)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 16
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (17, 0, N'Mobile (Personal)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 17
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (18, 0, N'Text Payment Advice', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 18
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (19, 0, N'Email (Work)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 19
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (20, 0, N'Email Payslip', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsStartDate' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (21, 0, N'Start Date', '', 11, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsLeavingDate' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (22, 0, N'Leaving Date', '', 11, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (23, 0, N'Leaving Reason', '', 12, NULL, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 18
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (24, 0, N'Email (Personal)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 15
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (25, 0, N'Telephone (Work)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 16
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (26, 0, N'Mobile (Work)', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsJobTitle' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (27, 0, N'Job Title', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 57
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (28, 0, N'Known As', '', 12, NULL, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 81
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (29, 0, N'Employment Type', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsHRegion' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (30, 0, N'Region', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (97, 0, N'Outreach Area Name', '', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (31, 0, N'Manager Staff Number', '', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (39, 0, N'Picture', '', -3, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (40, 0, N'Holiday Start Period', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (41, 0, N'Holiday Remaining (Hours)', '', 2, 1, 6, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (42, 0, N'Holiday Taken (Hours)', '', 2, 1, 6, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (43, 0, N'Holiday Entitlement (Hours)', '', 2, 1, 6, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (44, 0, N'Ethnicity', '', 12, 1, 6, 1, @columnID)
		EXEC dbo.spsys_getaccordmodulesetting 'MODULE_PERSONNEL', 'Param_FieldsDepartment', 'PType_ColumnID', @columnID OUTPUT;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (77, 0, N'Department', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (119, 0, N'Nationality', '', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (109, 0, N'Actual Hours', '', 2, 1, 4, 2, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (110, 0, N'Continuous Service Date', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (111, 0, N'Full time / Part Time', '', 12, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (112, 0, N'Director Start Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 23
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (113, 0, N'Payment Method', '', 12, 1, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 63
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (114, 0, N'Hours Per Day', '', 2, 1, 4, 2, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 64
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (115, 0, N'Hours per Month', '', 2, 1, 4, 2, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 61
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (116, 0, N'P11D', '', -7, 1, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 83
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (117, 0, N'OSP Contract Type', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 0 AND [transferfieldid] = 55
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (118, 0, N'Director', '', -7, 1, NULL, 0, @columnID)

		-- Appointment/Working Pattern/Contract elements
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (32, 2, N'Contract Name', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 65 AND [transferfieldid] = 3
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (33, 2, N'From Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 65 AND [transferfieldid] = 15
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (34, 2, N'To Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (35, 2, N'Department', '', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (36, 2, N'Primary Site', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnid = ASRColumnID FROM asrsysAccordtransferfielddefinitions WHERE [transfertypeid] = 65 AND [transferfieldid] = 12
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (37, 2, N'Contracted Hours (Week)', '', 2, 1, NULL, 0, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (38, 2, N'Maximum Hours (Week)', '', 2, 1, NULL, 0, @columnID)

		-- Training Booking elements
		SELECT @columnID =  NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_TrainBookCourseTitle' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (45, 4, N'Course Name', '', 12, 1, NULL, 1, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (46, 4, N'Start Date', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (47, 4, N'End Date', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (48, 4, N'Valid From', 'Date from which this training becomes valid', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (64, 4, N'Valid To', 'Date to which this training remains valid', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (49, 4, N'Reference', '', 12, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (50, 4, N'Outcome', 'Outcome of the course', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (51, 4, N'Requested By', '', 12, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (52, 4, N'Requested Date', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (53, 4, N'Accepted By', '', 12, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (54, 4, N'Accepted Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_TrainBookStatus' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (55, 4, N'Booking Status', 'Status of this delegate booking', 12, 1, NULL, 1, @columnID)
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (63, 4, N'Attended', 'Did the delegate attend this course?', -7, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (78, 4, N'Outcome Name', '', 12, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (79, 4, N'Outcome Value', '', 12, NULL, NULL, 1, @columnID)

		-- Courses
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseCancelDate' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (56, 5, N'Cancelation Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseCancelledBy' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (57, 5, N'Cancelled By', '', 12, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseEndDate' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (58, 5, N'End Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseMaxNumber' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (59, 5, N'Maximum Delegates', '', 2, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseNumberBooked' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (60, 5, N'Delegates Booked', '', 2, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseStartDate' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (61, 5, N'Start Date', '', 11, 1, NULL, 0, @columnID)
		SELECT @columnID = NULLIF(parametervalue,'') FROM asrsysModuleSetup WHERE ModuleKey = 'MODULE_TRAININGBOOKING' AND ParameterKey = 'Param_CourseTitle' AND parametertype = 'PType_ColumnID';
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (62, 5, N'Course Title', '', 12, 1, NULL, 1, @columnID)

		-- Documents
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (65, 6, N'Document name', '', 12, 1, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (66, 6, N'Valid From', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (67, 6, N'Valid To', '', 11, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (68, 6, N'Document Reference', '', 12, 1, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (69, 6, N'Secondary Document Reference', '', 12, 1, NULL, 0, @columnID)
		
		-- Timesheet
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (70, 7, N'timesheet Type', '', 12, NULL, NULL, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (71, 7, N'Timesheet Date', '', 11, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (72, 7, N'Hours Planned', '', 2, NULL, 6, 2, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (73, 7, N'Hours Worked', '', 2, NULL, 6, 2, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (74, 7, N'Hours Accured in TOIL', '', 2, NULL, 6, 2, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (75, 7, N'Holiday Taken (Hours)', '', 2, NULL, 6, 2, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (76, 7, N'Hours taken in TOIL', '', 2, NULL, 6, 2, 0, @columnID)
		
		-- Contacts
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (80, 8, N'Title', '', 12, NULL, NULL, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (81, 8, N'Forenames', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (82, 8, N'Surname', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (83, 8, N'Description', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (84, 8, N'Relationship', '', 12, NULL, NULL, NULL, 1, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (85, 8, N'Phone Number (Work Mobile)', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (86, 8, N'Phone Number (Personal Mobile)', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (87, 8, N'Phone Number (Work)', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (88, 8, N'Phone Number (Home)', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (89, 8, N'Email', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (90, 8, N'Notes', '', 12, NULL, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (91, 8, N'Address Line 1', '', 12, NULL, 50, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (92, 8, N'Address Line 2', '', 12, NULL, 50, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (93, 8, N'Address Line 3', '', 12, NULL, 50, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (94, 8, N'Address Line 4', '', 12, NULL, 50, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (95, 8, N'Address Line 5', '', 12, NULL, 50, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (96, 8, N'Postcode', '', 12, NULL, 15, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Precision], [Lookup], [ColumnID]) VALUES (98, 8, N'Contact Type', '', 12, NULL, NULL, NULL, 1, @columnID)

		-- Salary elements
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (99, 1, N'Start Date', '', 11, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (100, 1, N'End Date', '', 11, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (101, 1, N'Amount', '', 2, NULL, NULL, 0, @columnID)
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID]) VALUES (102, 1, N'Grade', '', 12, NULL, NULL, 0, @columnID)

		-- Title
		SELECT TOP 1 @columnID = columnid FROM asrsyscolumns c 
			INNER JOIN fusion.category cat ON cat.TableID = c.TableID
			WHERE cat.ID = 9 AND c.columnname <> 'ID' AND datatype = 12 ORDER BY c.columnID
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (103, 9, N'Title', '', 12, NULL, NULL, 0, @columnID)

		-- Contact Type
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (106, 10, N'Contact Type', '', 12, NULL, NULL, 0, @columnID)
		
		-- Department
		SELECT @columnID = NULL;
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (107, 11, N'Department', '', 12, NULL, NULL, 0, @columnID)
		
		-- Marital Status
		SELECT TOP 1 @columnID = columnid FROM asrsyscolumns c 
			INNER JOIN fusion.category cat ON cat.TableID = c.TableID
			WHERE cat.ID = 12 AND c.columnname <> 'ID' AND datatype = 12 ORDER BY c.columnID
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (104, 12, N'Marital Status', '', 12, NULL, NULL, 0, @columnID)

		-- Ethnicity
		SELECT TOP 1 @columnID = columnid FROM asrsyscolumns c 
			INNER JOIN fusion.category cat ON cat.TableID = c.TableID
			WHERE cat.ID = 13 AND c.columnname <> 'ID' AND datatype = 12 ORDER BY c.columnID
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (105, 13, N'Ethnicity', '', 12, NULL, NULL, 0, @columnID)
		
		-- Leaving Reason
		SELECT TOP 1 @columnID = columnid FROM asrsyscolumns c 
			INNER JOIN fusion.category cat ON cat.TableID = c.TableID
			WHERE cat.ID = 13 AND c.columnname <> 'ID' AND datatype = 12 ORDER BY c.columnID
		INSERT [fusion].[Element] ([ID], [CategoryID], [Name], [Description], [DataType], [MinSize], [MaxSize], [Lookup], [ColumnID])
			VALUES (108, 14, N'Leaving Reason', '', 12, NULL, NULL, 0, @columnID)

	END

	-- Message table
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'Message' AND type in (N'U') AND schema_id = @fusionschemaID)
	BEGIN

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

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe
				, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (1, 'StaffChange', 'Staff details', 'staff'
				, 0x
				, 1, 1, 1, 1, 1, 1, 1
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe
				, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (2, 'StaffContractChange', 'Staff contract details', 'staffContract'
				, 0x
				, 1, 1, 1, 1, 1, 1, 1
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')
	 
		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (3, 'StaffPictureChange', 'Staff picture details', 'staffPicture'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (4, 'StaffSkillChange', 'Staff Skills', 'staffSkill'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (5, 'StaffLegalDocumentChange', 'Staff documents', 'staffLegalDocument'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (6, 'StaffTimesheetPerContractChange', 'Staff timesheet details', 'staffTimesheetPerContract'
				,0x
				, 1, 0, 1, 0, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (7, 'StaffHolidayBalanceRemainingChange', 'Staff holiday balance', 'staffHolidayBalanceRemaining'
				,0x
				, 1, 1, 0, 1, 0, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (8, 'StaffContactChange', 'Staff contacts', 'staffContact'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (9, 'OutreachAreaChange', 'Outreach areas', 'outreachArea'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (10, 'PersonTitleChange', 'Person Title', 'personTitle'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (11, 'PersonMaritalStatusChange', 'Person marital status', 'personMaritalStatus'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (12, 'PersonEthnicityChange', 'Person ethnicity status', 'personEthnicity'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (13, 'PersonContactTypeChange', 'Person contact type', 'personContactType'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (14, 'StaffLeavingReasonChange', 'Leaving Reason', 'staffLeavingReason'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (15, 'StaffDepartmentChange', 'Staff Department', 'staffDepartment'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

		INSERT fusion.[Message] (ID, Name, Description, [DataNodeKey], [Schema], Version, AllowPublish, AllowSubscribe, Publish, Subscribe, StopDeletion, BypassValidation, xmlns, xmlschemaLocation, xmlxsi)
			VALUES (16, 'StaffTrainingOutcomeChange', 'Staff Training', 'staffTrainingOutcome'
				,0x
				, 1, 1, 1, 1, 1, 1, 1, 'http://advancedcomputersoftware.com/xml/fusion/socialCare'
				, 'http://advancedcomputersoftware.com/xml/fusion/socialCare https://rlo.advanced365.com/FUSION/Message%20Specifications/Data%20Examples/'
				, 'http://www.w3.org/2001/XMLSchema-instance')		

	END

	-- Message Relations table
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageRelations' AND type in (N'U') AND schema_id = @fusionschemaID)
	BEGIN

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

		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (1, 'staffRef', 0, 1)	
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (2, 'staffContractRef', 2, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (2, 'staffRef', 0, 0)		
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (3, 'staffRef', 0, 1)	
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (4, 'staffSkillRef', 4, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (4, 'staffRef', 0, 0)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (5, 'documentRef', 5, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (5, 'staffRef', 0, 0)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (6, 'staffRef', 0, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (7, 'staffRef', 0, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (8, 'staffRef', 0, 0)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (9, 'outreachAreaRef', 0, 1)	
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (10, 'personTitleRef', 9, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (11, 'personMaritalStatusRef', 12, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (12, 'personEthnicityRef', 13, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (13, 'personContactTypeRef', 10, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (14, 'staffLeavingReasonRef', 14, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey) VALUES (15, 'staffDepartmentRef', 11, 1)
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)	VALUES (16, 'staffTrainingOutcomeRef', 4, 1)	

	END

	-- MessageElements table
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageElements' AND type in (N'U') AND schema_id = @fusionschemaID)
	BEGIN	

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


		-- staffChange message
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 5, 'title', 1, 12, 1, 1, 1, 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 4, 'forenames', 1, 12, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 3, 'surname', 2, 12, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 28, 'preferredName', 3, 12, 1, 0, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 2, 'payrollNumber', 4, 12, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 7, 'DOB', 5, 11, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 1, 'employeeType', 5, 12, 0, 1, 1, 20, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 16, 'homePhoneNumber', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 26, 'workMobile', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 17, 'personalMobile', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 19, 'email', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 24, 'personalEmail', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 8, 'gender', 5, 12, 0, 1, 1, 20, 50, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 21, 'startDate', 5, 11, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 22, 'leavingDate', 5, 11, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 23, 'leavingReason', 5, 12, 1, 1, 1, 20, 50, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 1, 'companyName', 5, 12, 0, 1, 1, 20, 50, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 27, 'jobTitle', 5, 12, 0, 1, 1, 20, 50, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 31, 'managerRef', 5, 12, 1, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 10, 'addressLine1', 5, 12, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 11, 'addressLine2', 5, 12, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 12, 'addressLine3', 5, 12, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 13, 'addressLine4', 5, 12, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 14, 'addressLine5', 5, 12, 0, 1, 1, 20, 50, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (1, 15, 'postCode', 5, 12, 0, 1, 1, 20, 50, 0)

		-- staffContractChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 32 , 'name', 2, 12, 0, 1, 1, 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 33 , 'from', 3, 11, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 34, 'to', 4, 11, 0, 1, 1, 1, 0, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 35, 'department', 5, 12, 1, 1, 1, 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 36, 'primarySite', 6, 12, 0, 1, 1, 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 37, 'contractedHoursperWeek', 7, 12, 0, 1, 1, 1, 6, 0)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Position, DataType, Nillable, MinOccurs, MaxOccurs, MinSize, MaxSize, [Lookup])
			VALUES (2, 38, 'maximumHoursperWeek', 8, 12, 0, 1, 1, 1, 6, 0)

		-- staffPictureChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (3, 39, 'picture', 1, 0, 1)


		-- staffSkillChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 45, 'name', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 46, 'trainingStart', 0, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 47, 'trainingEnd', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 48, 'validFrom', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 64, 'validTo', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 49, 'reference', 1, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 50, 'outcome', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 55, 'didNotAttend', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 51, 'requestedBy', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 52, 'requestedDate', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 53, 'acceptedBy', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (4, 54, 'acceptedDate', 1, 0, 1)
		
		-- staffLegalDocumentChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (5, 65, 'name', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (5, 66, 'validFrom', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (5, 67, 'validTo', 1, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (5, 68, 'documentReference', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (5, 69, 'secondaryReference', 1, 0, 1)

		-- staffTimesheetPerContractChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 70, 'contract', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 71, 'timesheetDate', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 72, 'plannedHours', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 73, 'workedHours', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 74, 'toilHoursAccrued', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 75, 'holidayHoursTaken', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (6, 76, 'toilHoursTaken', 0, 1, 1)

		-- staffHolidayBalanceRemainingChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (7, 40, 'asOf', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (7, 41, 'holidayHoursRemaining', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (7, 42, 'holidayHoursTaken', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (7, 43, 'holidayHoursEntitlement', 0, 1, 1)


		-- StaffContactChange
		INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
			VALUES (8, 'staffContactRef', 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 80, 'title', 1, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 81, 'forenames', 1, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 82, 'surname', 1, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 83, 'description', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 84, 'relationship', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 85, 'workMobile', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 86, 'personalMobile', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 87, 'workPhoneNumber', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 88, 'homePhoneNumber', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 89, 'e-Mail', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 90, 'notes', 1, 0, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 91, 'addressLine1', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 92, 'addressLine2', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 93, 'addressLine3', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 94, 'addressLine4', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 95, 'addressLine5', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (8, 96, 'postCode', 0, 1, 1)

		-- outeachAreaChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (9, 97, 'areaName', 0, 1, 1)

		-- personTitleChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (10, 103, 'title', 0, 1, 1)

		-- staffMaritalStatusChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (11, 104, 'maritalStatus', 0, 1, 1)

		-- staffEthnicityChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (12, 105, 'ethnicity', 0, 1, 1)

		-- personContactTypeChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (13, 106, 'contactType', 0, 1, 1)

		-- staffLeavingReasonChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (14, 108, 'leavingReason', 0, 1, 1)

		-- staffDepartmentChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (15, 107, 'department', 0, 1, 1)
		
		-- staffTrainingOutcomeChange
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (16, 78, 'TrainingOutcomeName', 0, 1, 1)
		INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
			VALUES (16, 79, 'TrainingOutcomeValue', 0, 1, 1)



		
	END


	IF EXISTS (SELECT * FROM sys.objects WHERE name = 'MessageTracking' AND type in (N'U') AND schema_id = @fusionschemaID)
		DROP TABLE [fusion].[MessageTracking]

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


----------------------------------------------------------------------------
-- Connector specifics
----------------------------------------------------------------------------

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessage]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessage];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pSendMessageCheckContext]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pSendMessageCheckContext];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastProcessedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingSetLastGeneratedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastMessageDates]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageTrackingGetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogCheck]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogCheck];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pMessageLogAdd]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pMessageLogAdd];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateSetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateSetBusRef];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetLocalId]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetLocalId];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[pIdTranslateGetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[pIdTranslateGetBusRef];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pConvertData]') AND xtype = 'P')
		DROP FUNCTION [fusion].[pConvertData]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSetFusionContext]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[pSetFusionContext]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spSendFusionMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[spSendFusionMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[spGetDataForMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[spGetDataForMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[pSetDataForMessage]') AND xtype = 'P')
		DROP PROCEDURE [fusion].[pSetDataForMessage]

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[fusion].[makeXMLSafe]') AND xtype = 'FN')
		DROP FUNCTION [fusion].[makeXMLSafe]

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pSetFusionContext
	--
	-- Purpose: Sets current connection context to indicate message processing 
	--          underway
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pSetFusionContext]
		(
			@MessageType varchar(50)
		)
	AS
	BEGIN
		SET NOCOUNT ON;
		
		DECLARE @ContextInfo varbinary(128)
	 
		SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );
	 
		SET CONTEXT_INFO @ContextInfo
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pIdTranslateGetBusRef
	--
	-- Purpose: Converts a local identifier into a uniqueidentifier for the bus, 
	--			returning consistent value for all future conversions.  
	--          This will create a new identifier where one is not found where
	--			@CanGenerate = 1
	--
	-- Returns: 0 = success, 1 = failure
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateGetBusRef]
		(
			@TranslationName varchar(50),
			@LocalId varchar(25),
			@BusRef uniqueidentifier output,
			@DidGenerate bit = 0 output,
			@CanGenerate bit = 1
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		SET @BusRef = NULL;
		SET @DidGenerate = 0;
		
		SELECT @BusRef = BusRef from [fusion].IdTranslation
			WHERE TranslationName = @TranslationName AND LocalId = @LocalId;
		
		IF @@ROWCOUNT = 0
		BEGIN
			IF @CanGenerate = 1
			BEGIN
				SET @BusRef = NEWID();

							
				INSERT fusion.IdTranslation (TranslationName, LocalId, BusRef) 
						VALUES (@TranslationName, @LocalId, @BusRef);
				
				SET @DidGenerate = 1;
						
				RETURN 0;
			END
			RETURN 1;
		END

		RETURN 0;
	END';

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pIdTranslateGetLocalId
	--
	-- Purpose: Finds the local id equivelant for the given Bus reference number, 
	--          assuming it has previous been created through spIdTranslateSetBusRef
	--
	-- Returns: 
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateGetLocalId]
		(
			@TranslationName varchar(50),
			@BusRef uniqueidentifier,
			@LocalId varchar(25) output
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SET @LocalId = null;
	
		SELECT @LocalId = LocalId from [fusion].IdTranslation 
			WHERE TranslationName = @TranslationName and BusRef = @BusRef;
	END';

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spIdTranslateSetBusRef
	--
	-- Purpose: Sets the conversion of a given local reference into the given bus ref
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pIdTranslateSetBusRef]
		(
			@TranslationName varchar(50),
			@LocalId varchar(25),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		--BEGIN TRAN;
	
		DELETE fusion.IdTranslation
			WHERE TranslationName = @TranslationName and LocalId = @LocalId;
		
		INSERT fusion.IdTranslation(TranslationName, LocalId, BusRef) 
			VALUES (@TranslationName, @LocalId, @BusRef);

		--COMMIT TRAN;
	END	'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageLogAdd
	--
	-- Purpose: Adds fact that message has been processed to local message log
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageLogAdd]
		(
			@MessageType varchar(50),
			@MessageRef uniqueidentifier,
			@Originator varchar(50) = NULL
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		INSERT fusion.MessageLog (MessageType, MessageRef, Originator, ReceivedDate) VALUES (@MessageType, @MessageRef, @Originator, GETUTCDATE());

	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageLogCheck
	--
	-- Purpose: Checks whether message has been processed before
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageLogCheck]
		(
			@MessageType varchar(50),
			@MessageRef uniqueidentifier,
			@ReceivedBefore bit output
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		IF EXISTS ( SELECT * FROM fusion.MessageLog WHERE MessageType = @MessageType AND MessageRef = @MessageRef )
		BEGIN
			SET @ReceivedBefore = 1
		END
		ELSE
		BEGIN
			SET @ReceivedBefore = 0
		END
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingGetLastGeneratedXml
	--
	-- Purpose: Gets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingGetLastGeneratedXml]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SELECT LastGeneratedXml
			FROM fusion.MessageTracking
			WHERE MessageType = @MessageType AND BusRef = @BusRef;

	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingGetLastMessageDates
	--
	-- Purpose: Gets the last processing date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingGetLastMessageDates]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SELECT LastProcessedDate, LastGeneratedDate
			FROM fusion.MessageTracking
			WHERE MessageType = @MessageType AND BusRef = @BusRef;

	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastGeneratedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedDate]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastGeneratedDate datetime
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM [fusion].MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE [fusion].MessageTracking
			   SET LastGeneratedDate = @LastGeneratedDate
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT [fusion].MessageTracking (MessageType, BusRef, LastGeneratedDate)
				VALUES (@MessageType, @BusRef, @LastGeneratedDate)
		END		
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastGeneratedXml
	--
	-- Purpose: Sets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastGeneratedXml]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastGeneratedXml varchar(max)
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM fusion.MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE fusion.MessageTracking
			   SET LastGeneratedXml = @LastGeneratedXml
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT fusion.MessageTracking (MessageType, BusRef, LastGeneratedXml)
				VALUES (@MessageType, @BusRef, @LastGeneratedXml)
		END		
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pMessageTrackingSetLastProcessedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pMessageTrackingSetLastProcessedDate]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastProcessedDate datetime
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM fusion.MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE fusion.MessageTracking
			   SET LastProcessedDate = @LastProcessedDate
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT fusion.MessageTracking (MessageType, BusRef, LastProcessedDate)
				VALUES (@MessageType, @BusRef, @LastProcessedDate)
		END		
	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pSendMessage
	--
	-- Purpose: Triggers a message to be sent
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pSendMessage]
		(
			@MessageType varchar(50),
			@LocalId int
		)
	AS
	BEGIN
		SET NOCOUNT ON;
	
		DECLARE @DialogHandle uniqueidentifier;
		SET @DialogHandle = NEWID();

		BEGIN DIALOG @DialogHandle 
			FROM SERVICE FusionApplicationService 
			TO SERVICE ''FusionConnectorService''
			ON CONTRACT TriggerFusionContract
			WITH ENCRYPTION = OFF;
		
		DECLARE @msg varchar(max);

		SET @msg = (SELECT	@MessageType AS MessageType, 
							@LocalId as LocalId,
							CONVERT(varchar(50), GETUTCDATE(), 126)+''Z'' as TriggerDate 
						FOR XML PATH(''SendFusionMessage''));	
		
		SEND ON CONVERSATION @DialogHandle
			MESSAGE TYPE TriggerFusionSend (@msg);
	 
		END CONVERSATION @DialogHandle;

	END'

	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    pSendMessageCheckContext
	--
	-- Purpose: Triggers a message to be sent, checking context
	--          to see if we are in the process of updating according to
	--          this same message being received (preventing multi-master
	--          re-publish scenario)
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[pSendMessageCheckContext]
		(
			@MessageType varchar(50),
			@LocalId int
		)
	AS
	BEGIN
		SET NOCOUNT ON;

		DECLARE @ContextInfo varbinary(128);
 
		SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );
 
		IF CONTEXT_INFO() IS NULL OR CONTEXT_INFO() <> @ContextInfo
		BEGIN	
			EXEC fusion.pSendMessage @MessageType, @LocalId;
		END
	END'

GO


CREATE PROCEDURE fusion.spSendFusionMessage(@TableID integer, @RecordID integer)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @messageName varchar(255);

	DECLARE MessageCursor CURSOR LOCAL FAST_FORWARD FOR SELECT DISTINCT m.Name
		FROM ASRSysColumns c
		INNER JOIN fusion.element e ON e.ColumnID = c.ColumnID
		INNER JOIN fusion.messageelements me ON me.ElementID = e.ID		
		INNER JOIN fusion.message m ON me.MessageID = m.ID
		WHERE c.TableID = @tableID;
	
	OPEN MessageCursor;
	FETCH NEXT FROM MessageCursor INTO @messageName;
	WHILE @@FETCH_STATUS = 0 
	BEGIN 
		EXEC fusion.[pSendMessageCheckContext] @MessageType=@messageName, @LocalId=@RecordID
	    FETCH NEXT FROM MessageCursor INTO @messageName;
	END
	CLOSE MessageCursor;
	DEALLOCATE MessageCursor;

END
GO

CREATE FUNCTION fusion.makeXMLSafe(@input varchar(MAX))
	RETURNS VARCHAR(MAX)
	BEGIN
	RETURN 
		Replace(Replace(Replace(Replace(Replace(@input,'&','&amp;'),'<', '&lt;'),'>', '&gt;'),'"', '&quot;'), '''', '&#39;')
	END

GO

CREATE PROCEDURE [fusion].[pSetDataForMessage](@messagetype varchar(255), @id integer OUTPUT, @xml varchar(MAX), @parentguid varchar(255))
AS
BEGIN

	SET NOCOUNT ON;

DECLARE @xmlCode xml;

DECLARE @ParmDefinition nvarchar(500);
DECLARE @ssql nvarchar(MAX),
		@sInsert nvarchar(MAX),
		@sUpdate nvarchar(MAX),
		@sColumns nvarchar(MAX),
		@sTableName nvarchar(255),
		@messagename nvarchar(255),
		@datanodeKey nvarchar(255),
		@foreignKeyName nvarchar(255),
		@foreignkeyvalue nvarchar(255),
		@executeCode nvarchar(MAX);

SET @messagename = @messagetype

SET @executeCode = '';
SET @ssql = '0 AS ID';
SET @sInsert = '0 AS ID';
SET	@sUpdate = '';
SET @xmlCode = convert(xml, @xml);

SELECT @datanodeKey = DataNodeKey FROM fusion.message WHERE Name = @messagetype;

SELECT @foreignKeyName = 'ID_' + convert(varchar(4), c.TableID) FROM fusion.message m
	INNER JOIN fusion.[MessageRelations] mr ON mr.messageID = m.ID
	INNER JOIN fusion.[category] c ON c.ID = mr.categoryID
	WHERE mr.IsPrimaryKey = 0 AND m.name = @messagetype;

IF LEN(@foreignKeyName) > 0
BEGIN
	SELECT @foreignkeyvalue = LocalID FROM fusion.idtranslation WHERE busRef = @parentguid
	SET @foreignkeyvalue = ISNULL(@foreignkeyvalue,0)
END


-- Temp table
SET @ssql = 'DECLARE @mytable TABLE (ID integer'
SELECT @ssql = @ssql + ', [' + nodekey + '] nvarchar(MAX)'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @ssql = @ssql + ');'
SET @executeCode = @executeCode + @ssql + CHAR(13);

-- Insert
SET @sInsert = 'INSERT @mytable (ID '
SELECT @sInsert = @sInsert + ', [' + nodekey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename
SET @sInsert = @sInsert + ')'

SET @ssql = '';
SELECT @ssql = @ssql + ',c.value(''nsWithXNS:' + nodekey + '[1]'', ''nvarchar(MAX)'') AS [' + nodekey + ']' + CHAR(13) 
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SET @ssql = 'WITH XMLNAMESPACES (''http://advancedcomputersoftware.com/xml/fusion/socialCare'' AS nsWithXNS)' + CHAR(13) +
	@sInsert +
	'SELECT 0' + @ssql + 'FROM @xmlCode.nodes(''nsWithXNS:' + @datanodeKey + ''') AS mytable(c)'

SET @executeCode = @executeCode + @ssql + CHAR(13);

SELECT TOP 1 @sTableName = t.tablename
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename

SELECT @sInsert = CASE WHEN LEN(@foreignKeyName) > 0 THEN @foreignKeyName ELSE '' END
SELECT @sInsert = @sInsert + CASE WHEN LEN(@sInsert) > 0 THEN ', ' ELSE '' END + ' [' + c.columnname + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sInsert= 'INSERT ' + @sTableName + ' ( ' + @sInsert + ') SELECT ';

SELECT @sColumns = CASE WHEN LEN(@foreignKeyName) > 0 THEN @foreignkeyvalue ELSE '' END
SELECT @sColumns = @sColumns + CASE WHEN LEN(@sColumns) > 0 THEN ', ' ELSE '' END + '[' + e.NodeKey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sInsert = @sInsert + @sColumns + ' FROM @mytable';

SELECT @sUpdate = @sUpdate + CASE WHEN LEN(@sUpdate) > 0 THEN ', ' ELSE '' END
	+ @sTableName + '.[' + c.columnname + '] = message.[' + e.nodekey + ']'
	FROM fusion.messageElements e
	INNER JOIN fusion.message m ON m.id = e.messageid
	INNER JOIN fusion.element lm ON e.ElementID = lm.ID
	INNER JOIN asrsyscolumns c ON c.columnid = lm.columnid
	INNER JOIN asrsystables t ON c.tableID = t.tableID
	WHERE lm.columnID IS NOT NULL AND m.name = @messagename;
SET @sUpdate = 'UPDATE ' + @sTableName + ' SET ' + @sUpdate + ' FROM @mytable message WHERE ' + @sTableName + '.ID = @ID;'

SET @executeCode = @executeCode 
	+ 'IF (@ID > 0)' + CHAR(13)
	+ @sUpdate + CHAR(13) 
	+ ' ELSE ' + CHAR(13)
	+ ' BEGIN ' + CHAR(13)
	+ @sInsert  + CHAR(13)
	+ ' SELECT @ID = MAX(ID) FROM ' + @sTableName
	+ ' END' + CHAR(13)
	+ ' SELECT @ID;';


SET @ParmDefinition = N'@xmlCode xml, @ID integer OUTPUT';

IF LEN(@executeCode) > 0
	EXEC sp_executeSQL @executeCode, @ParmDefinition, @xmlcode = @xmlcode, @id = @id
ELSE
	SELECT 0

END
GO

CREATE PROCEDURE [fusion].[spGetDataForMessage](@messagetype varchar(255), @ID integer, @ID_Parent1 integer, @ID_Parent2 integer, @ID_Parent3 integer)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @ssql nvarchar(MAX),
			@linesepcode nvarchar(255);

	DECLARE @xmlMessageBody	varchar(MAX),
			@xmllastmessage		varchar(MAX),
			@xmlMessageCode		varchar(MAX),
			@fusiontypeID		integer,
			@effectivefrom		datetime,
			@postguid			varchar(255),
			@tablename			varchar(255),
			@selectcode			varchar(MAX);

	DECLARE @messageID		int,
			@xmlns			varchar(255),
			@schemaLocation	varchar(255),
			@xmlxsi			varchar(255),
			@dataNodeKey	varchar(255),		
			@primaryKey		varchar(255),
			@foreignKey		varchar(MAX),
			@parentKey		varchar(255),
			@parentID		varchar(10),
			@version		int;

	SET @effectivefrom = '';
	SET @ssql = '';
	
	-- Details for this message
	SELECT @messageID =	ID
			, @xmlns = [xmlns]
			, @schemaLocation = [xmlschemalocation]
			, @version = [version]
			, @xmlxsi = [xmlxsi]
			, @dataNodeKey = [DataNodeKey]
		FROM fusion.[message] WHERE name = @messagetype

	SET @linesepcode = ' + CHAR(13)+CHAR(10) + ''				'' + ';

	-- Get the creation date for this record
	--SELECT @effectivefrom = ISNULL(effectivefrom, GETDATE())
	--	FROM fusion.IdTranslation
	--	WHERE translationname = @messagetype AND localid = @ID;
	SELECT @effectivefrom = GETDATE();
	SET @ssql = '';

	---- Last XML message
	--SELECT TOP 1 @xmllastmessage = ISNULL(mt.LastGeneratedXml,'') 
	--	FROM fusion.messagetracking mt
	--		INNER JOIN fusion.IdTranslation tr ON tr.LocalId = @ID AND tr.BusRef = mt.BusRef
	--	WHERE tr.TranslationName = @messagetype
	--	ORDER BY mt.LastProcessedDate DESC;

	SET @xmllastmessage = '';

	-- Get table name
	SELECT @tablename = t.tablename, @primaryKey = mr.NodeKey FROM fusion.MessageRelations mr
		INNER JOIN fusion.Category c ON c.ID = mr.CategoryID
		INNER JOIN ASRSysTables t ON t.TableID = c.TableID
		WHERE messageID = @messageID AND mr.IsPrimaryKey = 1

	-- Get relationship data
	SET @foreignKey = ''
	SELECT @foreignKey = NodeKey
		FROM fusion.MessageRelations mr
		INNER JOIN fusion.message m ON mr.MessageID = m.ID
		WHERE mr.IsPrimaryKey = 0 AND m.Name = @messagetype;

	SET @parentKey = '';
	SELECT @parentKey = 'ID_' + convert(varchar(10), c.TableID)
		FROM fusion.MessageRelations mr
		INNER JOIN fusion.message m ON mr.MessageID = m.ID
		INNER JOIN fusion.category c ON c.ID = mr.CategoryID
		WHERE mr.IsPrimaryKey = 0 AND m.Name = @messagetype;


	-- Build message body
	SET @ssql = '';
	SELECT @ssql = @ssql + CASE LEN(@ssql) WHEN 0 THEN '' ELSE ' + ' END +
		CASE 
			WHEN NULLIF(x.value, '') IS NOT NULL
				THEN @linesepcode + '''<' + x.xmlnodekey + '>' + x.value + '</' + x.xmlnodekey + '>''' 

			WHEN c.datatype = 2
				THEN 'CASE ISNULL([' + c.ColumnName + '],0) WHEN 0 THEN '''' ELSE '
					+ @linesepcode + '''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

			WHEN x.minoccurs = 0 AND x.nilable = 0 AND c.datatype = 11
				THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
					+ @linesepcode + '''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

			WHEN x.minoccurs = 0 AND x.nilable = 0
				THEN 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN '''' ELSE '
					+ @linesepcode + '''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 

			WHEN x.nilable = 0 AND c.datatype = 11
				THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' +  x.xmlnodekey + '/>'' ELSE ''<' + x.xmlnodekey 
					+ '>'' + convert(varchar(10),[' + c.ColumnName + '], 120) + ''</' + x.xmlnodekey + '>'' END' 

			WHEN x.nilable = 0
				THEN + @linesepcode + '''<' + x.xmlnodekey + '>'' + ISNULL(fusion.makeXMLSafe([' + c.ColumnName + ']),'''') + ''</' + x.xmlnodekey + '>''' 

			WHEN x.nilable = 1 AND c.datatype = 11
				THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
					+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + convert(varchar(10),[' + c.ColumnName + '],120) + ''</' + x.xmlnodekey + '>'' END' 

			WHEN x.nilable = 1
				THEN + @linesepcode + 'CASE ISNULL([' + c.ColumnName + '],'''') WHEN '''' THEN ''<' + x.xmlnodekey 
					+ ' xsi:nil="true"/>'' ELSE ''<' + x.xmlnodekey + '>'' + fusion.makeXMLSafe([' + c.ColumnName + ']) + ''</' + x.xmlnodekey + '>'' END' 
	 		
			ELSE '''UNKNOWN FIELD TYPE'''  END  
			-- + ' AS [column_' + convert(varchar(3), x.position) + ']'

		FROM [fusion].[MessageDefinition] x
			INNER JOIN ASRSysColumns c ON c.columnID = x.columnid
			INNER JOIN ASRSysTables t ON t.TableID = x.tableid
			WHERE xmlMessageID = @messagetype;


	IF LEN(@ssql) > 0 
	BEGIN
	
		SELECT @ssql = N'SELECT ' + CASE WHEN LEN(@parentKey) > 0 THEN ' @parentid = ' + @parentKey + ', ' ELSE '' END
			+ '@xml = ' + @ssql + ' FROM [' + @tablename + ']  WHERE ID = ' + convert(varchar(10),@ID)

		EXECUTE sp_executeSQL @ssql, N'@xml nvarchar(MAX) OUTPUT, @parentID int OUTPUT'
			, @xml = @xmlMessageBody OUTPUT, @parentID = @parentID OUTPUT;

		SELECT N'<?xml version="1.0" encoding="utf-8"?>
		<' + @messagetype + ' version="' + convert(varchar(2),@version) + '" ' + @primarykey + '="{0}" '
			+ CASE WHEN LEN(@foreignKey) > 0 THEN @foreignKey + '="{1}"' ELSE '' END +
			' xsi:schemaLocation="' + @schemaLocation + @messagetype + '.xsd"
			xmlns="' + @xmlns + '"
			xmlns:xsi="' + @xmlxsi + '">
			<data auditUserName="' + CURRENT_USER + '" recordStatus="Active" effectiveFrom="' + convert(varchar(10),@effectivefrom, 120) + '">
				<' + @dataNodeKey + '>'
				+ @xmlMessageBody +
				'</' + @dataNodeKey + '>
			</data>
			</' + @messagetype + '>' AS XMLCode
			, @parentID AS ParentID

	END
	ELSE
	BEGIN
		SELECT 'whoops - you ain''t configured this thing properly. Contact Harpenden QA on...'
	
	END

END




GO

