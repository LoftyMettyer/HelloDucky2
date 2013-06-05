--select * from asrsysAccordtransferfielddefinitions where transfertypeid = 0 order by description 
-- select * from ASRSysAccordTransferTypes  order by transfertype
-- select * FROM asrsysModuleSetup

--select * from [fusion].[Element] 

DECLARE @columnID integer,
		@tableID integer;

DELETE FROM fusion.[MessageRelations]
DELETE FROM fusion.[MessageElements]
DELETE FROM fusion.[Message]

DELETE FROM [fusion].[Element]
DELETE FROM [fusion].[Category]

DBCC CHECKIDENT ('fusion.[MessageElements]', RESEED, 0)





	-- Categories
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




	-- Messages (based on xsd definitions)
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

	-- staffChange message
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (1, 'staffRef', 0, 1)	
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (2, 'staffContractRef', 2, 1)
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (2, 'staffRef', 0, 0)
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (1, 'staffRef', 0, 1)	
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (3, 39, 'picture', 1, 0, 1)


	-- staffSkillChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (4, 'staffSkillRef', 4, 1)
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (4, 'staffRef', 0, 0)
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (5, 'documentRef', 5, 1)
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (5, 'staffRef', 0, 0)
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (6, 'staffRef', 0, 1)
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (7, 'staffRef', 0, 1)
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
		VALUES (8, 'staffRef', 0, 0)
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
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (9, 'outreachAreaRef', 0, 1)	
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (9, 97, 'areaName', 0, 1, 1)

	-- personTitleChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (10, 'personTitleRef', 9, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (10, 103, 'title', 0, 1, 1)

	-- staffMaritalStatusChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (11, 'personMaritalStatusRef', 12, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (11, 104, 'maritalStatus', 0, 1, 1)

	-- staffEthnicityChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (12, 'personEthnicityRef', 13, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (12, 105, 'ethnicity', 0, 1, 1)

	-- personContactTypeChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (13, 'personContactTypeRef', 10, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (13, 106, 'contactType', 0, 1, 1)

	-- staffLeavingReasonChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (14, 'staffLeavingReasonRef', 14, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (14, 108, 'leavingReason', 0, 1, 1)

	-- staffDepartmentChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (15, 'staffDepartmentRef', 11, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (15, 107, 'department', 0, 1, 1)
	
	-- staffTrainingOutcomeChange
	INSERT fusion.[MessageRelations] (MessageID, NodeKey, CategoryID, IsPrimaryKey)
		VALUES (16, 'staffTrainingOutcomeRef', 4, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (16, 78, 'TrainingOutcomeName', 0, 1, 1)
	INSERT fusion.[MessageElements] (MessageID, ElementID, NodeKey, Nillable, MinOccurs, MaxOccurs)
		VALUES (16, 79, 'TrainingOutcomeValue', 0, 1, 1)

